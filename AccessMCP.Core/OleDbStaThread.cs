using System.Collections.Concurrent;
using System.Diagnostics;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>
/// Runs all OLE DB / ACE / Jet work on a single dedicated STA thread.
/// <para>
/// The ACE/Jet OLE DB provider uses apartment-threaded COM objects. Plain <c>async</c>/<c>await</c> code
/// resumes continuations on arbitrary thread-pool (MTA) threads, so creating and using OleDb connections
/// from there forces cross-apartment COM marshaling that can corrupt internal native state, leading to
/// <see cref="System.ExecutionEngineException"/> or <c>E_FAIL(0x80004005)</c>.
/// </para>
/// <para>
/// By funnelling every OleDb call through a single persistent STA thread, all COM objects live in the same
/// apartment and no marshaling is needed. This pattern was proven against the same ACE/Jet provider in
/// another codebase; plain async code only appears safe while a single caller runs one tool call at a time.
/// AccessOleDb's own connection-per-call design still applies serialization at the STA-thread level (only
/// one call proceeds at a time, queued).
/// </para>
/// <para>
/// Queued work runs under the <b>caller's</b> <see cref="ExecutionContext"/>. AsyncLocal state does not
/// flow to a raw <see cref="Thread"/> that invokes a delegate directly, so without the capture/restore
/// below anything ambient (transaction/connection scopes, logging correlation) silently reads as its
/// default inside a work item. Where this pattern was first used, that manifested as a nested helper missing the
/// in-flight transaction and opening a second OleDb connection to the same file, corrupting ACE/Jet.
/// </para>
/// </summary>
internal static class OleDbStaThread
{
    private static readonly ConcurrentQueue<Action> _queue = new();
    private static readonly AutoResetEvent _workAvailable = new(false);
    private static readonly Thread _thread;

    static OleDbStaThread()
    {
        _thread = new Thread(ProcessQueue)
        {
            IsBackground = true,
            Name = "AccessMCP-OleDb-STA"
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    private static void ProcessQueue()
    {
        // WaitHandle.WaitOne pumps COM messages on STA threads via CoWaitForMultipleHandles, unlike
        // SemaphoreSlim/Monitor which do not.
        while (true)
        {
            _workAvailable.WaitOne();
            while (_queue.TryDequeue(out var action))
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    // Work items funnel their own failures into their TaskCompletionSource, so reaching
                    // here means the dispatch wrapper faulted. Swallow it: if this thread unwinds, every
                    // later Execute call blocks forever on a TCS nothing will complete.
                    Debug.Fail($"OleDb STA dispatch faulted: {ex}");
                }
            }
        }
    }

    /// <summary>
    /// Executes <paramref name="work"/> on the dedicated STA thread and returns the result.
    /// Blocks the caller until the work completes.
    /// If already on the STA thread (e.g. nested call), executes inline.
    /// </summary>
    public static T Execute<T>(Func<T> work)
    {
        if (Thread.CurrentThread == _thread)
            return work();

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

        // Capture here, on the caller's thread — see the class remarks. A null capture means the caller
        // suppressed flow, in which case running bare is the correct behaviour.
        var context = ExecutionContext.Capture();

        _queue.Enqueue(() =>
        {
            try
            {
                if (context is null)
                {
                    tcs.TrySetResult(work());
                }
                else
                {
                    var result = default(T)!;
                    ExecutionContext.Run(context, _ => result = work(), null);
                    tcs.TrySetResult(result);
                }
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }
        });

        _workAvailable.Set();

        // Block the caller synchronously; the caller is expected to be a background/pool thread (this is
        // always invoked from inside a Task.Run in AccessOleDb) — the actual OleDb work runs on the STA
        // thread while this thread waits.
        return tcs.Task.GetAwaiter().GetResult();
    }
}
