using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

/// <summary>
/// ACE/Jet is apartment-threaded COM, so all OLE DB work must run on the one STA thread. These pin the
/// funnel's contract: same thread every time, STA, nested calls inline, failures and context flow back.
/// </summary>
public sealed class OleDbStaThreadTests
{
    [Fact]
    public async Task WorkRunsOnOneDedicatedStaThread()
    {
        var first = await Task.Run(() => OleDbStaThread.Execute(() => (Thread.CurrentThread.ManagedThreadId, Thread.CurrentThread.GetApartmentState(), Thread.CurrentThread.Name)));
        var second = await Task.Run(() => OleDbStaThread.Execute(() => Thread.CurrentThread.ManagedThreadId));

        Assert.Equal(ApartmentState.STA, first.Item2);
        Assert.Equal("AccessMCP-OleDb-STA", first.Name);
        Assert.Equal(first.ManagedThreadId, second);
        Assert.NotEqual(Environment.CurrentManagedThreadId, second);
    }

    [Fact]
    public async Task ConcurrentCallersAreSerialized()
    {
        var active = 0;
        var maxActive = 0;

        var callers = Enumerable.Range(0, 8).Select(i => Task.Run(() => OleDbStaThread.Execute(() =>
        {
            var now = Interlocked.Increment(ref active);
            InterlockedMax(ref maxActive, now);
            Thread.Sleep(5);
            Interlocked.Decrement(ref active);
            return i;
        })));

        var results = await Task.WhenAll(callers);

        Assert.Equal(1, maxActive);
        Assert.Equal(Enumerable.Range(0, 8), results.Order());
    }

    [Fact]
    public async Task NestedCallRunsInlineInsteadOfDeadlocking()
    {
        var execute = Task.Run(() => OleDbStaThread.Execute(() =>
        {
            var outer = Environment.CurrentManagedThreadId;
            var inner = OleDbStaThread.Execute(() => Environment.CurrentManagedThreadId);
            return (outer, inner);
        }));

        var completed = await Task.WhenAny(execute, Task.Delay(TimeSpan.FromSeconds(10)));

        Assert.Same(execute, completed);
        var (outer, inner) = await execute;
        Assert.Equal(outer, inner);
    }

    [Fact]
    public async Task ExceptionsPropagateAndTheThreadSurvives()
    {
        await Assert.ThrowsAsync<InvalidOperationException>(
            () => Task.Run(() => OleDbStaThread.Execute<int>(() => throw new InvalidOperationException("boom"))));

        Assert.Equal(42, await Task.Run(() => OleDbStaThread.Execute(() => 42)));
    }

    [Fact]
    public async Task AsyncLocalValuesFlowIntoTheWorkItem()
    {
        var local = new AsyncLocal<string>();

        var seen = await Task.Run(() =>
        {
            local.Value = "caller";
            return OleDbStaThread.Execute(() => local.Value);
        });

        Assert.Equal("caller", seen);
    }

    private static void InterlockedMax(ref int target, int value)
    {
        int current;
        while ((current = Volatile.Read(ref target)) < value &&
               Interlocked.CompareExchange(ref target, value, current) != current)
        {
        }
    }
}
