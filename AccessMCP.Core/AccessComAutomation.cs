using System.Diagnostics;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>
/// Runs Access COM automation in a dedicated STA thread and provides helpers for common Access constants.
/// </summary>
internal sealed class AccessComAutomation(AccessMcpOptions options)
{
    private const int AcViewNormal = 0;
    private const int AcViewDesign = 1;
    private const int AcViewPreview = 2;
    private const int AcViewReport = 5;
    private const int AcViewLayout = 6;
    private const int AcSaveYes = 1;
    private const int AcSaveNo = 2;
    private const int AcFormObjectType = 2;
    private const int AcReportObjectType = 3;
    private const int AcCmdSaveRecord = 97;
    private const int AcCmdCompileAllModules = 125;
    private const int AcCmdCompileAndSaveAllModules = 126;
    private const int AcCmdSaveAllModules = 280;

    // DAO database property used by Access to decide which form (if any) to open on start-up.
    private const string StartUpFormPropertyName = "StartUpForm";
    private const int DbText = 10;

    private static readonly IReadOnlyDictionary<string, int> ControlTypes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
    {
        ["label"] = 100,
        ["line"] = 102,
        ["image"] = 103,
        ["commandbutton"] = 104,
        ["optionbutton"] = 105,
        ["checkbox"] = 106,
        ["optiongroup"] = 107,
        ["textbox"] = 109,
        ["listbox"] = 110,
        ["combobox"] = 111,
        ["subform"] = 112,
        ["tabcontrol"] = 123,
        ["togglebutton"] = 122,
        ["rectangle"] = 119
    };

    private static readonly IReadOnlyDictionary<string, int> Sections = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
    {
        ["detail"] = 0,
        ["header"] = 1,
        ["footer"] = 2,
        ["pageheader"] = 3,
        ["pagefooter"] = 4,
        ["grouplevel1header"] = 5,
        ["grouplevel1footer"] = 6,
        ["grouplevel2header"] = 7,
        ["grouplevel2footer"] = 8
    };

    private readonly AccessMcpOptions _options = options;

    // Access COM automation is effectively single-instance, and start-up-form suppression opens the
    // database exclusively and edits one of its properties. Serialize sessions so two overlapping
    // calls can't skip suppression (because the other holds the file) or restore the property while
    // another Access instance still has the database open.
    private static readonly SemaphoreSlim AutomationGate = new(1, 1);

    /// <summary>
    /// Opens the target Access database in a background STA thread and executes the provided action.
    /// </summary>
    public async Task<T> RunInAccessAsync<T>(string databasePath, Func<dynamic, T> action, CancellationToken cancellationToken)
    {
        if (!OperatingSystem.IsWindows())
        {
            throw new PlatformNotSupportedException("Access COM automation is only supported on Windows.");
        }

        if (!_options.EnableComAutomation)
        {
            throw new InvalidOperationException("Access COM automation is disabled by configuration.");
        }

        cancellationToken.ThrowIfCancellationRequested();

        // Held for the whole session (released in the worker's outer finally) so only one Access
        // instance touches the database at a time; see the note on AutomationGate. Waited on
        // asynchronously so a UI-thread caller isn't blocked while another session runs.
        await AutomationGate.WaitAsync(cancellationToken).ConfigureAwait(false);

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            dynamic? app = null;
            int? processId = null;
            try
            {
                cancellationToken.ThrowIfCancellationRequested();

                var type = Type.GetTypeFromProgID(_options.AccessComProgId)
                    ?? throw new InvalidOperationException("Microsoft Access is not installed.");

                app = Activator.CreateInstance(type)
                    ?? throw new InvalidOperationException("Failed to create Access COM instance.");

                // Captured up front so the finally block can force this specific instance closed
                // even if Quit() doesn't fully release it (see EnsureProcessExited).
                processId = TryGetProcessId(app);

                app.Visible = false;
                app.UserControl = false;

                if (!string.IsNullOrWhiteSpace(_options.WorkgroupPath))
                {
                    app.DBEngine.SystemDB = _options.WorkgroupPath;
                }

                // A configured StartUpForm runs as soon as Access opens the database and can sit on a
                // modal dialog until a human dismisses it, which would hang automation. Remove the
                // property via DAO, open, then restore it through the live database so the file is left
                // exactly as it was and normal interactive start-up behaviour is unchanged.
                string? startUpFormValue = null;
                var startUpFormRemoved = _options.SuppressStartupForm
                    && TryRemoveStartUpForm((object)app, databasePath, out startUpFormValue);
                try
                {
                    app.OpenCurrentDatabase(databasePath, false);
                }
                finally
                {
                    if (startUpFormRemoved)
                    {
                        RestoreStartUpForm((object)app, databasePath, startUpFormValue);
                    }
                }

                var result = action(app);
                tcs.TrySetResult(result);
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }
            finally
            {
                Timer? cleanupTimer = null;
                try
                {
                    if (app is not null)
                    {
                        // Quit can block on a VBE "Save As" prompt. Bound the cleanup window so the
                        // automation gate is always released; the process is a throwaway instance.
                        if (processId is int cleanupPid)
                        {
                            cleanupTimer = new Timer(
                                _ =>
                                {
                                    try
                                    {
                                        using var process = Process.GetProcessById(cleanupPid);
                                        process.Kill(entireProcessTree: true);
                                    }
                                    catch
                                    {
                                    }
                                },
                                null,
                                TimeSpan.FromSeconds(Math.Max(1, _options.VbaCleanupTimeoutSeconds)),
                                Timeout.InfiniteTimeSpan);
                        }

                        try
                        {
                            app.CloseCurrentDatabase();
                            app.Quit();
                        }
                        finally
                        {
                            try { cleanupTimer?.Dispose(); } catch { }
                            cleanupTimer = null;
                        }
                    }
                }
                catch
                {
                }
                finally
                {
                    ReleaseComObject(app);
                    app = null;

                    // The dynamic/DLR COM interop used throughout this class can leave outstanding
                    // RCWs (e.g. from transient property-chain access) that keep the out-of-process
                    // MSACCESS.EXE alive even after Quit() returns. This runs after tcs.TrySetResult
                    // above, so it never adds latency to the caller - it only affects how long this
                    // background thread lingers. Force a GC pass to release anything reclaimable,
                    // then fall back to killing this specific process by PID if it's still running,
                    // so repeated automation calls never accumulate orphaned hidden Access instances.
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    if (processId is int pid)
                    {
                        EnsureProcessExited(pid);
                    }

                    AutomationGate.Release();
                }
            }
        })
        {
            IsBackground = true
        };

        thread.SetApartmentState(ApartmentState.STA);
        StartHoldingGate(thread);

        using var registration = cancellationToken.Register(() => tcs.TrySetCanceled(cancellationToken));
        return await tcs.Task.ConfigureAwait(false);
    }

    /// <summary>Starts a session worker, releasing the gate if the thread never runs to release it.</summary>
    private static void StartHoldingGate(Thread thread)
    {
        try
        {
            thread.Start();
        }
        catch
        {
            AutomationGate.Release();
            throw;
        }
    }

    /// <summary>
    /// Runs an Access action with a caller-side watchdog and a structured outcome. Unlike
    /// <see cref="RunInAccessAsync{T}"/>, a hang returns <c>TimedOut</c> instead of never completing:
    /// the caller waits on the action and, on expiry, kills the Access process by id. Cleanup in the
    /// worker is bounded too, so a modal dialog during Quit cannot hold the automation gate forever.
    /// </summary>
    public async Task<AccessRunOutcome<T>> RunInAccessWithWatchdogAsync<T>(
        string? databasePath,
        Func<dynamic, T> action,
        AccessRunOptions? runOptions,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(action);

        if (!OperatingSystem.IsWindows())
        {
            throw new PlatformNotSupportedException("Access COM automation is only supported on Windows.");
        }

        if (!_options.EnableComAutomation)
        {
            throw new InvalidOperationException("Access COM automation is disabled by configuration.");
        }

        var options = runOptions ?? AccessRunOptions.Default;

        if (options.OpenDatabase && string.IsNullOrWhiteSpace(databasePath))
        {
            throw new ArgumentException("Database path is required when OpenDatabase is true.", nameof(databasePath));
        }

        cancellationToken.ThrowIfCancellationRequested();
        await AutomationGate.WaitAsync(cancellationToken).ConfigureAwait(false);

        var actionTcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        var processTcs = new TaskCompletionSource<Process?>(TaskCreationOptions.RunContinuationsAsynchronously);
        var dialogText = new StrongBox<string?>();

        var thread = new Thread(() =>
        {
            dynamic? app = null;
            int? processId = null;
            ModalDialogWatcher? watcher = null;
            Timer? cleanupTimer = null;

            try
            {
                cancellationToken.ThrowIfCancellationRequested();

                var type = Type.GetTypeFromProgID(_options.AccessComProgId)
                    ?? throw new InvalidOperationException("Microsoft Access is not installed.");

                app = Activator.CreateInstance(type)
                    ?? throw new InvalidOperationException("Failed to create Access COM instance.");

                processId = TryGetProcessId(app);
                if (processId is int pid)
                {
                    try
                    {
                        processTcs.TrySetResult(Process.GetProcessById(pid));
                    }
                    catch
                    {
                        processTcs.TrySetResult(null);
                    }
                }
                else
                {
                    processTcs.TrySetResult(null);
                }

                app.Visible = options.Visible;
                app.UserControl = false;

                if (options.LowAutomationSecurity)
                {
                    // Loosen macro security for this scratch session only. The database is a copy.
                    try { app.AutomationSecurity = 1; } catch { }
                }

                var workgroup = options.WorkgroupPath ?? _options.WorkgroupPath;
                if (!string.IsNullOrWhiteSpace(workgroup))
                {
                    app.DBEngine.SystemDB = workgroup;
                }

                var openPath = databasePath is null ? null : Path.GetFullPath(databasePath);
                string? startUpFormValue = null;
                var startUpFormRemoved = false;
                if (options.OpenDatabase && openPath is not null)
                {
                    if (options.SuppressStartupForm && _options.SuppressStartupForm)
                    {
                        startUpFormRemoved = TryRemoveStartUpForm((object)app, openPath, out startUpFormValue);
                    }

                    try
                    {
                        app.OpenCurrentDatabase(openPath, false);
                    }
                    finally
                    {
                        if (startUpFormRemoved)
                        {
                            RestoreStartUpForm((object)app, openPath, startUpFormValue);
                        }
                    }
                }

                if (options.WatchDialogs && processId is int watcherPid)
                {
                    watcher = new ModalDialogWatcher(watcherPid);
                    watcher.Start();
                }

                var result = action(app);
                dialogText.Value = watcher?.LastText;
                actionTcs.TrySetResult(result);
            }
            catch (Exception ex)
            {
                dialogText.Value = watcher?.LastText;
                actionTcs.TrySetException(ex);
            }
            finally
            {
                try
                {
                    if (app is not null)
                    {
                        // A modal dialog during Quit (for example the VBE "Save As" prompt) would
                        // block this worker forever and keep the automation gate. Arm a kill timer
                        // for the cleanup window; it is disarmed as soon as cleanup completes.
                        if (processId is int cleanupPid)
                        {
                            var timeout = options.CleanupTimeout;
                            cleanupTimer = new Timer(
                                _ =>
                                {
                                    try
                                    {
                                        using var process = Process.GetProcessById(cleanupPid);
                                        process.Kill(entireProcessTree: true);
                                    }
                                    catch
                                    {
                                        // Already exited.
                                    }
                                },
                                null,
                                timeout,
                                Timeout.InfiniteTimeSpan);
                        }

                        try
                        {
                            try { app.CloseCurrentDatabase(); } catch { }
                            try { app.Quit(); } catch { }
                        }
                        finally
                        {
                            try { cleanupTimer?.Dispose(); } catch { }
                            cleanupTimer = null;
                        }
                    }
                }
                catch
                {
                }
                finally
                {
                    ReleaseComObject(app);
                    app = null;

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    try { watcher?.Dispose(); } catch { }

                    if (processId is int pid)
                    {
                        EnsureProcessExited(pid);
                    }

                    AutomationGate.Release();
                }
            }
        })
        {
            IsBackground = true
        };

        thread.SetApartmentState(ApartmentState.STA);
        StartHoldingGate(thread);

        var timeout = options.Timeout ?? TimeSpan.FromSeconds(Math.Max(1, _options.VbaTimeoutSeconds));

        Task completed;
        using (var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
        {
            var actionTask = actionTcs.Task;
            var delayTask = Task.Delay(timeout, timeoutCts.Token);
            completed = await Task.WhenAny(actionTask, delayTask).ConfigureAwait(false);
            if (completed == actionTask)
            {
                timeoutCts.Cancel();
            }
        }

        if (completed != actionTcs.Task)
        {
            var process = await WaitForProcessAsync(processTcs, TimeSpan.FromSeconds(2)).ConfigureAwait(false);
            KillProcess(process);

            // The worker may complete later with the COM error caused by the kill; observe it so it
            // does not surface as an unobserved task exception, and let the gate release.
            _ = actionTcs.Task.ContinueWith(static t => _ = t.Exception, TaskScheduler.Default);
            actionTcs.TrySetCanceled();

            return AccessRunOutcome<T>.TimedOutResult(timeout, dialogText.Value);
        }

        if (actionTcs.Task.IsCanceled || cancellationToken.IsCancellationRequested)
        {
            var process = await WaitForProcessAsync(processTcs, TimeSpan.FromSeconds(2)).ConfigureAwait(false);
            KillProcess(process);
            return AccessRunOutcome<T>.CancelledResult(dialogText.Value);
        }

        if (actionTcs.Task.IsFaulted)
        {
            var exception = actionTcs.Task.Exception?.GetBaseException() ?? new InvalidOperationException("Access action failed.");

            if (exception is OperationCanceledException)
            {
                var process = await WaitForProcessAsync(processTcs, TimeSpan.FromSeconds(2)).ConfigureAwait(false);
                KillProcess(process);
                return AccessRunOutcome<T>.CancelledResult(dialogText.Value);
            }

            var process2 = await WaitForProcessAsync(processTcs, TimeSpan.FromSeconds(2)).ConfigureAwait(false);
            var crashed = IsCrashException(exception);
            var exitCode = crashed ? TryGetExitCode(process2) : null;

            if (crashed)
            {
                KillProcess(process2);
            }

            return AccessRunOutcome<T>.Failure(exception, crashed, exitCode, dialogText.Value);
        }

        return AccessRunOutcome<T>.Success(actionTcs.Task.Result, dialogText.Value);
    }

    private static async Task<Process?> WaitForProcessAsync(TaskCompletionSource<Process?> processTcs, TimeSpan timeout)
    {
        var completed = await Task.WhenAny(processTcs.Task, Task.Delay(timeout)).ConfigureAwait(false);
        return completed == processTcs.Task && processTcs.Task.Status == TaskStatus.RanToCompletion
            ? processTcs.Task.Result
            : null;
    }

    private static void KillProcess(Process? process)
    {
        if (process is null)
        {
            return;
        }

        try
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
                process.WaitForExit(5000);
            }
        }
        catch
        {
            // Raced with exit or access denied; nothing more to do.
        }
    }

    private static int? TryGetExitCode(Process? process)
    {
        try
        {
            return process is { HasExited: true } ? process.ExitCode : null;
        }
        catch
        {
            return null;
        }
    }

    internal static bool IsCrashException(Exception exception)
    {
        if (exception is not COMException com)
        {
            return false;
        }

        // RPC server unavailable / call failed / disconnected: the Access process died or stopped
        // answering, rather than VBA raising a runtime error.
        return (uint)com.HResult is 0x800706BA or 0x800706BE or 0x80010108 or 0x800706BF;
    }

    /// <summary>
    /// Temporarily removes the database's StartUpForm property so opening it in Access doesn't run the
    /// configured start-up form (which can sit on a modal dialog and block automation). Uses the Access
    /// instance's own DAO engine so bitness and workgroup settings match. Returns true when the caller
    /// must restore <paramref name="originalValue"/> after opening the database: either a property was
    /// removed now, or an earlier session removed one and died before putting it back.
    /// </summary>
    private static bool TryRemoveStartUpForm(object appObject, string databasePath, out string? originalValue)
    {
        originalValue = null;

        dynamic app = appObject;

        // Probe read-only first, so a database without a start-up form is never opened exclusively or
        // written to.
        var current = TryReadStartUpForm(appObject, databasePath);
        if (string.IsNullOrEmpty(current))
        {
            originalValue = StartUpFormRecovery.Default.TryLoad(databasePath);
            return originalValue is not null;
        }

        // Record the value before deleting it, so a kill before RestoreStartUpForm is recoverable.
        try { StartUpFormRecovery.Default.Save(databasePath, current); } catch { }

        dynamic? database = null;
        try
        {
            // Exclusive: editing a database property requires the file to be locked.
            database = app.DBEngine.OpenDatabase(databasePath, true, false);
            database.Properties.Delete(StartUpFormPropertyName);
            originalValue = current;
            return true;
        }
        catch
        {
            // Best effort: the file can't be locked exclusively right now (another user has it open).
            // Leave it untouched and let Access start normally.
            StartUpFormRecovery.Default.Clear(databasePath);
            return false;
        }
        finally
        {
            if (database is not null)
            {
                try { database.Close(); } catch { }
                ReleaseComObject(database);
            }
        }
    }

    /// <summary>Reads StartUpForm through a shared, read-only DAO open. Null when there is none or the file can't be read.</summary>
    private static string? TryReadStartUpForm(object appObject, string databasePath)
    {
        dynamic app = appObject;
        dynamic? database = null;
        try
        {
            database = app.DBEngine.OpenDatabase(databasePath, false, true);
            return database.Properties.Item(StartUpFormPropertyName).Value as string;
        }
        catch
        {
            return null;
        }
        finally
        {
            if (database is not null)
            {
                try { database.Close(); } catch { }
                ReleaseComObject(database);
            }
        }
    }

    /// <summary>
    /// Restores a StartUpForm property removed by <see cref="TryRemoveStartUpForm"/>, preferring the
    /// live database (so the file's normal start-up behaviour is preserved) and falling back to a
    /// direct DAO open if Access never opened it. The recovery record is cleared only once the value
    /// is back, so a failed restore is retried by the next session.
    /// </summary>
    private static void RestoreStartUpForm(object appObject, string databasePath, string? originalValue)
    {
        if (originalValue is null)
        {
            return;
        }

        dynamic app = appObject;
        dynamic? database = null;
        var openedDirectly = false;
        try
        {
            try
            {
                database = app.CurrentDb();
            }
            catch
            {
                database = null;
            }

            if (database is null)
            {
                database = app.DBEngine.OpenDatabase(databasePath, true, false);
                openedDirectly = true;
            }

            var properties = database.Properties;
            try { properties.Delete(StartUpFormPropertyName); } catch { }

            var property = database.CreateProperty(StartUpFormPropertyName, DbText, originalValue);
            properties.Append(property);
            ReleaseComObject(property);
            StartUpFormRecovery.Default.Clear(databasePath);
        }
        catch
        {
            // The recovery record is still there, so the next session on this database retries.
        }
        finally
        {
            if (openedDirectly && database is not null)
            {
                try { database.Close(); } catch { }
            }

            ReleaseComObject(database);
        }
    }

    /// <summary>
    /// Maps a friendly control type name (for example, textbox) to the Access control type constant.
    /// </summary>
    public static int ResolveControlType(string controlType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(controlType);

        if (ControlTypes.TryGetValue(controlType.Trim(), out var value))
        {
            return value;
        }

        throw new ArgumentException("Unsupported control type.", nameof(controlType));
    }

    /// <summary>
    /// Resolves the friendly name for a control type constant, or returns the numeric value.
    /// </summary>
    public static string GetControlTypeName(int controlType)
    {
        foreach (var (name, value) in ControlTypes)
        {
            if (value == controlType)
            {
                return name;
            }
        }

        return controlType.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Maps a friendly section name (detail, header, footer) to the Access section constant.
    /// </summary>
    public static int ResolveSection(string section)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(section);

        if (int.TryParse(section, NumberStyles.Integer, CultureInfo.InvariantCulture, out var numeric))
        {
            return numeric;
        }

        if (Sections.TryGetValue(section.Trim(), out var value))
        {
            return value;
        }

        throw new ArgumentException("Unsupported section name.", nameof(section));
    }

    /// <summary>
    /// Resolves the Access form view constant for design or normal view.
    /// </summary>
    public static int ResolveFormView(bool designView)
        => designView ? AcViewDesign : AcViewNormal;

    /// <summary>
    /// Resolves the Access report view constant for the requested view name.
    /// </summary>
    public static int ResolveReportView(string view)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(view);

        return view.Trim().ToLowerInvariant() switch
        {
            "design" => AcViewDesign,
            "layout" => AcViewLayout,
            "preview" => AcViewPreview,
            "report" => AcViewReport,
            "normal" => AcViewNormal,
            _ => throw new ArgumentException("Unsupported report view.", nameof(view))
        };
    }

    /// <summary>
    /// Resolves the Access close behavior constant for whether to save design changes.
    /// </summary>
    public static int ResolveCloseSave(bool saveDesignChanges)
        => saveDesignChanges ? AcSaveYes : AcSaveNo;

    /// <summary>
    /// Access object type constant for forms.
    /// </summary>
    public static int FormObjectType => AcFormObjectType;

    /// <summary>
    /// Access object type constant for reports.
    /// </summary>
    public static int ReportObjectType => AcReportObjectType;

    /// <summary>
    /// Access DoCmd command constant for saving a record.
    /// </summary>
    public static int SaveRecordCommand => AcCmdSaveRecord;

    /// <summary>
    /// Access DoCmd command constant for compiling VBA modules.
    /// </summary>
    public static int CompileAllModulesCommand => AcCmdCompileAllModules;

    /// <summary>
    /// Access DoCmd command constant for compiling and saving VBA modules.
    /// </summary>
    public static int CompileAndSaveAllModulesCommand => AcCmdCompileAndSaveAllModules;

    /// <summary>
    /// Access DoCmd command constant for saving all VBA modules.
    /// </summary>
    public static int SaveAllModulesCommand => AcCmdSaveAllModules;

    /// <summary>
    /// Releases a COM object to avoid keeping Access running after automation.
    /// </summary>
    public static void ReleaseComObject(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject))
        {
            Marshal.FinalReleaseComObject(comObject);
        }
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    /// <summary>
    /// Resolves the OS process id of a freshly-created Access.Application instance via its main
    /// window handle, so it can be forcibly closed later if COM cleanup doesn't fully release it.
    /// Best-effort: returns null rather than throwing if anything about this lookup fails.
    /// </summary>
    private static int? TryGetProcessId(dynamic app)
    {
        try
        {
            var hwnd = new IntPtr(Convert.ToInt64(app.hWndAccessApp, CultureInfo.InvariantCulture));
            if (hwnd == IntPtr.Zero)
            {
                return null;
            }

            GetWindowThreadProcessId(hwnd, out var pid);
            return pid == 0 ? null : (int)pid;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Waits briefly for a specific Access process to exit on its own after Quit(), and kills it
    /// if it hasn't - a deterministic safety net against orphaned MSACCESS.EXE instances left behind
    /// by incomplete COM release. Runs on the automation thread after the caller's result is already
    /// set, so this never adds latency to a tool call.
    /// </summary>
    private static void EnsureProcessExited(int processId, int timeoutMs = 5000)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            if (!process.WaitForExit(timeoutMs))
            {
                process.Kill();
            }
        }
        catch (ArgumentException)
        {
            // Already exited - GetProcessById throws when the PID is no longer running.
        }
        catch (InvalidOperationException)
        {
            // Already exited between the check and Kill().
        }
    }

    // Access generates a COM type library per form/report, listing every control and every VBA
    // procedure as a member. The C# `dynamic` binder eagerly scans that whole type library
    // (IDispatchComObject.EnsureScanDefinedMethods) the first time a form/report object takes part in
    // any dynamic operation, and that scan blows up on real-world forms: duplicate member names that
    // differ only by case throw "Item has already been added", and forms whose module won't yield
    // full type info throw COMException 0x800A88D2 out of ITypeInfo.GetFuncDesc. Neither is
    // recoverable and neither has anything to do with the member we actually wanted.
    //
    // The helpers below go through IDispatch by name (GetIDsOfNames/Invoke) via Type.InvokeMember,
    // which resolves one member at a time and never scans the type library. Form, report, control and
    // module objects MUST be held in statically-typed `object` locals and read through these helpers
    // — putting one into a `dynamic` expression reintroduces the scan and the crash.

    private const System.Reflection.BindingFlags GetPropertyFlags =
        System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.InvokeMethod;

    /// <summary>
    /// Reads a COM property by name without invoking the C# dynamic binder. Throws if the member is missing.
    /// </summary>
    public static object? GetProperty(object target, string propertyName, params object?[] args)
    {
        ArgumentNullException.ThrowIfNull(target);

        return target.GetType().InvokeMember(
            propertyName,
            GetPropertyFlags,
            binder: null,
            target,
            args.Length == 0 ? null : args,
            CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Calls a COM method by name without invoking the C# dynamic binder.
    /// </summary>
    public static object? InvokeMethod(object target, string methodName, params object?[] args)
    {
        ArgumentNullException.ThrowIfNull(target);

        return target.GetType().InvokeMember(
            methodName,
            System.Reflection.BindingFlags.InvokeMethod,
            binder: null,
            target,
            args.Length == 0 ? null : args,
            CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Reads an item from a COM collection by index or key without invoking the C# dynamic binder.
    /// </summary>
    public static object? GetItem(object collection, object index)
        => GetProperty(collection, "Item", index);

    /// <summary>
    /// Reads a COM collection's Count without invoking the C# dynamic binder.
    /// </summary>
    public static int GetCount(object collection)
        => Convert.ToInt32(GetProperty(collection, "Count"), CultureInfo.InvariantCulture);

    /// <summary>
    /// Reads a COM property by name, returning null when the property is missing or unreadable.
    /// </summary>
    public static object? TryGetProperty(object target, string propertyName)
    {
        try
        {
            return GetProperty(target, propertyName);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Sets a COM property by name without invoking the C# dynamic binder. Throws if the assignment fails.
    /// </summary>
    public static void SetProperty(object target, string propertyName, object? value)
    {
        ArgumentNullException.ThrowIfNull(target);

        target.GetType().InvokeMember(
            propertyName,
            System.Reflection.BindingFlags.SetProperty,
            binder: null,
            target,
            [value],
            CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Attempts to set a COM property and returns whether the assignment succeeded.
    /// </summary>
    public static bool TrySetProperty(object target, string propertyName, object? value)
    {
        try
        {
            SetProperty(target, propertyName, value);
            return true;
        }
        catch
        {
            return false;
        }
    }
}

/// <summary>
/// Options for a single Access automation session. The defaults preserve the behaviour of the
/// existing <see cref="AccessComAutomation.RunInAccessAsync{T}"/> path.
/// </summary>
internal sealed record AccessRunOptions
{
    /// <summary>Wall-clock limit for the action. Null uses <see cref="AccessMcpOptions.VbaTimeoutSeconds"/>.</summary>
    public TimeSpan? Timeout { get; init; }

    /// <summary>How long Access gets to close cleanly before it is killed.</summary>
    public TimeSpan CleanupTimeout { get; init; } = TimeSpan.FromSeconds(10);

    /// <summary>Whether the Access window is visible. False by default.</summary>
    public bool Visible { get; init; }

    /// <summary>Open <c>databasePath</c> before running the action. False for instances that only build an MDE.</summary>
    public bool OpenDatabase { get; init; } = true;

    /// <summary>Remove and restore the configured StartUpForm around opening the database.</summary>
    public bool SuppressStartupForm { get; init; } = true;

    /// <summary>Set AutomationSecurity to msoAutomationSecurityLow (1) before opening the database.</summary>
    public bool LowAutomationSecurity { get; init; } = true;

    /// <summary>Start the Win32 modal-dialog watcher for the session.</summary>
    public bool WatchDialogs { get; init; } = true;

    /// <summary>Override the configured .mdw workgroup file.</summary>
    public string? WorkgroupPath { get; init; }

    public static AccessRunOptions Default { get; } = new();
}

/// <summary>
/// Structured result of an Access session: a value, a timeout, a crash, or a normal exception.
/// VBA runtime errors surface as <see cref="Exception"/> with <see cref="Crashed"/> false.
/// </summary>
internal sealed record AccessRunOutcome<T>(
    T? Value,
    bool TimedOut,
    bool Crashed,
    bool Cancelled,
    int? ExitCode,
    Exception? Exception,
    string? DialogText)
{
    public bool Ok => !TimedOut && !Crashed && !Cancelled && Exception is null;

    public static AccessRunOutcome<T> Success(T value, string? dialogText = null)
        => new(value, false, false, false, null, null, dialogText);

    public static AccessRunOutcome<T> TimedOutResult(TimeSpan timeout, string? dialogText = null)
        => new(default, true, false, false, null, new TimeoutException($"Access did not finish within {timeout.TotalSeconds:0.#}s."), dialogText);

    public static AccessRunOutcome<T> CancelledResult(string? dialogText = null)
        => new(default, false, false, true, null, new OperationCanceledException("The Access session was cancelled."), dialogText);

    public static AccessRunOutcome<T> Failure(Exception exception, bool crashed, int? exitCode, string? dialogText = null)
        => new(default, false, crashed, false, exitCode, exception, dialogText);
}
