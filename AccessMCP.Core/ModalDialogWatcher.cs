using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>A button on a Win32 dialog, identified by its standard control id.</summary>
internal enum DialogButton
{
    None = 0,
    Ok = 1,
    Cancel = 2,
    Yes = 6,
    No = 7
}

/// <summary>One child control of a dialog.</summary>
internal sealed record DialogControl(int Id, string ClassName, string Text);

/// <summary>Snapshot of a modal dialog owned by the Access process.</summary>
internal sealed record DialogInfo(
    long Handle,
    int ProcessId,
    string ClassName,
    string Title,
    string StaticText,
    IReadOnlyList<DialogControl> Controls)
{
    public DialogControl? FindControl(int id) => Controls.FirstOrDefault(c => c.Id == id);

    public bool HasEdit(int id) => Controls.Any(c => c.Id == id && c.ClassName.Contains("Edit", StringComparison.OrdinalIgnoreCase));

    public string CombinedText => string.IsNullOrEmpty(StaticText) ? Title : Title + " " + StaticText;
}

/// <summary>What the watcher should do with a dialog.</summary>
internal sealed record DialogAction(DialogButton Button, string? EditText = null)
{
    public static DialogAction None { get; } = new(DialogButton.None);
    public static DialogAction Ok { get; } = new(DialogButton.Ok);
    public static DialogAction Cancel { get; } = new(DialogButton.Cancel);
}

/// <summary>
/// Watches top-level windows owned by a process and dismisses the known Access/VBE modal dialogs
/// that would otherwise wedge COM automation: the VBE "Save As" module prompt and the
/// compile-error/save-changes message boxes.
/// <para>
/// Runs on its own background thread because the COM caller is blocked inside the modal call. It
/// deliberately matches control ids and window classes, never captions, so it is localization-safe.
/// </para>
/// </summary>
internal sealed class ModalDialogWatcher : IDisposable
{
    private const int WM_SETTEXT = 0x000C;
    private const int BM_CLICK = 0x00F5;
    private const int SaveAsModuleNameEditId = 2020;

    private readonly int _processId;
    private readonly Func<DialogInfo, DialogAction> _policy;
    private readonly TimeSpan _pollInterval;
    private readonly ConcurrentQueue<DialogInfo> _seen = new();
    private readonly CancellationTokenSource _stop = new();
    private readonly Thread _thread;

    public ModalDialogWatcher(int processId, Func<DialogInfo, DialogAction>? policy = null, TimeSpan? pollInterval = null)
    {
        _processId = processId;
        _policy = policy ?? DefaultPolicy;
        _pollInterval = pollInterval ?? TimeSpan.FromMilliseconds(150);
        _thread = new Thread(Loop)
        {
            IsBackground = true,
            Name = "AccessMCP.ModalDialogWatcher"
        };
    }

    public IReadOnlyList<DialogInfo> Seen => _seen.ToArray();

    public string? LastText => _seen.LastOrDefault()?.CombinedText;

    public void Start() => _thread.Start();

    public void Dispose()
    {
        try { _stop.Cancel(); } catch { }
        try { _thread.Join(TimeSpan.FromSeconds(2)); } catch { }
        _stop.Dispose();
    }

    /// <summary>
    /// The harness policy. A VBE "Save As" prompt (a RichEdit module-name edit with id 2020) is
    /// cancelled: the database itself is saved by Access, and accepting would write stray .bas files.
    /// Everything else with an OK/Yes button is acknowledged so the caller can continue.
    /// </summary>
    public static DialogAction DefaultPolicy(DialogInfo dialog)
    {
        if (dialog.HasEdit(SaveAsModuleNameEditId) && dialog.FindControl(ButtonId(DialogButton.Cancel)) is not null)
        {
            return DialogAction.Cancel;
        }

        if (dialog.FindControl(1) is not null)
        {
            return DialogAction.Ok;
        }

        return DialogAction.None;
    }

    private void Loop()
    {
        var handled = new HashSet<long>();
        while (!_stop.IsCancellationRequested)
        {
            try
            {
                foreach (var dialog in EnumerateDialogs(_processId))
                {
                    if (!handled.Add(dialog.Handle))
                    {
                        continue;
                    }

                    _seen.Enqueue(dialog);

                    DialogAction action;
                    try
                    {
                        action = _policy(dialog);
                    }
                    catch
                    {
                        action = DialogAction.None;
                    }

                    if (action.Button != DialogButton.None)
                    {
                        Act(dialog, action);
                    }
                }
            }
            catch
            {
                // Enumeration races with window creation/destruction; ignore and retry.
            }

            _stop.Token.WaitHandle.WaitOne(_pollInterval);
        }
    }

    private static void Act(DialogInfo dialog, DialogAction action)
    {
        var handle = new IntPtr(dialog.Handle);

        if (action.EditText is not null && dialog.FindControl(SaveAsModuleNameEditId) is not null)
        {
            var editHandle = FindChildHandle(handle, SaveAsModuleNameEditId);
            if (editHandle != IntPtr.Zero)
            {
                _ = SendMessage(editHandle, WM_SETTEXT, IntPtr.Zero, action.EditText);
            }
        }

        var buttonId = ButtonId(action.Button);
        var buttonHandle = FindChildHandle(handle, buttonId);
        if (buttonHandle != IntPtr.Zero)
        {
            _ = PostMessage(buttonHandle, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
        }
    }

    private static IEnumerable<DialogInfo> EnumerateDialogs(int processId)
    {
        var results = new List<DialogInfo>();

        _ = EnumWindows((hwnd, lParam) =>
        {
            GetWindowThreadProcessId(hwnd, out var pid);
            if (pid != (uint)processId)
            {
                return true;
            }

            var className = GetWindowTextCore(hwnd, GetClassName);
            if (!className.Equals("#32770", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var title = GetWindowTextCore(hwnd, GetWindowText);
            var controls = new List<DialogControl>();
            _ = EnumChildWindows(hwnd, (child, childParam) =>
            {
                var childClass = GetWindowTextCore(child, GetClassName);
                var text = GetWindowTextCore(child, GetWindowText);
                controls.Add(new DialogControl(GetDlgCtrlID(child), childClass, text));
                return true;
            }, IntPtr.Zero);

            var staticText = string.Join(
                Environment.NewLine,
                controls
                    .Where(c => c.ClassName.Equals("Static", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(c.Text))
                    .Select(c => c.Text));

            results.Add(new DialogInfo(
                hwnd.ToInt64(),
                processId,
                className,
                title,
                staticText,
                controls));

            return true;
        }, IntPtr.Zero);

        return results;
    }

    private static IntPtr FindChildHandle(IntPtr parent, int controlId)
    {
        var found = IntPtr.Zero;
        _ = EnumChildWindows(parent, (child, _) =>
        {
            if (GetDlgCtrlID(child) == controlId)
            {
                found = child;
                return false;
            }

            return true;
        }, IntPtr.Zero);

        return found;
    }

    private static string GetWindowTextCore(IntPtr hwnd, Func<IntPtr, StringBuilder, int, int> read)
    {
        var builder = new StringBuilder(1024);
        _ = read(hwnd, builder, builder.Capacity);
        return builder.ToString();
    }

    private static int GetWindowText(IntPtr hwnd, StringBuilder text, int count)
        => NativeGetWindowText(hwnd, text, count);

    private static int GetClassName(IntPtr hwnd, StringBuilder text, int count)
        => NativeGetClassName(hwnd, text, count);

    private static int ButtonId(DialogButton button) => button switch
    {
        DialogButton.Ok => 1,
        DialogButton.Cancel => 2,
        DialogButton.Yes => 6,
        DialogButton.No => 7,
        _ => 0
    };

    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetWindowTextW")]
    private static extern int NativeGetWindowText(IntPtr hwnd, StringBuilder text, int count);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetClassNameW")]
    private static extern int NativeGetClassName(IntPtr hwnd, StringBuilder text, int count);

    [DllImport("user32.dll")]
    private static extern int GetDlgCtrlID(IntPtr hwnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")]
    private static extern IntPtr SendMessage(IntPtr hwnd, int message, IntPtr wParam, string lParam);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hwnd, int message, IntPtr wParam, IntPtr lParam);
}
