using System.Runtime.InteropServices;
using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

/// <summary>The COM-side helpers that are pure mapping logic and run without Access.</summary>
public sealed class AccessComHelpersTests : IDisposable
{
    private readonly TempDirectory _dir = new();

    public void Dispose() => _dir.Dispose();

    [Theory]
    [InlineData("textbox", 109)]
    [InlineData("TextBox", 109)]
    [InlineData(" combobox ", 111)]
    [InlineData("label", 100)]
    [InlineData("checkbox", 106)]
    [InlineData("subform", 112)]
    public void ResolveControlType_MapsFriendlyNames(string name, int expected)
    {
        Assert.Equal(expected, AccessComAutomation.ResolveControlType(name));
    }

    [Fact]
    public void ResolveControlType_RejectsUnknownAndBlankNames()
    {
        Assert.Throws<ArgumentException>(() => AccessComAutomation.ResolveControlType("hologram"));
        Assert.ThrowsAny<ArgumentException>(() => AccessComAutomation.ResolveControlType(" "));
    }

    [Theory]
    [InlineData(109, "textbox")]
    [InlineData(104, "commandbutton")]
    [InlineData(999, "999")]
    public void GetControlTypeName_RoundTripsOrFallsBackToTheNumber(int type, string expected)
    {
        Assert.Equal(expected, AccessComAutomation.GetControlTypeName(type));
    }

    [Theory]
    [InlineData("detail", 0)]
    [InlineData("Header", 1)]
    [InlineData("pagefooter", 4)]
    [InlineData("grouplevel2footer", 8)]
    [InlineData("12", 12)]
    public void ResolveSection_AcceptsNamesAndNumbers(string section, int expected)
    {
        Assert.Equal(expected, AccessComAutomation.ResolveSection(section));
    }

    [Fact]
    public void ResolveSection_RejectsUnknownNames()
    {
        Assert.Throws<ArgumentException>(() => AccessComAutomation.ResolveSection("sidebar"));
    }

    [Theory]
    [InlineData("normal", 0)]
    [InlineData("design", 1)]
    [InlineData("Preview", 2)]
    [InlineData("report", 5)]
    [InlineData("layout", 6)]
    public void ResolveReportView_MapsViewNames(string view, int expected)
    {
        Assert.Equal(expected, AccessComAutomation.ResolveReportView(view));
    }

    [Fact]
    public void ResolveReportView_RejectsUnknownViews()
    {
        Assert.Throws<ArgumentException>(() => AccessComAutomation.ResolveReportView("datasheet"));
    }

    [Fact]
    public void FormViewAndCloseSaveConstantsMatchAccess()
    {
        Assert.Equal(0, AccessComAutomation.ResolveFormView(designView: false));
        Assert.Equal(1, AccessComAutomation.ResolveFormView(designView: true));
        Assert.Equal(1, AccessComAutomation.ResolveCloseSave(saveDesignChanges: true));
        Assert.Equal(2, AccessComAutomation.ResolveCloseSave(saveDesignChanges: false));
        Assert.Equal(2, AccessComAutomation.FormObjectType);
        Assert.Equal(3, AccessComAutomation.ReportObjectType);
        Assert.Equal(126, AccessComAutomation.CompileAndSaveAllModulesCommand);
    }

    [Theory]
    [InlineData(0x800706BA)]
    [InlineData(0x800706BE)]
    [InlineData(0x80010108)]
    [InlineData(0x800706BF)]
    public void IsCrashException_RecognisesDeadServerHResults(uint hresult)
    {
        Assert.True(AccessComAutomation.IsCrashException(new COMException("gone", unchecked((int)hresult))));
    }

    [Fact]
    public void IsCrashException_IgnoresVbaErrorsAndNonComExceptions()
    {
        // 0x800A000B is VBA error 11 (division by zero) surfaced through COM.
        Assert.False(AccessComAutomation.IsCrashException(new COMException("div", unchecked((int)0x800A000B))));
        Assert.False(AccessComAutomation.IsCrashException(new InvalidOperationException()));
    }

    [Fact]
    public void ReleaseComObject_IgnoresNullAndNonComObjects()
    {
        AccessComAutomation.ReleaseComObject(null);
        AccessComAutomation.ReleaseComObject(new object());
    }

    [Fact]
    public void DialogPolicy_CancelsTheVbeSaveAsPrompt()
    {
        var dialog = Dialog(
            new DialogControl(2020, "RichEdit20W", ""),
            new DialogControl(1, "Button", "OK"),
            new DialogControl(2, "Button", "Cancel"));

        Assert.Equal(DialogButton.Cancel, ModalDialogWatcher.DefaultPolicy(dialog).Button);
    }

    [Fact]
    public void DialogPolicy_AcknowledgesMessageBoxesWithAnOkButton()
    {
        var dialog = Dialog(new DialogControl(1, "Button", "OK"), new DialogControl(65535, "Static", "Compile error"));

        Assert.Equal(DialogButton.Ok, ModalDialogWatcher.DefaultPolicy(dialog).Button);
    }

    [Fact]
    public void DialogPolicy_LeavesUnknownDialogsAlone()
    {
        var dialog = Dialog(new DialogControl(6, "Button", "Yes"), new DialogControl(7, "Button", "No"));

        Assert.Equal(DialogButton.None, ModalDialogWatcher.DefaultPolicy(dialog).Button);
    }

    [Fact]
    public void DialogPolicy_EditWithoutCancelIsNotTheSaveAsPrompt()
    {
        // Matches control ids and classes, never captions, so it stays localization-safe.
        var dialog = Dialog(new DialogControl(2020, "Edit", ""), new DialogControl(1, "Button", "OK"));

        Assert.Equal(DialogButton.Ok, ModalDialogWatcher.DefaultPolicy(dialog).Button);
    }

    [Fact]
    public void DialogInfo_CombinedTextJoinsTitleAndBody()
    {
        Assert.Equal("Microsoft Access", Dialog().CombinedText);
        Assert.Equal("Microsoft Access Body", (Dialog() with { StaticText = "Body" }).CombinedText);
    }

    [Fact]
    public void ModalDialogWatcher_StartsAndStopsForAProcessWithNoDialogs()
    {
        using var watcher = new ModalDialogWatcher(Environment.ProcessId, pollInterval: TimeSpan.FromMilliseconds(10));
        watcher.Start();
        Thread.Sleep(50);

        Assert.Null(watcher.LastText);
    }

    [Theory]
    [InlineData(@"C:\Data\Sample.accdb", ".accde")]
    [InlineData(@"C:\Data\Sample.ACCDB", ".accde")]
    [InlineData(@"C:\Data\Sample.mdb", ".mde")]
    [InlineData(@"C:\Data\Sample.accde", ".accde")]
    [InlineData(@"C:\Data\Sample.mde", ".mde")]
    public void MdeDefaultExtension_FollowsTheSourceFormat(string source, string expected)
    {
        Assert.Equal(expected, MdeBuilder.DefaultExtension(source), ignoreCase: true);
    }

    [Fact]
    public void UsePrebuilt_CopiesACompiledDatabase()
    {
        var prebuilt = _dir.Touch("Sample.accde", [4, 5, 6]);
        var output = _dir.Combine(Path.Combine("out", "copy.accde"));

        var result = MdeBuilder.UsePrebuilt(prebuilt, output);

        Assert.True(result.Ok);
        Assert.True(result.Verified);
        Assert.Equal("prebuilt", result.Strategy);
        Assert.Equal([4, 5, 6], File.ReadAllBytes(output));
    }

    [Fact]
    public void UsePrebuilt_SamePathIsANoOp()
    {
        var prebuilt = _dir.Touch("Sample.mde", [1]);

        Assert.True(MdeBuilder.UsePrebuilt(prebuilt, prebuilt).Ok);
    }

    [Fact]
    public void UsePrebuilt_RefusesMissingOrUncompiledFiles()
    {
        var missing = MdeBuilder.UsePrebuilt(_dir.Combine("missing.accde"), _dir.Combine("out.accde"));
        var uncompiled = MdeBuilder.UsePrebuilt(_dir.Touch("Sample.accdb"), _dir.Combine("out.accde"));

        Assert.False(missing.Ok);
        Assert.Contains("not found", missing.Error);
        Assert.False(uncompiled.Ok);
        Assert.Contains(".mde or .accde", uncompiled.Error);
        Assert.False(File.Exists(_dir.Combine("out.accde")));
    }

    private static DialogInfo Dialog(params DialogControl[] controls)
        => new(1, 1, "#32770", "Microsoft Access", string.Empty, controls);
}
