using System.Globalization;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>A reference in the VBA project that cannot be resolved.</summary>
internal sealed record VbaBrokenReference(string Name, string? FullPath, bool IsBroken);

/// <summary>Result of the compile gate.</summary>
internal sealed record VbaCompileResult(
    bool Ok,
    bool IsCompiled,
    string? Module,
    int? Line,
    string? Message,
    string? DialogText,
    IReadOnlyList<VbaBrokenReference> BrokenReferences)
{
    /// <summary>True when Access itself reports a broken reference (Application.BrokenReference).</summary>
    public bool BrokenReferenceFlag { get; init; }

    /// <summary>The VBE compiler's error number, when the add-in could read the record.</summary>
    public int? ErrorNumber { get; init; }

    /// <summary>The VBA project name from the record, when available.</summary>
    public string? Project { get; init; }
}

/// <summary>
/// Compiles the VBA project and reports whether it succeeded. Access has no structured compiler
/// API: <c>Application.IsCompiled</c> is the reliable success check, and the code pane's caret is a
/// best-effort hint at the failing module. A compile error normally returns in milliseconds without
/// a modal dialog; the dialog watcher exists for the cases where Access does show one.
/// </summary>
internal static class VbaCompiler
{
    public const int CompileAndSaveAllModulesCommand = 126;
    private const int AcModuleObjectType = 4;
    private const int AcViewDesign = 1;
    private const int AcSaveYes = 1;

    public static VbaCompileResult Compile(object appObject, ModalDialogWatcher? watcher = null)
    {
        dynamic app = appObject;

        var brokenReference = ReadBrokenReferenceFlag((object)app);
        var references = ReadReferences((object)app);
        var broken = references.Where(r => r.IsBroken).ToArray();
        if (brokenReference)
        {
            return new VbaCompileResult(
                false,
                false,
                null,
                null,
                "The VBA project has broken references: " + string.Join(", ", broken.Select(b => b.Name)) + ".",
                watcher?.LastText,
                references)
            {
                BrokenReferenceFlag = true
            };
        }

        var anchor = FindAnchor((object)app);

        object? vbe = null;
        VbeAddinClient? addin = null;
        try
        {
            try
            {
                vbe = ((dynamic)app).VBE;
                addin = VbeAddinClient.TryConnect(vbe);
            }
            catch
            {
                addin = null;
            }

            TryOpenAnchor((object)app, anchor);
            RunCompile((object)app);
            var compiled = ReadIsCompiled((object)app);
            var record = addin?.ReadErrorRecord();

            if (compiled == true)
            {
                return new VbaCompileResult(true, true, null, null, null, watcher?.LastText, references)
                {
                    BrokenReferenceFlag = brokenReference
                };
            }

            if (record is { Present: true })
            {
                return RecordFailure(record, references, watcher);
            }

            if (compiled == false)
            {
                // Some projects only compile fully when a code pane (or a form in design view) is open.
                // Mirror AutoBuilder's one nudge before giving up.
                if (!TryNudge((object)app, anchor))
                {
                    RunCompile((object)app);
                }

                compiled = ReadIsCompiled((object)app);
                record = addin?.ReadErrorRecord();

                if (compiled == true)
                {
                    return new VbaCompileResult(true, true, null, null, null, watcher?.LastText, references)
                    {
                        BrokenReferenceFlag = brokenReference
                    };
                }

                if (record is { Present: true })
                {
                    return RecordFailure(record, references, watcher);
                }
            }

            if (compiled != false)
            {
                return new VbaCompileResult(true, compiled ?? true, null, null, null, watcher?.LastText, references)
                {
                    BrokenReferenceFlag = brokenReference
                };
            }

            var caret = TryGetCaret((object)app);
            var message = "The VBA project did not compile (Application.IsCompiled is false).";

            var dialog = watcher?.LastText;
            if (!string.IsNullOrWhiteSpace(dialog))
            {
                message += " Access reported: " + dialog;
            }

            return new VbaCompileResult(false, false, caret.Module, caret.Line, message, dialog, references)
            {
                BrokenReferenceFlag = brokenReference
            };
        }
        finally
        {
            addin?.Dispose();
            AccessComAutomation.ReleaseComObject(vbe);
        }
    }

    private static VbaCompileResult RecordFailure(
        VbeCompileRecord record,
        IReadOnlyList<VbaBrokenReference> references,
        ModalDialogWatcher? watcher)
    {
        var message = !string.IsNullOrWhiteSpace(record.Message)
            ? record.Message
            : record.ErrorNumber is int number && VbeCompileErrorTable.TryGet(number, out var known)
                ? known
                : record.ErrorNumber is int unknownNumber
                    ? $"VBE compile error {unknownNumber}."
                    : "The VBA project did not compile.";

        return new VbaCompileResult(
            false,
            false,
            string.IsNullOrWhiteSpace(record.Module) ? null : record.Module,
            record.Line is > 0 ? record.Line : null,
            message,
            watcher?.LastText,
            references)
        {
            ErrorNumber = record.ErrorNumber,
            Project = record.Project
        };
    }

    /// <summary>Reads the compile flag, or null when Access does not expose it.</summary>
    public static bool? ReadIsCompiled(object appObject)
    {
        try
        {
            dynamic app = appObject;
            return Convert.ToBoolean(app.IsCompiled, CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }

    private static void RunCompile(object appObject)
    {
        dynamic app = appObject;
        dynamic? doCmd = null;
        try
        {
            doCmd = app.DoCmd;
            doCmd.RunCommand(CompileAndSaveAllModulesCommand);
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(doCmd);
        }
    }

    private static bool TryNudge(object appObject, string? anchor)
    {
        dynamic app = appObject;
        dynamic? doCmd = null;
        try
        {
            doCmd = app.DoCmd;

            if (anchor is not null && !anchor.StartsWith("Form_", StringComparison.OrdinalIgnoreCase))
            {
                try { doCmd.OpenModule(anchor); } catch { }
                try { doCmd.SelectObject(AcModuleObjectType, anchor); } catch { }
                doCmd.RunCommand(CompileAndSaveAllModulesCommand);
                return true;
            }

            var formName = FindFormName(app);
            if (formName is not null)
            {
                try
                {
                    doCmd.OpenForm(formName, AcViewDesign);
                    doCmd.RunCommand(CompileAndSaveAllModulesCommand);
                    doCmd.Close(2, formName, AcSaveYes);
                    return true;
                }
                catch
                {
                    try { doCmd.Close(2, formName, AcSaveYes); } catch { }
                }
            }

            return false;
        }
        catch
        {
            return false;
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(doCmd);
        }
    }

    private static void TryOpenAnchor(object appObject, string? anchor)
    {
        dynamic app = appObject;
        if (anchor is null)
        {
            return;
        }

        dynamic? doCmd = null;
        try
        {
            doCmd = app.DoCmd;
            if (anchor.StartsWith("Form_", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var formName = anchor["Form_".Length..];
                    doCmd.OpenForm(formName, AcViewDesign);
                }
                catch
                {
                }
            }
            else
            {
                try { doCmd.OpenModule(anchor); } catch { }
                try { doCmd.SelectObject(AcModuleObjectType, anchor); } catch { }
            }
        }
        catch
        {
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(doCmd);
        }
    }

    /// <summary>
    /// Chooses a code context for the compiler: the first standard module with code, else the first
    /// form/report document module.
    /// </summary>
    private static string? FindAnchor(object appObject)
    {
        dynamic app = appObject;
        object? vbe = null;
        object? project = null;
        object? components = null;

        try
        {
            vbe = app.VBE;
            project = VbaProjectEditor.GetProject(vbe, null);
            components = AccessComAutomation.GetProperty(project, "VBComponents");
            var count = AccessComAutomation.GetCount(components!);

            string? form = null;
            for (var i = 1; i <= count; i++)
            {
                object? component = null;
                object? codeModule = null;
                try
                {
                    component = AccessComAutomation.GetItem(components!, i)!;
                    var name = Convert.ToString(AccessComAutomation.GetProperty(component, "Name"), CultureInfo.InvariantCulture) ?? string.Empty;
                    var type = Convert.ToInt32(AccessComAutomation.GetProperty(component, "Type"), CultureInfo.InvariantCulture);
                    codeModule = AccessComAutomation.GetProperty(component, "CodeModule");
                    var lines = Convert.ToInt32(AccessComAutomation.GetProperty(codeModule!, "CountOfLines"), CultureInfo.InvariantCulture);

                    if (type == VbaProjectEditor.StandardModule && lines > 0)
                    {
                        return name;
                    }

                    if (form is null && type is VbaProjectEditor.FormModule or VbaProjectEditor.DocumentModule)
                    {
                        form = name;
                    }
                }
                finally
                {
                    AccessComAutomation.ReleaseComObject(codeModule);
                    AccessComAutomation.ReleaseComObject(component);
                }
            }

            return form;
        }
        catch
        {
            return null;
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(components);
            AccessComAutomation.ReleaseComObject(project);
            AccessComAutomation.ReleaseComObject(vbe);
        }
    }

    private static string? FindFormName(object appObject)
    {
        dynamic app = appObject;
        try
        {
            dynamic? allForms = null;
            try
            {
                allForms = app.CurrentProject.AllForms;
                var count = Convert.ToInt32(allForms.Count, CultureInfo.InvariantCulture);
                if (count > 0)
                {
                    var form = allForms.Item(0);
                    try
                    {
                        return Convert.ToString(form.Name, CultureInfo.InvariantCulture);
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(form);
                    }
                }
            }
            finally
            {
                AccessComAutomation.ReleaseComObject(allForms);
            }
        }
        catch
        {
        }

        return null;
    }

    private static bool ReadBrokenReferenceFlag(object appObject)
    {
        try
        {
            dynamic app = appObject;
            return Convert.ToBoolean(app.BrokenReference, CultureInfo.InvariantCulture);
        }
        catch
        {
            return false;
        }
    }

    private static IReadOnlyList<VbaBrokenReference> ReadReferences(object appObject)
    {
        dynamic app = appObject;
        var result = new List<VbaBrokenReference>();
        object? references = null;

        try
        {
            references = app.References;
            var count = Convert.ToInt32(AccessComAutomation.GetProperty(references!, "Count"), CultureInfo.InvariantCulture);
            for (var i = 1; i <= count; i++)
            {
                object? reference = null;
                try
                {
                    reference = AccessComAutomation.GetItem(references!, i);
                    if (reference is null)
                    {
                        continue;
                    }

                    var name = Convert.ToString(AccessComAutomation.TryGetProperty(reference, "Name"), CultureInfo.InvariantCulture) ?? $"#{i}";
                    var path = Convert.ToString(AccessComAutomation.TryGetProperty(reference, "FullPath"), CultureInfo.InvariantCulture);
                    var isBroken = Convert.ToBoolean(AccessComAutomation.TryGetProperty(reference, "IsBroken") ?? false, CultureInfo.InvariantCulture);

                    result.Add(new VbaBrokenReference(name, path, isBroken));
                }
                finally
                {
                    AccessComAutomation.ReleaseComObject(reference);
                }
            }
        }
        catch
        {
            // References are best-effort diagnostics.
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(references);
        }

        return result;
    }

    private static (string? Module, int? Line) TryGetCaret(object appObject)
    {
        dynamic app = appObject;
        object? pane = null;
        object? codeModule = null;
        object? parent = null;

        try
        {
            pane = app.VBE.ActiveCodePane;
            if (pane is null)
            {
                return (null, null);
            }

            codeModule = AccessComAutomation.GetProperty(pane, "CodeModule");
            parent = AccessComAutomation.GetProperty(codeModule!, "Parent");
            var moduleName = Convert.ToString(AccessComAutomation.GetProperty(parent!, "Name"), CultureInfo.InvariantCulture);

            // GetSelection uses four out parameters. The DLR handles them for a COM dispatch call.
            dynamic dynamicPane = pane;
            int startLine = 0, startColumn = 0, endLine = 0, endColumn = 0;
            dynamicPane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);

            var lines = Convert.ToInt32(AccessComAutomation.GetProperty(codeModule!, "CountOfLines"), CultureInfo.InvariantCulture);
            if (startLine < 1 || startLine > lines)
            {
                return (moduleName, null);
            }

            var text = Convert.ToString(
                AccessComAutomation.GetProperty(codeModule!, "Lines", startLine, 1),
                CultureInfo.InvariantCulture) ?? string.Empty;

            // The caret is often parked past the real error. Only report a line that is a plausible
            // statement, never an option, blank or End line.
            var trimmed = text.Trim();
            if (trimmed.Length == 0 ||
                trimmed.StartsWith("Option ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("End ", StringComparison.OrdinalIgnoreCase))
            {
                return (moduleName, null);
            }

            return (moduleName, startLine);
        }
        catch
        {
            return (null, null);
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(parent);
            AccessComAutomation.ReleaseComObject(codeModule);
            AccessComAutomation.ReleaseComObject(pane);
        }
    }
}
