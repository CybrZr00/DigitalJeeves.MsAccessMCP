using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>One VBE component, with the Access/VBIDE type number and a friendly name.</summary>
internal sealed record VbaComponentInfo(string Name, int Type, string TypeName, int LineCount)
{
    public bool IsStandardModule => Type == VbaProjectEditor.StandardModule;
    public bool IsClassModule => Type == VbaProjectEditor.ClassModule;
    public bool IsDocumentModule => Type == VbaProjectEditor.DocumentModule;
}

/// <summary>A procedure discovered by parsing a module's source text.</summary>
internal sealed record VbaProcedureInfo(
    string Name,
    string Kind,
    int StartLine,
    bool IsPublic,
    bool IsTest,
    bool IsSkipped,
    bool IsExpectedFailure);

/// <summary>Result of a verified module write.</summary>
internal sealed record VbaWriteResult(
    string ModuleName,
    bool Created,
    int Lines,
    string Sha256,
    string? BackupPath,
    string Code)
{
    /// <summary>
    /// True when the database knows the component as a module object. A component created with
    /// <c>VBComponents.Add</c> stays unregistered until a successful compile-and-save, and an
    /// unregistered component is never written to the file - Access asks for a module name in a
    /// modal "Save As" dialog instead, which automation cannot answer.
    /// </summary>
    public bool Registered { get; init; }

    /// <summary>True when Access only re-cased identifiers; the content itself matched.</summary>
    public bool CaseNormalized { get; init; }

    /// <summary>Number of lines whose only difference after the write was identifier casing.</summary>
    public int CaseNormalizedLines { get; init; }

    /// <summary>Tabs expanded to spaces before the write (VBE turns a tab into a line break).</summary>
    public int ExpandedTabs { get; init; }

    /// <summary>Lines whose trailing whitespace Access would have trimmed anyway.</summary>
    public int TrimmedLines { get; init; }
}

/// <summary>Source prepared for a <c>CodeModule</c> write, with what had to be normalized.</summary>
internal sealed record PreparedWrite(string Code, int ExpandedTabs, int TrimmedLines);

/// <summary>Outcome of comparing intended source with what <c>CodeModule</c> read back.</summary>
internal sealed record CodeComparison(
    bool Same,
    bool CaseOnlyDifferences,
    int CaseDifferenceLines,
    int? Line,
    string? Expected,
    string? Actual,
    int ExpectedLines,
    int ActualLines);

/// <summary>Thrown when source cannot be written through <c>CodeModule</c> at all.</summary>
internal sealed class ModuleWriteRefusedException(string message) : InvalidOperationException(message);

/// <summary>
/// Reads and writes VBA modules through the VBE object model with the same verified line-level
/// write used by <see cref="AccessTools"/>: never <c>AddFromString</c>, always read back and roll
/// back on mismatch. Also discovers <c>'@test</c> procedures for the test runner.
/// </summary>
internal static class VbaProjectEditor
{
    public const int StandardModule = 1;
    public const int ClassModule = 2;
    public const int FormModule = 3;
    public const int DocumentModule = 100;

    private static readonly Regex ProcedureRegex = new(
        @"^\s*(?:(?<access>Public|Private|Friend)\s+)?(?:(?<static>Static)\s+)?(?<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex DirectiveRegex = new(
        @"^\s*'@(?<directive>test(?:-skip|-xfail)?)\b",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex AttributeRegex = new(
        @"^\s*Attribute\s+VB_",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex OptionRegex = new(
        @"^\s*Option\s+(?:Explicit|Compare\s+\w+|Base\s+\d+|Private\s+Module)\b",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex VersionRegex = new(
        @"^\s*VERSION\s+\d+\.\d+\s+CLASS\b",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    /// <summary>
    /// <c>CodeModule.ReplaceLine</c>/<c>InsertLines</c> silently split anything longer than this, so a
    /// longer line can never be written as one line and is refused up front.
    /// </summary>
    public const int MaxCodeModuleLineLength = 1023;

    /// <summary>VBE's editor tab width; a tab advances to the next stop, it is not just four spaces.</summary>
    private const int TabWidth = 4;

    private const int CompileAndSaveAllModulesCommand = 126;

    /// <summary>
    /// Resolves the VBA project by name, or the first project when no name is supplied. The caller
    /// owns the returned COM object and must release it.
    /// </summary>
    public static object GetProject(object vbe, string? projectName)
    {
        object? projects = null;
        object? selected = null;

        try
        {
            projects = AccessComAutomation.GetProperty(vbe, "VBProjects");
            var count = AccessComAutomation.GetCount(projects!);

            // VBIDE collections are 1-based, unlike Access's own collections.
            for (var i = 1; i <= count; i++)
            {
                var project = AccessComAutomation.GetItem(projects!, i)!;
                var name = Convert.ToString(AccessComAutomation.GetProperty(project, "Name"), CultureInfo.InvariantCulture) ?? string.Empty;
                var matches = projectName is null ? selected is null : name.Equals(projectName, StringComparison.OrdinalIgnoreCase);

                if (matches)
                {
                    if (selected is not null)
                    {
                        AccessComAutomation.ReleaseComObject(selected);
                    }

                    selected = project;
                }
                else
                {
                    AccessComAutomation.ReleaseComObject(project);
                }
            }

            if (selected is null)
            {
                throw new ArgumentException($"VBA project '{projectName}' was not found.", nameof(projectName));
            }

            return selected;
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(projects);
        }
    }

    /// <summary>Lists components in a project. The caller owns nothing further.</summary>
    public static IReadOnlyList<VbaComponentInfo> ListComponents(object project)
    {
        object? components = null;
        var result = new List<VbaComponentInfo>();

        try
        {
            components = AccessComAutomation.GetProperty(project, "VBComponents");
            var count = AccessComAutomation.GetCount(components!);
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
                    result.Add(new VbaComponentInfo(name, type, GetTypeName(type), lines));
                }
                finally
                {
                    AccessComAutomation.ReleaseComObject(codeModule);
                    AccessComAutomation.ReleaseComObject(component);
                }
            }
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(components);
        }

        result.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
        return result;
    }

    /// <summary>
    /// Ensures a module exists and has exactly <paramref name="code"/>. Existing modules are backed
    /// up and written through the verified line-level writer; new modules are created with
    /// <c>VBComponents.Add</c> and named before the body is written. Access's own Option lines are
    /// preserved unless the supplied code specifies its own.
    /// </summary>
    public static VbaWriteResult EnsureModule(
        object project,
        string databasePath,
        string moduleName,
        string kind,
        string code,
        bool createIfMissing = true,
        bool allowOverwrite = true,
        string? backupDirectory = null,
        object? accessApp = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(moduleName);
        ValidateModuleName(moduleName);

        var descriptor = kind.Trim().ToLowerInvariant() switch
        {
            "standard" or "std" or "bas" => StandardModule,
            "class" or "cls" => ClassModule,
            _ => throw new ArgumentException("kind must be 'standard' or 'class'.", nameof(kind))
        };

        object? components = null;
        object? component = null;
        object? codeModule = null;

        try
        {
            components = AccessComAutomation.GetProperty(project, "VBComponents");
            component = FindComponent(components!, moduleName);

            var created = component is null;
            if (created && !createIfMissing)
            {
                throw new InvalidOperationException($"Module '{moduleName}' does not exist and createIfMissing is false.");
            }

            if (!created && !allowOverwrite)
            {
                throw new InvalidOperationException($"Module '{moduleName}' already exists and allowOverwrite is false.");
            }

            // Read the current body and prepare/validate the new one before touching the project, so a
            // refused write (an over-long line, say) cannot leave a newly created empty module behind.
            var existingCode = string.Empty;
            if (!created)
            {
                codeModule = AccessComAutomation.GetProperty(component!, "CodeModule");
                existingCode = ReadCode(codeModule!);
            }

            var prepared = PrepareForWrite(code, existingCode);
            ValidateLineLengths(prepared.Code, moduleName);

            var registered = false;
            if (created)
            {
                component = AccessComAutomation.InvokeMethod(components!, "Add", descriptor)
                    ?? throw new InvalidOperationException($"VBComponents.Add returned null for '{moduleName}'.");
                AccessComAutomation.SetProperty(component, "Name", moduleName);

                // Register the (still empty) component before any code is written - see the note on
                // CompileAndSave. Compiling an empty module cannot fail because of its own content, so
                // the body is free to be code that does not compile yet and still be persisted.
                if (accessApp is not null)
                {
                    registered = CompileAndSave(accessApp);
                }

                codeModule = AccessComAutomation.GetProperty(component, "CodeModule");
            }

            var backupPath = created
                ? null
                : BackupModuleCode(backupDirectory ?? DefaultBackupDirectory(databasePath), moduleName, existingCode);

            var comparison = WriteAndVerify(codeModule!, existingCode, prepared.Code);

            var finalReadBack = ReadCode(codeModule!);
            return new VbaWriteResult(
                moduleName,
                created,
                CountLines(finalReadBack),
                Sha256Hex(finalReadBack),
                backupPath,
                finalReadBack)
            {
                Registered = registered,
                CaseNormalized = comparison.CaseOnlyDifferences,
                CaseNormalizedLines = comparison.CaseDifferenceLines,
                ExpandedTabs = prepared.ExpandedTabs,
                TrimmedLines = prepared.TrimmedLines
            };
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(codeModule);
            AccessComAutomation.ReleaseComObject(component);
            AccessComAutomation.ReleaseComObject(components);
        }
    }

    /// <summary>
    /// Imports a .bas or .cls file through <c>VBComponents.Import</c>, preserving the file's
    /// Attribute VB_Name and class headers. Returns the imported component.
    /// </summary>
    public static VbaWriteResult ImportFile(object project, string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        var fullPath = Path.GetFullPath(filePath);

        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"Module file not found: {fullPath}", fullPath);
        }

        var extension = Path.GetExtension(fullPath);
        if (!extension.Equals(".bas", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".cls", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Only .bas and .cls files can be imported; forms need a designer.", nameof(filePath));
        }

        object? components = null;
        object? component = null;
        object? codeModule = null;

        try
        {
            components = AccessComAutomation.GetProperty(project, "VBComponents");
            component = AccessComAutomation.InvokeMethod(components!, "Import", fullPath)
                ?? throw new InvalidOperationException($"VBComponents.Import returned null for '{fullPath}'.");

            var name = Convert.ToString(AccessComAutomation.GetProperty(component, "Name"), CultureInfo.InvariantCulture) ?? string.Empty;
            codeModule = AccessComAutomation.GetProperty(component, "CodeModule");
            var code = ReadCode(codeModule!);

            return new VbaWriteResult(name, true, CountLines(code), Sha256Hex(code), null, code);
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(codeModule);
            AccessComAutomation.ReleaseComObject(component);
            AccessComAutomation.ReleaseComObject(components);
        }
    }

    /// <summary>Exports a module through <c>VBComponent.Export</c>, preserving headers.</summary>
    public static (string FilePath, int Lines, string Sha256) ExportFile(object project, string moduleName, string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(moduleName);
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        var fullPath = Path.GetFullPath(filePath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath) ?? ".");

        object? components = null;
        object? component = null;
        try
        {
            components = AccessComAutomation.GetProperty(project, "VBComponents");
            component = FindComponent(components!, moduleName)
                ?? throw new InvalidOperationException($"Module '{moduleName}' was not found.");

            AccessComAutomation.InvokeMethod(component, "Export", fullPath);
            var text = File.ReadAllText(fullPath);
            return (fullPath, CountLines(text), Sha256Hex(text));
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(component);
            AccessComAutomation.ReleaseComObject(components);
        }
    }

    /// <summary>
    /// Removes a module. Form/report code-behind is refused unconditionally; when
    /// <paramref name="onlyHarnessModules"/> is true, only modules whose name starts with
    /// <c>AccessMcp</c> can be removed.
    /// </summary>
    public static bool RemoveModule(object project, string moduleName, bool onlyHarnessModules)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(moduleName);

        if (moduleName.StartsWith("Form_", StringComparison.OrdinalIgnoreCase) ||
            moduleName.StartsWith("Report_", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"'{moduleName}' is form/report code-behind and cannot be removed by this tool.");
        }

        if (onlyHarnessModules && !moduleName.StartsWith("AccessMcp", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"'{moduleName}' was not created by the harness; pass onlyHarnessModules=false to remove it deliberately.");
        }

        object? components = null;
        object? component = null;
        try
        {
            components = AccessComAutomation.GetProperty(project, "VBComponents");
            component = FindComponent(components!, moduleName)
                ?? throw new InvalidOperationException($"Module '{moduleName}' was not found.");

            var type = Convert.ToInt32(AccessComAutomation.GetProperty(component, "Type"), CultureInfo.InvariantCulture);
            if (type == DocumentModule)
            {
                throw new InvalidOperationException(
                    $"'{moduleName}' is an Access document (form/report) module and cannot be removed by this tool.");
            }

            AccessComAutomation.InvokeMethod(components!, "Remove", component);
            return true;
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(component);
            AccessComAutomation.ReleaseComObject(components);
        }
    }

    /// <summary>Reads a module's full code.</summary>
    public static string ReadModuleCode(object project, string moduleName)
    {
        object? components = null;
        object? component = null;
        object? codeModule = null;
        try
        {
            components = AccessComAutomation.GetProperty(project, "VBComponents");
            component = FindComponent(components!, moduleName)
                ?? throw new InvalidOperationException($"Module '{moduleName}' was not found.");
            codeModule = AccessComAutomation.GetProperty(component, "CodeModule");
            return ReadCode(codeModule!);
        }
        finally
        {
            AccessComAutomation.ReleaseComObject(codeModule);
            AccessComAutomation.ReleaseComObject(component);
            AccessComAutomation.ReleaseComObject(components);
        }
    }

    /// <summary>
    /// Finds the module that owns a public procedure. Access's <c>Application.Run</c> only accepts
    /// unqualified procedure names, so a caller-facing "Module.Procedure" has to be resolved here.
    /// </summary>
    public static string? FindProcedureModule(object project, string procedureName)
    {
        foreach (var component in ListComponents(project))
        {
            if (component.IsDocumentModule && !component.Name.StartsWith("Form_", StringComparison.OrdinalIgnoreCase) &&
                !component.Name.StartsWith("Report_", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string code;
            try
            {
                code = ReadModuleCode(project, component.Name);
            }
            catch
            {
                continue;
            }

            if (FindProcedures(code).Any(p => p.IsPublic && p.Name.Equals(procedureName, StringComparison.OrdinalIgnoreCase)))
            {
                return component.Name;
            }
        }

        return null;
    }

    /// <summary>Parses procedures and <c>'@test</c> directives from module source.</summary>
    public static IReadOnlyList<VbaProcedureInfo> FindProcedures(string code)
    {
        var result = new List<VbaProcedureInfo>();
        var lines = SplitLines(code);

        var pendingDirective = (string?)null;
        var blankSinceDirective = false;

        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            var trimmed = line.Trim();

            if (trimmed.Length == 0)
            {
                if (pendingDirective is not null)
                {
                    blankSinceDirective = true;
                }

                continue;
            }

            if (trimmed.StartsWith("'", StringComparison.Ordinal) || trimmed.StartsWith("Rem ", StringComparison.OrdinalIgnoreCase))
            {
                if (!blankSinceDirective)
                {
                    var match = DirectiveRegex.Match(line);
                    if (match.Success)
                    {
                        pendingDirective = match.Groups["directive"].Value.ToLowerInvariant();
                    }
                }

                continue;
            }

            var procedure = ProcedureRegex.Match(line);
            if (!procedure.Success)
            {
                // Any other code ends the directive's applicability.
                pendingDirective = null;
                blankSinceDirective = false;
                continue;
            }

            var access = procedure.Groups["access"].Value;
            var isPublic = access.Length == 0 || access.Equals("Public", StringComparison.OrdinalIgnoreCase);
            var kind = Regex.Replace(procedure.Groups["kind"].Value, @"\s+", " ");

            result.Add(new VbaProcedureInfo(
                procedure.Groups["name"].Value,
                kind,
                i + 1,
                isPublic,
                pendingDirective is "test" or "test-skip" or "test-xfail",
                pendingDirective == "test-skip",
                pendingDirective == "test-xfail"));

            pendingDirective = null;
            blankSinceDirective = false;
        }

        return result;
    }

    /// <summary>Discovers public <c>'@test</c> procedures in standard modules.</summary>
    public static IReadOnlyList<(string Module, VbaProcedureInfo Procedure)> DiscoverTests(object project)
    {
        var tests = new List<(string, VbaProcedureInfo)>();

        foreach (var component in ListComponents(project))
        {
            if (!component.IsStandardModule)
            {
                continue;
            }

            string code;
            try
            {
                code = ReadModuleCode(project, component.Name);
            }
            catch
            {
                continue;
            }

            foreach (var procedure in FindProcedures(code).Where(p => p.IsPublic && p.IsTest))
            {
                tests.Add((component.Name, procedure));
            }
        }

        return tests;
    }

    public static string GetTypeName(int type) => type switch
    {
        StandardModule => "StandardModule",
        ClassModule => "ClassModule",
        FormModule => "FormModule",
        DocumentModule => "DocumentModule",
        _ => $"Type{type}"
    };

    /// <summary>
    /// Reads a CodeModule's text. CodeModule.Lines returns \r\n line breaks.
    /// </summary>
    public static string ReadCode(object codeModule)
    {
        var count = Convert.ToInt32(AccessComAutomation.GetProperty(codeModule, "CountOfLines"), CultureInfo.InvariantCulture);
        return count > 0
            ? Convert.ToString(AccessComAutomation.GetProperty(codeModule, "Lines", 1, count), CultureInfo.InvariantCulture) ?? string.Empty
            : string.Empty;
    }

    /// <summary>
    /// Replaces a module's text as a minimal line-level edit. Never empties the module first and
    /// never uses AddFromString. See the reliability note in AccessTools.WriteModuleCode.
    /// </summary>
    public static void WriteCode(object codeModule, string code)
    {
        var newLines = string.IsNullOrEmpty(code) ? [] : code.Replace("\r\n", "\n").Split('\n');

        var oldCount = Convert.ToInt32(AccessComAutomation.GetProperty(codeModule, "CountOfLines"), CultureInfo.InvariantCulture);
        var oldLines = oldCount == 0
            ? []
            : (Convert.ToString(AccessComAutomation.GetProperty(codeModule, "Lines", 1, oldCount), CultureInfo.InvariantCulture) ?? string.Empty)
                .Replace("\r\n", "\n").Split('\n');

        var common = Math.Min(oldLines.Length, newLines.Length);
        for (var i = 0; i < common; i++)
        {
            if (!string.Equals(oldLines[i], newLines[i], StringComparison.Ordinal))
            {
                AccessComAutomation.InvokeMethod(codeModule, "ReplaceLine", i + 1, newLines[i]);
            }
        }

        if (newLines.Length > oldLines.Length)
        {
            for (var i = oldLines.Length; i < newLines.Length; i++)
            {
                AccessComAutomation.InvokeMethod(codeModule, "InsertLines", i + 1, newLines[i]);
            }
        }
        else if (oldLines.Length > newLines.Length)
        {
            AccessComAutomation.InvokeMethod(codeModule, "DeleteLines", newLines.Length + 1, oldLines.Length - newLines.Length);
        }

        // Inserting past the old end can leave one phantom empty line after the real last line.
        var count = Convert.ToInt32(AccessComAutomation.GetProperty(codeModule, "CountOfLines"), CultureInfo.InvariantCulture);
        while (count > newLines.Length)
        {
            var last = Convert.ToString(AccessComAutomation.GetProperty(codeModule, "Lines", count, 1), CultureInfo.InvariantCulture) ?? string.Empty;
            if (last.Length != 0)
            {
                break;
            }

            AccessComAutomation.InvokeMethod(codeModule, "DeleteLines", count, 1);
            count--;
        }
    }

    /// <summary>
    /// Writes <paramref name="intendedCode"/> and verifies it by reading the module back. Access's
    /// automation normalizes code on the way in (it trims trailing whitespace, re-cases identifiers and
    /// splits over-long lines), so the comparison is semantic: trailing whitespace is ignored and
    /// identifier casing alone counts as a match. Anything else is a failure, and the exception says
    /// exactly which line differs and how, then the previous content is restored.
    /// </summary>
    private static CodeComparison WriteAndVerify(object codeModule, string previousCode, string intendedCode)
    {
        CodeComparison? comparison = null;
        Exception? writeFailure = null;
        try
        {
            WriteCode(codeModule, intendedCode);
            comparison = CompareCode(intendedCode, ReadCode(codeModule));
            if (!comparison.Same)
            {
                writeFailure = new InvalidOperationException(
                    "write completed but content did not verify (Access's automation altered it) - "
                    + DescribeMismatch(comparison));
            }
        }
        catch (Exception ex)
        {
            writeFailure = ex;
        }

        if (writeFailure is null)
        {
            return comparison!;
        }

        var rolledBack = false;
        try
        {
            WriteCode(codeModule, previousCode);
            rolledBack = CompareCode(previousCode, ReadCode(codeModule)).Same;
        }
        catch
        {
            rolledBack = false;
        }

        throw new InvalidOperationException(
            rolledBack
                ? $"Write failed ({writeFailure.Message}) and was rolled back to the previous content."
                : $"Write failed ({writeFailure.Message}), AND rollback did not verify. The module may be inconsistent.",
            writeFailure);
    }

    /// <summary>
    /// Compares intended source with what <c>CodeModule</c> actually holds. Line endings, trailing
    /// whitespace and trailing blank lines are ignored, and a line that differs only by letter case
    /// matches (VBA is case-insensitive and the VBE re-cases identifiers to their declarations).
    /// </summary>
    internal static CodeComparison CompareCode(string intended, string actual)
    {
        var expected = TrimTrailingBlankLines(SplitLines(intended));
        var got = TrimTrailingBlankLines(SplitLines(actual));

        var caseDifferences = 0;
        var lineCount = Math.Max(expected.Count, got.Count);
        for (var i = 0; i < lineCount; i++)
        {
            var wanted = i < expected.Count ? expected[i] : null;
            var found = i < got.Count ? got[i] : null;

            if (wanted is null || found is null)
            {
                return new CodeComparison(false, false, caseDifferences, i + 1, wanted, found, expected.Count, got.Count);
            }

            if (string.Equals(wanted.TrimEnd(), found.TrimEnd(), StringComparison.Ordinal))
            {
                continue;
            }

            if (string.Equals(wanted.TrimEnd(), found.TrimEnd(), StringComparison.OrdinalIgnoreCase))
            {
                caseDifferences++;
                continue;
            }

            return new CodeComparison(false, false, caseDifferences, i + 1, wanted, found, expected.Count, got.Count);
        }

        return new CodeComparison(true, caseDifferences > 0, caseDifferences, null, null, null, expected.Count, got.Count);
    }

    /// <summary>A one-line, quotable account of the first difference, for error messages.</summary>
    internal static string DescribeMismatch(CodeComparison comparison)
    {
        if (comparison.Line is not int line)
        {
            return "the read-back differed but no line could be singled out";
        }

        return $"line {line} differs: expected {Quote(comparison.Expected)}, module has {Quote(comparison.Actual)}"
            + $" (intended {comparison.ExpectedLines} lines, module now has {comparison.ActualLines}).";
    }

    private static string Quote(string? text)
        => text is null
            ? "<no line>"
            : "'" + (text.Length <= 120 ? text : text[..120] + "...") + "'";

    /// <summary>
    /// Builds the body that is safe to write through CodeModule: strips export-only headers
    /// (Attribute VB_Name, VERSION/BEGIN blocks) and keeps a single Option block at the top, using
    /// the caller's Option lines when supplied and the existing module's otherwise.
    /// </summary>
    internal static string NormalizeForWrite(string code, string existingCode)
        => PrepareForWrite(code, existingCode).Code;

    /// <summary>
    /// <see cref="NormalizeForWrite"/> plus the normalizations the VBE applies itself, so the write is
    /// idempotent and verifies first time: tabs are expanded to space tab stops (a tab passed to
    /// InsertLines/ReplaceLine is treated as a line break, not as indentation) and trailing whitespace
    /// is dropped from code lines (Access drops it anyway; whitespace-only lines are kept as-is).
    /// </summary>
    internal static PreparedWrite PrepareForWrite(string code, string existingCode)
    {
        var expanded = PrepareCodeLines(code);
        var existingOptions = ExtractOptionLines(PrepareCodeLines(existingCode).Code);
        var inputOptions = ExtractOptionLines(expanded.Code);
        var body = StripExportHeadersAndOptions(expanded.Lines);

        var options = inputOptions.Count > 0 ? inputOptions : existingOptions;
        var builder = new StringBuilder();
        foreach (var option in options)
        {
            builder.AppendLine(option);
        }

        if (options.Count > 0 && body.Count > 0 && body[0].Trim().Length != 0)
        {
            builder.AppendLine();
        }

        foreach (var line in body)
        {
            builder.AppendLine(line);
        }

        return new PreparedWrite(
            builder.ToString().TrimEnd('\r', '\n'),
            expanded.ExpandedTabs,
            expanded.TrimmedLines);
    }

    /// <summary>Expands tabs and trims trailing whitespace from lines that have content.</summary>
    internal static (string Code, string[] Lines, int ExpandedTabs, int TrimmedLines) PrepareCodeLines(string code)
    {
        var raw = SplitLines(code);
        var lines = new string[raw.Length];
        var expandedTabs = 0;
        var trimmedLines = 0;

        for (var i = 0; i < raw.Length; i++)
        {
            var line = raw[i];

            if (line.Contains('\t', StringComparison.Ordinal))
            {
                expandedTabs += line.Count(static c => c == '\t');
                line = ExpandTabs(line);
            }

            if (line.Trim().Length != 0)
            {
                var withoutTrailing = line.TrimEnd();
                if (withoutTrailing.Length != line.Length)
                {
                    trimmedLines++;
                    line = withoutTrailing;
                }
            }

            lines[i] = line;
        }

        return (string.Join('\n', lines), lines, expandedTabs, trimmedLines);
    }

    /// <summary>Advances a tab to the next stop, the way the VBE editor does, so columns line up.</summary>
    private static string ExpandTabs(string line)
    {
        var builder = new StringBuilder(line.Length + 8);
        var column = 0;

        foreach (var character in line)
        {
            if (character == '\t')
            {
                var spaces = TabWidth - (column % TabWidth);
                builder.Append(' ', spaces);
                column += spaces;
            }
            else
            {
                builder.Append(character);
                column++;
            }
        }

        return builder.ToString();
    }

    /// <summary>
    /// Refuses source that <c>CodeModule</c> cannot hold on one line, naming the line and its length
    /// instead of letting Access split the text and then reporting an anonymous verify mismatch.
    /// </summary>
    internal static void ValidateLineLengths(string code, string moduleName)
    {
        var lines = SplitLines(code);
        for (var i = 0; i < lines.Length; i++)
        {
            if (lines[i].Length > MaxCodeModuleLineLength)
            {
                throw new ModuleWriteRefusedException(
                    $"Line {i + 1} of module '{moduleName}' is {lines[i].Length} characters long. Access's CodeModule " +
                    $"keeps at most {MaxCodeModuleLineLength} characters per line and silently splits anything longer, " +
                    "so the write was refused. Shorten that line, or use access_vba_import_file, which imports through " +
                    "VBComponents.Import and preserves long lines.");
            }
        }
    }

    /// <summary>
    /// Saves the VBA project through Access's only headless save path: compile-and-save. It is also
    /// what registers a new component as a database module object. Returns whether the project
    /// compiled, which is also whether the save completed - Access does not save a project it cannot
    /// compile, and an unsaved project is discarded (or answered with a modal Save As dialog) on close.
    /// </summary>
    public static bool CompileAndSave(object appObject)
    {
        dynamic app = appObject;
        dynamic? doCmd = null;
        try
        {
            doCmd = app.DoCmd;
            doCmd.RunCommand(CompileAndSaveAllModulesCommand);
            return Convert.ToBoolean(app.IsCompiled, CultureInfo.InvariantCulture);
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

    private static List<string> TrimTrailingBlankLines(string[] lines)
    {
        var list = new List<string>(lines);
        while (list.Count > 0 && list[^1].Trim().Length == 0)
        {
            list.RemoveAt(list.Count - 1);
        }

        return list;
    }


    private static List<string> StripExportHeadersAndOptions(string[] lines)
    {
        var result = new List<string>();
        var inClassHeader = false;

        foreach (var line in lines)
        {
            var trimmed = line.Trim();

            if (inClassHeader)
            {
                if (trimmed.Equals("END", StringComparison.OrdinalIgnoreCase))
                {
                    inClassHeader = false;
                }

                continue;
            }

            if (VersionRegex.IsMatch(line))
            {
                inClassHeader = true;
                continue;
            }

            if (AttributeRegex.IsMatch(line) || OptionRegex.IsMatch(line))
            {
                continue;
            }

            // Skip leading blank lines; keep internal blank lines.
            if (result.Count == 0 && trimmed.Length == 0)
            {
                continue;
            }

            result.Add(line);
        }

        while (result.Count > 0 && result[^1].Trim().Length == 0)
        {
            result.RemoveAt(result.Count - 1);
        }

        return result;
    }

    private static List<string> ExtractOptionLines(string code)
    {
        var options = new List<string>();
        foreach (var line in SplitLines(code))
        {
            var trimmed = line.Trim();
            if (trimmed.Length == 0)
            {
                continue;
            }

            if (OptionRegex.IsMatch(line))
            {
                if (!options.Contains(trimmed, StringComparer.OrdinalIgnoreCase))
                {
                    options.Add(trimmed);
                }

                continue;
            }

            if (trimmed.StartsWith("'", StringComparison.Ordinal) || AttributeRegex.IsMatch(line))
            {
                continue;
            }

            break;
        }

        return options;
    }

    private static void ValidateModuleName(string moduleName)
    {
        if (!Regex.IsMatch(moduleName, @"^[A-Za-z_][A-Za-z0-9_]{0,30}$"))
        {
            throw new ArgumentException(
                "Module name must be a valid VBA identifier (letter/underscore, then letters/digits/underscore, at most 31 characters).",
                nameof(moduleName));
        }
    }

    private static object? FindComponent(object components, string moduleName)
    {
        var count = AccessComAutomation.GetCount(components!);
        for (var i = 1; i <= count; i++)
        {
            var component = AccessComAutomation.GetItem(components!, i)!;
            var name = Convert.ToString(AccessComAutomation.GetProperty(component, "Name"), CultureInfo.InvariantCulture) ?? string.Empty;
            if (name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
            {
                return component;
            }

            AccessComAutomation.ReleaseComObject(component);
        }

        return null;
    }

    private static string DefaultBackupDirectory(string databasePath)
        => Path.Combine(Path.GetDirectoryName(databasePath) ?? ".", "_accessmcp_backups");

    private static string BackupModuleCode(string directory, string moduleName, string code)
    {
        Directory.CreateDirectory(directory);
        var safeName = new string(moduleName.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c).ToArray());
        var timestamp = DateTime.UtcNow.ToString("yyyyMMdd_HHmmssfff", CultureInfo.InvariantCulture);
        var path = Path.Combine(directory, $"{safeName}_{timestamp}.bas");
        File.WriteAllText(path, code);
        return path;
    }

    internal static string Sha256Hex(string text)
    {
        var normalized = NormalizeForHash(text);
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(normalized));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }

    internal static string NormalizeForHash(string text)
    {
        var lines = SplitLines(text).Select(l => l.TrimEnd()).ToArray();
        return string.Join("\n", lines).TrimEnd('\n');
    }

    private static int CountLines(string text) => SplitLines(text).Length;

    private static string[] SplitLines(string text) => text.Replace("\r\n", "\n").Split('\n');
}
