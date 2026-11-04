using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>
/// MCP tools for compiling, running and regression-testing VBA inside a real Access process,
/// including compiled .mde/.accde databases. Execution tools are gated behind
/// <see cref="AccessMcpOptions.EnableVbaTestTools"/> because they run arbitrary code.
/// </summary>
[McpServerToolType]
public sealed class AccessVbaTools(AccessMcpOptions options, ILoggerFactory loggerFactory)
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };

    private readonly AccessMcpOptions _options = options;
    private readonly ILogger _logger = loggerFactory.CreateLogger("AccessMCP.AccessVbaTools");
    private readonly VbaTestHarness _harness = new(options, loggerFactory);

    [McpServerTool(Name = "access_vba_doctor", Title = "Access: VBA Doctor", ReadOnly = true, Idempotent = true)]
    [Description("Reports what VBA automation will and will not work: Access version/bitness/runtime, VBE trust, ACE provider bitness, scratch/cache directories, MDE capability and the PIA status. Works without COM automation enabled.")]
    public async Task<string> DoctorAsync(
        [Description("Optional database to include in the report. If omitted, uses DefaultDatabasePath when it exists.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        var info = await VbaDoctor.ProbeAsync(_options, databasePath, cancellationToken);
        return JsonSerializer.Serialize(info, JsonOptions);
    }

    [McpServerTool(Name = "access_vba_ensure_module", Title = "Access: VBA Ensure Module", ReadOnly = false, Idempotent = true)]
    [Description("Creates a standard or class module if it is missing and replaces its code with a verified line-level write. Backs up the previous code and rolls back on any read-back mismatch. Never uses AddFromString. A newly created module is registered with the database so the write reaches the file (compile=false still performs that one save). Tabs are expanded to spaces and over-long lines are refused with the line number. Compiles afterwards by default.")]
    public async Task<string> EnsureModuleAsync(
        [Description("Module name (a valid VBA identifier).")] string moduleName,
        [Description("Module code. Do not include Attribute VB_Name; export-only headers are stripped.")] string? code = null,
        [Description("File to read the code from (.bas/.cls). Mutually exclusive with code.")] string? codeFile = null,
        [Description("'standard' or 'class'. Ignored when the module already exists.")] string kind = "standard",
        [Description("Compile and save all modules after writing. Default true.")] bool compile = true,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        [Description("Timeout in seconds. Defaults to VbaTimeoutSeconds.")] int? timeoutSeconds = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            RequireComAutomation();
            RequireWritable();
            ArgumentException.ThrowIfNullOrWhiteSpace(moduleName);

            if (code is null && codeFile is null)
            {
                throw new ArgumentException("Either code or codeFile is required.");
            }

            if (code is not null && codeFile is not null)
            {
                throw new ArgumentException("Specify code or codeFile, not both.");
            }

            var moduleCode = code ?? ReadCodeFile(codeFile!);

            // Refuse over-long lines before Access is even started: the CodeModule write would split
            // them and the caller would get an anonymous verify mismatch instead of the real reason.
            VbaProjectEditor.ValidateLineLengths(moduleCode, moduleName);

            var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
            var com = new AccessComAutomation(_options);

            var outcome = await com.RunInAccessWithWatchdogAsync(
                dbPath,
                app =>
                {
                    dynamic dynamicApp = app;
                    object? vbe = null;
                    object? project = null;
                    try
                    {
                        vbe = dynamicApp.VBE;
                        project = VbaProjectEditor.GetProject(vbe!, null);
                        var write = VbaProjectEditor.EnsureModule(
                            project,
                            dbPath,
                            moduleName,
                            kind,
                            moduleCode,
                            createIfMissing: true,
                            allowOverwrite: true,
                            backupDirectory: null,
                            accessApp: (object)app);

                        var compileResult = compile ? VbaCompiler.Compile((object)app) : null;

                        // A new component is only in the file once the database has registered it, which
                        // happens on a successful compile-and-save. Existing components are saved when
                        // the database closes, so those are already durable at this point.
                        var saved = compileResult is not null
                            ? compileResult.IsCompiled
                            : !write.Created || write.Registered;

                        if (!saved)
                        {
                            saved = VbaProjectEditor.CompileAndSave((object)app);
                        }

                        return (Write: write, Compile: compileResult, Saved: saved);
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(project);
                        AccessComAutomation.ReleaseComObject(vbe);
                    }
                },
                new AccessRunOptions { Timeout = Timeout(timeoutSeconds) },
                cancellationToken).ConfigureAwait(false);

            if (!outcome.Ok)
            {
                return Failure(outcome);
            }

            var writeResult = outcome.Value.Write;
            var savedResult = outcome.Value.Saved;
            return JsonSerializer.Serialize(new
            {
                ok = savedResult && outcome.Value.Compile?.Ok != false,
                databasePath = dbPath,
                moduleName,
                kind = kind.ToLowerInvariant() is "class" or "cls" ? "class" : "standard",
                created = writeResult.Created,
                saved = savedResult,
                lines = writeResult.Lines,
                sha256 = writeResult.Sha256,
                backupPath = writeResult.BackupPath,
                registered = writeResult.Created ? writeResult.Registered : (bool?)null,
                caseNormalized = writeResult.CaseNormalized ? true : (bool?)null,
                caseNormalizedLines = writeResult.CaseNormalizedLines > 0 ? writeResult.CaseNormalizedLines : (int?)null,
                expandedTabs = writeResult.ExpandedTabs > 0 ? writeResult.ExpandedTabs : (int?)null,
                trimmedTrailingWhitespaceLines = writeResult.TrimmedLines > 0 ? writeResult.TrimmedLines : (int?)null,
                error = savedResult
                    ? null
                    : "The module was written in Access, but Access did not save it: the VBA project could not be "
                      + "compiled, and Access cannot persist a new module until it compiles. Fix the compile error and retry.",
                compile = outcome.Value.Compile
            }, JsonOptions);
        }
        catch (Exception ex)
        {
            return Error(ex);
        }
    }

    [McpServerTool(Name = "access_vba_remove_module", Title = "Access: VBA Remove Module", ReadOnly = false, Idempotent = false)]
    [Description("Removes a standard or class module. Form/report code-behind is always refused. By default only modules whose names start with 'AccessMcp' may be removed.")]
    public async Task<string> RemoveModuleAsync(
        [Description("Module name.")] string moduleName,
        [Description("Only allow modules created by the harness (name starts with AccessMcp). Default true.")] bool onlyHarnessModules = true,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            RequireComAutomation();
            RequireWritable();
            ArgumentException.ThrowIfNullOrWhiteSpace(moduleName);

            var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
            var com = new AccessComAutomation(_options);

            var outcome = await com.RunInAccessWithWatchdogAsync(
                dbPath,
                app =>
                {
                    dynamic dynamicApp = app;
                    object? vbe = null;
                    object? project = null;
                    try
                    {
                        vbe = dynamicApp.VBE;
                        project = VbaProjectEditor.GetProject(vbe!, null);
                        var removed = VbaProjectEditor.RemoveModule(project, moduleName, onlyHarnessModules);

                        // Save explicitly: a removal is only durable once the project is written back.
                        var saved = VbaProjectEditor.CompileAndSave((object)app);
                        return (Removed: removed, Saved: saved);
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(project);
                        AccessComAutomation.ReleaseComObject(vbe);
                    }
                },
                new AccessRunOptions { Timeout = Timeout(null) },
                cancellationToken).ConfigureAwait(false);

            return outcome.Ok
                ? JsonSerializer.Serialize(new
                {
                    ok = true,
                    databasePath = dbPath,
                    moduleName,
                    removed = outcome.Value.Removed
                }, JsonOptions)
                : Failure(outcome);
        }
        catch (Exception ex)
        {
            return Error(ex);
        }
    }

    [McpServerTool(Name = "access_vba_import_file", Title = "Access: VBA Import File", ReadOnly = false, Idempotent = false)]
    [Description("Imports a .bas or .cls file with VBComponents.Import, preserving its Attribute VB_Name and class headers. Forms need a designer and are rejected. The import is saved to the database (compile=false still performs that one save), and imported files may contain tabs and lines longer than CodeModule's 1023-character limit.")]
    public async Task<string> ImportFileAsync(
        [Description("Path to the .bas or .cls file.")] string filePath,
        [Description("Compile and save all modules after import. Default true.")] bool compile = true,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            RequireComAutomation();
            RequireWritable();
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

            var fullFile = Path.GetFullPath(filePath);
            var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
            var com = new AccessComAutomation(_options);

            var outcome = await com.RunInAccessWithWatchdogAsync(
                dbPath,
                app =>
                {
                    dynamic dynamicApp = app;
                    object? vbe = null;
                    object? project = null;
                    try
                    {
                        vbe = dynamicApp.VBE;
                        project = VbaProjectEditor.GetProject(vbe!, null);
                        var imported = VbaProjectEditor.ImportFile(project, fullFile);
                        var compileResult = compile ? VbaCompiler.Compile((object)app) : null;

                        // VBComponents.Import creates an unregistered component too, and an unregistered
                        // component is discarded when the database closes, so the import must be saved
                        // even when the caller did not ask for a compile.
                        var saved = compileResult is not null
                            ? compileResult.IsCompiled
                            : VbaProjectEditor.CompileAndSave((object)app);

                        return (Import: imported, Compile: compileResult, Saved: saved);
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(project);
                        AccessComAutomation.ReleaseComObject(vbe);
                    }
                },
                new AccessRunOptions { Timeout = Timeout(null) },
                cancellationToken).ConfigureAwait(false);

            if (!outcome.Ok)
            {
                return Failure(outcome);
            }

            return JsonSerializer.Serialize(new
            {
                ok = outcome.Value.Saved && outcome.Value.Compile?.Ok != false,
                databasePath = dbPath,
                filePath = fullFile,
                moduleName = outcome.Value.Import.ModuleName,
                saved = outcome.Value.Saved,
                lines = outcome.Value.Import.Lines,
                sha256 = outcome.Value.Import.Sha256,
                error = outcome.Value.Saved
                    ? null
                    : "The module was imported in Access, but Access did not save it: the VBA project could not be "
                      + "compiled, and Access cannot persist an imported module until it compiles.",
                compile = outcome.Value.Compile
            }, JsonOptions);
        }
        catch (Exception ex)
        {
            return Error(ex);
        }
    }

    [McpServerTool(Name = "access_vba_export_file", Title = "Access: VBA Export File", ReadOnly = true, Idempotent = false)]
    [Description("Exports a module with VBComponent.Export, preserving headers. Writes the file to disk; the database is not modified.")]
    public async Task<string> ExportFileAsync(
        [Description("Module name.")] string moduleName,
        [Description("Destination file path.")] string filePath,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        CancellationToken cancellationToken = default)
    {
        RequireComAutomation();
        ArgumentException.ThrowIfNullOrWhiteSpace(moduleName);
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var fullFile = Path.GetFullPath(filePath);
        var com = new AccessComAutomation(_options);

        var outcome = await com.RunInAccessWithWatchdogAsync(
            dbPath,
            app =>
            {
                dynamic dynamicApp = app;
                object? vbe = null;
                object? project = null;
                try
                {
                    vbe = dynamicApp.VBE;
                    project = VbaProjectEditor.GetProject(vbe!, null);
                    return VbaProjectEditor.ExportFile(project, moduleName, fullFile);
                }
                finally
                {
                    AccessComAutomation.ReleaseComObject(project);
                    AccessComAutomation.ReleaseComObject(vbe);
                }
            },
            new AccessRunOptions { Timeout = Timeout(null) },
            cancellationToken).ConfigureAwait(false);

        return outcome.Ok
            ? JsonSerializer.Serialize(new
            {
                ok = true,
                databasePath = dbPath,
                moduleName,
                filePath = outcome.Value.FilePath,
                lines = outcome.Value.Lines,
                sha256 = outcome.Value.Sha256
            }, JsonOptions)
            : Failure(outcome);
    }

    [McpServerTool(Name = "access_vba_compile", Title = "Access: VBA Compile", ReadOnly = false, Idempotent = false)]
    [Description("Compiles and saves all VBA modules and reports the real result via Application.IsCompiled, with broken-reference preflight and best-effort module/line diagnostics. Takes a per-call timeout, unlike access_compile_vba.")]
    public async Task<string> CompileAsync(
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        [Description("Timeout in seconds. Defaults to VbaCompileTimeoutSeconds.")] int? timeoutSeconds = null,
        CancellationToken cancellationToken = default)
    {
        RequireComAutomation();
        RequireWritable();

        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var com = new AccessComAutomation(_options);
        var timeout = timeoutSeconds ?? _options.VbaCompileTimeoutSeconds;

        var outcome = await com.RunInAccessWithWatchdogAsync(
            dbPath,
            app => VbaCompiler.Compile((object)app),
            new AccessRunOptions { Timeout = TimeSpan.FromSeconds(Math.Max(1, timeout)) },
            cancellationToken).ConfigureAwait(false);

        if (!outcome.Ok)
        {
            return Failure(outcome);
        }

        var result = outcome.Value!;
        return JsonSerializer.Serialize(new
        {
            ok = result.Ok,
            isCompiled = result.IsCompiled,
            module = result.Module,
            line = result.Line,
            message = result.Message,
            errorNumber = result.ErrorNumber,
            project = result.Project,
            dialogText = outcome.DialogText ?? result.DialogText,
            brokenReferenceFlag = result.BrokenReferenceFlag,
            brokenReferences = result.BrokenReferences
        }, JsonOptions);
    }

    [McpServerTool(Name = "access_vba_run", Title = "Access: VBA Run", ReadOnly = false, Idempotent = false)]
    [Description("Runs a public VBA procedure and returns a structured result. By default it runs in a scratch copy; set inPlace=true to run against the database itself, so its changes persist. Access's Application.Run takes unqualified procedure names; a 'Module.Procedure' entry point is accepted and the module part is used for validation. Requires EnableVbaTestTools.")]
    public async Task<string> RunAsync(
        [Description("Public procedure name, optionally 'Module.Procedure'.")] string procedure,
        [Description("Optional JSON array of arguments.")] JsonElement? args = null,
        [Description("Timeout in seconds. Defaults to VbaTimeoutSeconds.")] int? timeoutSeconds = null,
        [Description("Path to .mdb/.accdb/.mde/.accde. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        [Description("Run against the database itself instead of a scratch copy. Changes persist, nothing is rolled back. Default false.")] bool inPlace = false,
        CancellationToken cancellationToken = default)
    {
        RequireVbaExecution();
        ArgumentException.ThrowIfNullOrWhiteSpace(procedure);

        var dbPath = AccessDbPathPolicy.NormalizeComDatabasePath(databasePath, _options);
        var run = await _harness.RunProcedureAsync(dbPath, procedure, ParseArguments(args), timeoutSeconds, cancellationToken, inPlace);
        return JsonSerializer.Serialize(run, JsonOptions);
    }

    [McpServerTool(Name = "access_vba_eval", Title = "Access: VBA Eval", ReadOnly = false, Idempotent = false)]
    [Description("Evaluates a VBA expression with Application.Eval. By default in a scratch copy; set inPlace=true to evaluate against the database itself. Requires EnableVbaTestTools.")]
    public async Task<string> EvalAsync(
        [Description("VBA expression, for example 2*21 or SomeFunction().")] string expression,
        [Description("Timeout in seconds. Defaults to VbaTimeoutSeconds.")] int? timeoutSeconds = null,
        [Description("Path to .mdb/.accdb/.mde/.accde. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        [Description("Evaluate against the database itself instead of a scratch copy. Default false.")] bool inPlace = false,
        CancellationToken cancellationToken = default)
    {
        RequireVbaExecution();
        ArgumentException.ThrowIfNullOrWhiteSpace(expression);

        var dbPath = AccessDbPathPolicy.NormalizeComDatabasePath(databasePath, _options);
        var run = await _harness.EvalAsync(dbPath, expression, timeoutSeconds, cancellationToken, inPlace);
        return JsonSerializer.Serialize(run, JsonOptions);
    }

    [McpServerTool(Name = "access_vba_run_tests", Title = "Access: VBA Run Tests", ReadOnly = false, Idempotent = false)]
    [Description("Runs public '@test procedures in an .mdb/.accdb: discovers them through the VBE, injects the shared VbaTestKit and a generated runner, compiles, and returns pass/fail/skip counts plus per-test messages. By default this happens on a scratch copy; set inPlace=true to run in the database itself (the injected modules are removed afterwards). Requires EnableVbaTestTools.")]
    public async Task<string> RunTestsAsync(
        [Description("Optional test names to run; a name may be 'Procedure' or 'Module.Procedure'.")] string[]? tests = null,
        [Description("Timeout in seconds. Defaults to VbaTimeoutSeconds.")] int? timeoutSeconds = null,
        [Description("Keep the scratch workspace (databases, result NDJSON, log).")] bool keepArtifacts = false,
        [Description("Path to .mdb/.accdb. If omitted, uses DefaultDatabasePath.")] string? databasePath = null,
        [Description("Run against the database itself instead of a scratch copy. The harness removes its own injected modules afterwards. Default false.")] bool inPlace = false,
        CancellationToken cancellationToken = default)
    {
        RequireVbaExecution();
        var dbPath = AccessDbPathPolicy.NormalizeAndValidate(databasePath, _options);
        var result = await _harness.RunTestsAsync(dbPath, tests, timeoutSeconds, keepArtifacts, cancellationToken, inPlace);

        return JsonSerializer.Serialize(new
        {
            ok = result.Tests.Ok && result.Run.Ok,
            total = result.Tests.Total,
            passed = result.Tests.Passed,
            failed = result.Tests.Failed,
            skipped = result.Tests.Skipped,
            xfailed = result.Tests.XFailed,
            xpassed = result.Tests.XPassed,
            tests = result.Tests.Tests,
            logTail = result.Tests.LogTail,
            error = result.Tests.Error ?? result.Run.Error?.Description,
            timedOut = result.Run.TimedOut,
            crashed = result.Run.Crashed,
            exitCode = result.Run.ExitCode,
            elapsedMs = result.Run.ElapsedMs,
            dialogText = result.Run.DialogText,
            artifactsKept = result.Run.ArtifactsKept,
            artifactDirectory = result.Run.ArtifactDirectory
        }, JsonOptions);
    }

    [McpServerTool(Name = "access_vba_build_mde", Title = "Access: VBA Build MDE", ReadOnly = false, Idempotent = false)]
    [Description("Builds a compiled .mde/.accde from an .mdb/.accdb. By default on a scratch copy; set inPlace=true to compile the database itself. Uses the headless early-bound SysCmd(603) route after compile and close. 'prebuilt' copies an existing compiled database instead. Requires EnableVbaTestTools.")]
    public async Task<string> BuildMdeAsync(
        [Description("Source .mdb/.accdb (or .mde/.accde for the prebuilt strategy).")] string databasePath,
        [Description("Output path. Defaults to the source name with .mde/.accde beside the source.")] string? outputPath = null,
        [Description("'syscmd' or 'prebuilt'. Defaults to VbaMdeBuildStrategy.")] string? strategy = null,
        [Description("Run MSACCESS.EXE /compact /decompile on the copy before compiling. Defaults to VbaDecompileFirst.")] bool? decompileFirst = null,
        [Description("Timeout in seconds. Defaults to VbaTimeoutSeconds.")] int? timeoutSeconds = null,
        [Description("Build from the database itself instead of a scratch copy (it is compiled in the process). Default false.")] bool inPlace = false,
        CancellationToken cancellationToken = default)
    {
        RequireVbaExecution();
        RequireWritable();
        ArgumentException.ThrowIfNullOrWhiteSpace(databasePath);

        var result = await _harness.BuildMdeAsync(
            databasePath,
            outputPath,
            strategy ?? _options.VbaMdeBuildStrategy,
            decompileFirst ?? _options.VbaDecompileFirst,
            timeoutSeconds,
            cancellationToken,
            inPlace);

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool(Name = "access_vba_cache_info", Title = "Access: VBA Cache Info", ReadOnly = true, Idempotent = true)]
    [Description("Lists cached built MDEs: key, path, size, creation time, total size and the configured cap.")]
    public Task<string> CacheInfoAsync(CancellationToken cancellationToken = default)
    {
        var cache = new VbaMdeCache(_options);
        return Task.FromResult(JsonSerializer.Serialize(cache.Info(), JsonOptions));
    }

    [McpServerTool(Name = "access_vba_cache_clear", Title = "Access: VBA Cache Clear", ReadOnly = false, Idempotent = false)]
    [Description("Removes one cached MDE by key, or every entry when all=true.")]
    public Task<string> CacheClearAsync(
        [Description("Cache key to remove.")] string? key = null,
        [Description("Remove every cached MDE.")] bool all = false,
        CancellationToken cancellationToken = default)
    {
        if (!all && string.IsNullOrWhiteSpace(key))
        {
            throw new ArgumentException("Pass key, or all=true.");
        }

        var cache = new VbaMdeCache(_options);
        var (removed, freed) = cache.Clear(key, all);
        return Task.FromResult(JsonSerializer.Serialize(new { ok = true, removed, freedBytes = freed }, JsonOptions));
    }

    [McpServerTool(Name = "access_vba_test", Title = "Access: VBA Test Scenario", ReadOnly = false, Idempotent = false)]
    [Description("End-to-end VBA scenario: copy the database, inject modules, compile, optionally build an MDE with SysCmd(603), run an entry point or the generated test runner, and return structured prepare/run/test results. By default the source database is never modified; set inPlace=true to work on it directly (modules you supply persist, and the harness removes its own injected kit/runner afterwards). Requires EnableVbaTestTools.")]
    public async Task<string> TestAsync(
        [Description("Source .mdb/.accdb (or .mde/.accde for mode='mde' with mdeStrategy='prebuilt').")] string databasePath,
        [Description("Modules to write: [{ name, kind: 'standard'|'class', source }] or [{ name, kind, filePath }].")] JsonElement? modules = null,
        [Description("Entry point. Optional when '@test procedures are found: the generated runner is used. 'Module.Procedure' is accepted; the module part is used for validation.")] string? entryPoint = null,
        [Description("Optional JSON array of arguments for the entry point.")] JsonElement? args = null,
        [Description("'mdb' or 'mde'. Default 'mdb'.")] string mode = "mdb",
        [Description("'syscmd' or 'prebuilt'. Defaults to VbaMdeBuildStrategy.")] string? mdeStrategy = null,
        [Description("Decompile the scratch copy before compiling. Defaults to VbaDecompileFirst.")] bool? decompileFirst = null,
        [Description("Use the built-MDE cache. Default true.")] bool useCache = true,
        [Description("Force a cache miss for this call.")] bool rebuild = false,
        [Description("MDB only: run even when the compile gate fails. Access can still run an unbroken procedure.")] bool continueOnCompileError = false,
        [Description("Keep the scratch workspace.")] bool keepArtifacts = false,
        [Description("Timeout in seconds. Defaults to VbaTimeoutSeconds.")] int? timeoutSeconds = null,
        [Description("Run against the database itself instead of a scratch copy. Modules you inject persist; the harness removes its own test kit/runner afterwards. Default false.")] bool inPlace = false,
        CancellationToken cancellationToken = default)
    {
        RequireVbaExecution();
        ArgumentException.ThrowIfNullOrWhiteSpace(databasePath);

        var request = new VbaScenarioRequest(
            databasePath,
            ParseModules(modules),
            entryPoint,
            ParseArguments(args),
            mode,
            mdeStrategy ?? _options.VbaMdeBuildStrategy,
            decompileFirst ?? _options.VbaDecompileFirst,
            useCache,
            rebuild,
            continueOnCompileError,
            keepArtifacts,
            timeoutSeconds ?? _options.VbaTimeoutSeconds,
            inPlace);

        var result = await _harness.RunScenarioAsync(request, cancellationToken);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static IReadOnlyList<VbaModuleSpec> ParseModules(JsonElement? modules)
    {
        if (modules is null || modules.Value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
        {
            return [];
        }

        if (modules.Value.ValueKind != JsonValueKind.Array)
        {
            throw new ArgumentException("modules must be a JSON array.");
        }

        var result = new List<VbaModuleSpec>();
        foreach (var element in modules.Value.EnumerateArray())
        {
            if (element.ValueKind != JsonValueKind.Object)
            {
                throw new ArgumentException("Each module must be an object.");
            }

            var name = element.TryGetProperty("name", out var nameElement) ? nameElement.GetString() : null;
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("Each module needs a name.");
            }

            var kind = element.TryGetProperty("kind", out var kindElement) ? kindElement.GetString() ?? "standard" : "standard";
            var source = element.TryGetProperty("source", out var sourceElement) ? sourceElement.GetString() : null;
            var filePath = element.TryGetProperty("filePath", out var fileElement) ? fileElement.GetString() : null;

            if (source is null && filePath is null)
            {
                throw new ArgumentException($"Module '{name}' needs source or filePath.");
            }

            result.Add(new VbaModuleSpec(name, kind, source, filePath));
        }

        return result;
    }

    private static IReadOnlyList<object?> ParseArguments(JsonElement? args)
    {
        if (args is null || args.Value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
        {
            return [];
        }

        if (args.Value.ValueKind != JsonValueKind.Array)
        {
            throw new ArgumentException("args must be a JSON array.");
        }

        return args.Value.EnumerateArray().Select(ParseArgument).ToArray();
    }

    private static object? ParseArgument(JsonElement element) => element.ValueKind switch
    {
        JsonValueKind.Null or JsonValueKind.Undefined => null,
        JsonValueKind.String => element.GetString(),
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Number when element.TryGetInt32(out var intValue) => intValue,
        JsonValueKind.Number when element.TryGetInt64(out var longValue) => longValue,
        JsonValueKind.Number => element.GetDouble(),
        JsonValueKind.Array => element.EnumerateArray().Select(ParseArgument).ToArray(),
        _ => throw new ArgumentException("Only scalar JSON values and arrays can be passed to VBA.")
    };

    private static string Failure<T>(AccessRunOutcome<T> outcome)
    {
        var kind = outcome.TimedOut ? "timedOut" : outcome.Crashed ? "crashed" : outcome.Cancelled ? "cancelled" : "error";
        return JsonSerializer.Serialize(new
        {
            ok = false,
            outcome = kind,
            error = outcome.Exception?.Message,
            exitCode = outcome.ExitCode,
            dialogText = outcome.DialogText
        }, JsonOptions);
    }

    /// <summary>
    /// Turns a tool-level exception into the same structured error shape as a failed Access session.
    /// Without this the MCP client only sees "An error occurred invoking '&lt;tool&gt;'." and none of
    /// the reason - a missing file, a bad module name and a content failure all look identical.
    /// </summary>
    private static string Error(Exception exception)
        => JsonSerializer.Serialize(new
        {
            ok = false,
            outcome = "error",
            error = exception.Message,
            exceptionType = exception.GetType().Name
        }, JsonOptions);

    /// <summary>Reads module source from disk, with an error that names the file that is missing.</summary>
    private static string ReadCodeFile(string codeFile)
    {
        var fullPath = Path.GetFullPath(codeFile);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"Module file not found: {fullPath}", fullPath);
        }

        return File.ReadAllText(fullPath);
    }

    private void RequireComAutomation()
    {
        if (!_options.EnableComAutomation)
        {
            throw new InvalidOperationException("Access COM automation is disabled (EnableComAutomation is false).");
        }
    }

    private void RequireVbaExecution()
    {
        RequireComAutomation();
        if (!_options.EnableVbaTestTools)
        {
            throw new InvalidOperationException(
                "VBA execution is disabled. Set EnableVbaTestTools=true to allow arbitrary VBA to run.");
        }
    }

    private void RequireWritable()
    {
        if (_options.ReadOnly)
        {
            throw new InvalidOperationException("Server is configured as read-only.");
        }
    }

    private TimeSpan Timeout(int? seconds)
        => TimeSpan.FromSeconds(Math.Max(1, seconds ?? _options.VbaTimeoutSeconds));
}
