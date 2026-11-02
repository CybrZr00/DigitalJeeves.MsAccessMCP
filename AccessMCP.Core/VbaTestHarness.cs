using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Extensions.Logging;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>A module to write into the scratch database.</summary>
internal sealed record VbaModuleSpec(string Name, string Kind, string? Code, string? FilePath)
{
    public string ResolveCode()
    {
        if (Code is not null)
        {
            return Code;
        }

        if (string.IsNullOrWhiteSpace(FilePath))
        {
            return string.Empty;
        }

        var path = Path.GetFullPath(FilePath);
        if (!File.Exists(path))
        {
            throw new FileNotFoundException($"Module file not found: {path}", path);
        }

        return File.ReadAllText(path);
    }

    public string StableKind => Kind.Trim().ToLowerInvariant() switch
    {
        "class" or "cls" => "class",
        _ => "standard"
    };

    public string Sha256 => VbaProjectEditor.Sha256Hex(VbaProjectEditor.NormalizeForWrite(ResolveCode(), string.Empty));
}

internal sealed record VbaModuleState(string Name, string Kind, bool Created, int Lines, string Sha256, string? BackupPath);

internal sealed record VbaPrepareInfo(
    IReadOnlyList<VbaModuleState> Modules,
    VbaCompileResult? Compile,
    MdeBuildResult? Mde,
    bool CacheHit,
    bool KitInjected,
    string? RunnerModule);

internal sealed record VbaValueInfo(string Type, string Text, JsonElement? Json);

internal sealed record VbaErrorInfo(int? Number, string? Description, string? Source, string? Module, int? Line, string? Stack);

internal sealed record VbaRunInfo(
    bool Ok,
    string? Procedure,
    long ElapsedMs,
    VbaValueInfo? Result,
    VbaErrorInfo? Error,
    bool Crashed,
    int? ExitCode,
    bool TimedOut,
    IReadOnlyList<string> LogTail,
    string? DialogText)
{
    /// <summary>Scratch directory, when it was kept (requested or the run failed).</summary>
    public string? ArtifactDirectory { get; init; }

    /// <summary>True when the scratch directory was left on disk.</summary>
    public bool ArtifactsKept { get; init; }
}

internal sealed record VbaScenarioRequest(
    string DatabasePath,
    IReadOnlyList<VbaModuleSpec> Modules,
    string? EntryPoint,
    IReadOnlyList<object?> Args,
    string Mode,
    string MdeStrategy,
    bool DecompileFirst,
    bool UseCache,
    bool Rebuild,
    bool ContinueOnCompileError,
    bool KeepArtifacts,
    int TimeoutSeconds,
    bool InPlace);

internal sealed record VbaScenarioInfo(
    bool Ok,
    string Mode,
    string SourceDatabasePath,
    string WorkDatabasePath,
    string? MdePath,
    VbaPrepareInfo Prepare,
    VbaRunInfo? Run,
    VbaTestSummary? Tests,
    string? ArtifactDirectory,
    bool ArtifactsKept,
    bool Cleaned,
    IReadOnlyList<string> CleanupErrors,
    string? Error);

internal sealed record VbaTestRunInfo(VbaRunInfo Run, VbaTestSummary Tests);

/// <summary>
/// Orchestrates a VBA scenario: copy the database to a scratch workspace, prepare modules (and the
/// test kit/runner), compile, optionally build an MDE, run the entry point, collect results and clean
/// up. The developer's database is never modified.
/// </summary>
internal sealed class VbaTestHarness(AccessMcpOptions options, ILoggerFactory loggerFactory)
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true
    };

    private readonly AccessMcpOptions _options = options;
    private readonly ILogger _logger = loggerFactory.CreateLogger("AccessMCP.VbaTestHarness");

    public async Task<VbaRunInfo> RunProcedureAsync(
        string databasePath,
        string? entryPoint,
        IReadOnlyList<object?> args,
        int? timeoutSeconds,
        CancellationToken cancellationToken)
    {
        var source = AccessDbPathPolicy.NormalizeComDatabasePath(databasePath, _options);
        var compiled = AccessDbPathPolicy.IsCompiledDatabase(source);
        using var workspace = VbaWorkspace.Create(_options, source, createMdePath: false);
        workspace.Keep = _options.KeepVbaArtifacts;

        var execution = await RunInWorkspaceAsync(
            workspace,
            compiled,
            entryPoint,
            args,
            timeoutSeconds,
            cancellationToken).ConfigureAwait(false);

        workspace.Cleanup(keepOnFailure: true, failed: !execution.Ok);
        return execution with
        {
            ArtifactDirectory = workspace.Deleted ? null : workspace.Directory,
            ArtifactsKept = !workspace.Deleted
        };
    }

    public async Task<VbaRunInfo> EvalAsync(
        string databasePath,
        string expression,
        int? timeoutSeconds,
        CancellationToken cancellationToken)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(expression);
        var source = AccessDbPathPolicy.NormalizeComDatabasePath(databasePath, _options);
        using var workspace = VbaWorkspace.Create(_options, source, createMdePath: false);
        workspace.Keep = _options.KeepVbaArtifacts;

        var com = new AccessComAutomation(_options);
        var timeout = Timeout(timeoutSeconds);
        var outcome = await com.RunInAccessWithWatchdogAsync(
            workspace.WorkDatabasePath,
            app =>
            {
                var stopwatch = Stopwatch.StartNew();
                try
                {
                    var value = AccessComAutomation.InvokeMethod(app, "Eval", expression);
                    stopwatch.Stop();
                    return new RunExecution(value, null, stopwatch.ElapsedMilliseconds, "Eval");
                }
                catch (Exception ex) when (!AccessComAutomation.IsCrashException(ex))
                {
                    stopwatch.Stop();
                    return new RunExecution(null, ex, stopwatch.ElapsedMilliseconds, "Eval");
                }
            },
            new AccessRunOptions { Timeout = timeout },
            cancellationToken).ConfigureAwait(false);

        var run = MapOutcome(outcome, workspace, outcome.Value as RunExecution);
        workspace.Cleanup(keepOnFailure: true, failed: !run.Ok);
        return run with
        {
            ArtifactDirectory = workspace.Deleted ? null : workspace.Directory,
            ArtifactsKept = !workspace.Deleted
        };
    }

    public async Task<VbaTestRunInfo> RunTestsAsync(
        string databasePath,
        IReadOnlyList<string>? tests,
        int? timeoutSeconds,
        bool keepArtifacts,
        CancellationToken cancellationToken)
    {
        var source = AccessDbPathPolicy.NormalizeComDatabasePath(databasePath, _options);
        if (AccessDbPathPolicy.IsCompiledDatabase(source))
        {
            throw new InvalidOperationException(
                "Tests cannot be injected into a compiled database. Use access_vba_test with mode 'mde' and the 'syscmd' strategy so the runner is baked in before the build.");
        }

        using var workspace = VbaWorkspace.Create(_options, source, createMdePath: false);
        workspace.Keep = keepArtifacts || _options.KeepVbaArtifacts;

        var (run, summary, prepare) = await PrepareAndRunAsync(
            workspace,
            Array.Empty<VbaModuleSpec>(),
            entryPoint: null,
            args: [],
            tests,
            timeoutSeconds,
            cancellationToken).ConfigureAwait(false);

        workspace.Cleanup(keepOnFailure: true, failed: !run.Ok || summary is { Ok: false });
        run = run with
        {
            ArtifactDirectory = workspace.Deleted ? null : workspace.Directory,
            ArtifactsKept = !workspace.Deleted
        };
        _ = prepare;
        return new VbaTestRunInfo(run, summary ?? new VbaTestSummary(false, 0, 0, 0, 0, 0, 0, [], [], run.Error?.Description, null));
    }

    public async Task<VbaScenarioInfo> RunScenarioAsync(VbaScenarioRequest request, CancellationToken cancellationToken)
    {
        if (request.InPlace)
        {
            throw new NotSupportedException("inPlace runs are not implemented; the harness only works on a scratch copy.");
        }

        var source = AccessDbPathPolicy.NormalizeComDatabasePath(request.DatabasePath, _options);
        var mode = request.Mode.Trim().ToLowerInvariant();
        if (mode is not ("mdb" or "mde"))
        {
            throw new ArgumentException("mode must be 'mdb' or 'mde'.", nameof(request));
        }

        var strategy = request.MdeStrategy.Trim().ToLowerInvariant();
        if (strategy is not ("syscmd" or "prebuilt"))
        {
            throw new ArgumentException("mdeStrategy must be 'syscmd' or 'prebuilt'.", nameof(request));
        }

        if (mode == "mdb" && AccessDbPathPolicy.IsCompiledDatabase(source))
        {
            throw new ArgumentException("A compiled .mde/.accde cannot be used as the source for mode 'mdb'.", nameof(request));
        }

        if (mode == "mde" && strategy == "prebuilt" && !AccessDbPathPolicy.IsCompiledDatabase(source))
        {
            throw new ArgumentException("The prebuilt MDE strategy requires databasePath to be an .mde or .accde file.", nameof(request));
        }

        if (mode == "mde" && strategy == "prebuilt" && request.Modules.Count > 0)
        {
            throw new ArgumentException(
                "The prebuilt MDE strategy cannot inject modules; the compiled database already contains its code.",
                nameof(request));
        }

        if (mode == "mde" && strategy == "prebuilt" && string.IsNullOrWhiteSpace(request.EntryPoint))
        {
            throw new ArgumentException(
                "The prebuilt MDE strategy requires an entryPoint, because the harness cannot discover tests in a compiled database.",
                nameof(request));
        }

        using var workspace = VbaWorkspace.Create(_options, source, createMdePath: mode == "mde");
        workspace.Keep = request.KeepArtifacts || _options.KeepVbaArtifacts;
        workspace.Log($"scenario start mode={mode} strategy={strategy} source={source}");

        var prepared = default(VbaPrepareInfo);
        VbaRunInfo? run = null;
        VbaTestSummary? summary = null;
        string? error = null;

        try
        {
            if (request.DecompileFirst && mode != "mdb")
            {
                error = "decompileFirst only applies to mode 'mdb'.";
            }
            else if (request.DecompileFirst)
            {
                var executable = AccessProcess.FindExecutable()
                    ?? throw new InvalidOperationException("Microsoft Access is not installed.");
                workspace.Log("decompiling");
                var exit = AccessProcess.Decompile(executable, workspace.WorkDatabasePath, Timeout(request.TimeoutSeconds));
                workspace.Log($"decompile exit={exit}");
                if (exit is null || exit != 0)
                {
                    error = "MSACCESS.EXE /compact /decompile did not complete successfully.";
                }
            }

            if (error is null)
            {
                if (mode == "mdb")
                {
                    (run, summary, prepared) = await PrepareAndRunMdbAsync(workspace, request, cancellationToken).ConfigureAwait(false);
                }
                else if (strategy == "prebuilt")
                {
                    run = await RunCompiledAsync(workspace, request, cancellationToken).ConfigureAwait(false);
                    prepared = new VbaPrepareInfo([], null, new MdeBuildResult(true, "prebuilt", source, workspace.WorkDatabasePath, true, false, null), false, false, null);
                }
                else
                {
                    (run, summary, prepared) = await PrepareBuildAndRunMdeAsync(workspace, request, cancellationToken).ConfigureAwait(false);

                    // SysCmd(603) is undocumented. If Access drops the automation connection during
                    // the build, retry once in a fresh instance: the work database is already
                    // compiled and closed on disk, which is the state the builder needs.
                    if (prepared.Mde is { Ok: false } failedBuild && IsTransientMdeFailure(failedBuild))
                    {
                        workspace.Log("transient MDE build failure; retrying once: " + failedBuild.Error);
                        (run, summary, prepared) = await PrepareBuildAndRunMdeAsync(workspace, request, cancellationToken).ConfigureAwait(false);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            error = ex.Message;
            workspace.Log("scenario failed: " + ex);
        }

        var ok = error is null &&
                 prepared?.Compile?.Ok != false &&
                 prepared?.Mde?.Ok != false &&
                 run?.Ok == true &&
                 summary is not { Ok: false } &&
                 run is not { Crashed: true } &&
                 run is not { TimedOut: true };

        var failed = !ok;
        workspace.Cleanup(keepOnFailure: true, failed: failed);

        return new VbaScenarioInfo(
            ok,
            request.Mode,
            source,
            workspace.WorkDatabasePath,
            prepared?.Mde?.OutputPath ?? (mode == "mde" ? workspace.MdePath : null),
            prepared ?? new VbaPrepareInfo([], null, null, false, false, null),
            run,
            summary,
            workspace.Deleted ? null : workspace.Directory,
            !workspace.Deleted,
            workspace.Deleted,
            workspace.CleanupErrors,
            error);
    }

    public async Task<MdeBuildResult> BuildMdeAsync(
        string databasePath,
        string? outputPath,
        string strategy,
        bool decompileFirst,
        int? timeoutSeconds,
        CancellationToken cancellationToken)
    {
        var source = AccessDbPathPolicy.NormalizeComDatabasePath(databasePath, _options);
        var resolvedStrategy = string.IsNullOrWhiteSpace(strategy) ? _options.VbaMdeBuildStrategy : strategy;
        resolvedStrategy = resolvedStrategy.Trim().ToLowerInvariant();

        if (resolvedStrategy == "prebuilt")
        {
            var destination = outputPath is null
                ? Path.Combine(Path.GetDirectoryName(source) ?? ".", Path.GetFileNameWithoutExtension(source) + MdeBuilder.DefaultExtension(source))
                : Path.GetFullPath(outputPath);
            return MdeBuilder.UsePrebuilt(source, destination);
        }

        if (AccessDbPathPolicy.IsCompiledDatabase(source))
        {
            return new MdeBuildResult(false, resolvedStrategy, source, outputPath, false, true,
                "Cannot build an MDE from an already compiled database.");
        }

        using var workspace = VbaWorkspace.Create(_options, source, createMdePath: true);
        workspace.Keep = _options.KeepVbaArtifacts;

        if (decompileFirst || _options.VbaDecompileFirst)
        {
            var executable = AccessProcess.FindExecutable()
                ?? throw new InvalidOperationException("Microsoft Access is not installed.");
            var exit = AccessProcess.Decompile(executable, workspace.WorkDatabasePath, Timeout(timeoutSeconds));
            if (exit is null || exit != 0)
            {
                return new MdeBuildResult(false, resolvedStrategy, source, outputPath, false, true,
                    "MSACCESS.EXE /compact /decompile did not complete successfully.");
            }
        }

        var mdePath = workspace.MdePath
            ?? throw new InvalidOperationException("The workspace has no MDE output path.");

        var com = new AccessComAutomation(_options);
        var outcome = await com.RunInAccessWithWatchdogAsync(
            workspace.WorkDatabasePath,
            app =>
            {
                var compile = VbaCompiler.Compile((object)app);
                if (!compile.Ok)
                {
                    return new BuildExecution(compile, null);
                }

                dynamic dynamicApp = app;
                dynamicApp.CloseCurrentDatabase();
                var build = MdeBuilder.BuildWithSysCmd((object)app, workspace.WorkDatabasePath, mdePath);
                return new BuildExecution(compile, build);
            },
            new AccessRunOptions { Timeout = Timeout(timeoutSeconds) },
            cancellationToken).ConfigureAwait(false);

        if (!outcome.Ok || outcome.Value?.Build is null)
        {
            return new MdeBuildResult(false, resolvedStrategy, source, outputPath, false, true,
                outcome.Exception?.Message ?? outcome.Value?.Compile?.Message ?? "The MDE build did not complete.");
        }

        var buildResult = outcome.Value.Build;
        var destinationPath = outputPath is null
            ? Path.Combine(Path.GetDirectoryName(source) ?? ".", Path.GetFileNameWithoutExtension(source) + MdeBuilder.DefaultExtension(source))
            : Path.GetFullPath(outputPath);

        if (Path.GetFullPath(destinationPath).Equals(Path.GetFullPath(source), StringComparison.OrdinalIgnoreCase))
        {
            return new MdeBuildResult(false, resolvedStrategy, source, destinationPath, false, true,
                "The output path must not be the source database.");
        }

        Directory.CreateDirectory(Path.GetDirectoryName(destinationPath) ?? ".");
        File.Copy(buildResult.OutputPath!, destinationPath, overwrite: true);

        workspace.Cleanup(keepOnFailure: false, failed: false);
        return buildResult with { OutputPath = destinationPath, SourcePath = source };
    }

    private async Task<(VbaRunInfo Run, VbaTestSummary? Summary, VbaPrepareInfo Prepare)> PrepareAndRunMdbAsync(
        VbaWorkspace workspace,
        VbaScenarioRequest request,
        CancellationToken cancellationToken)
    {
        return await PrepareAndRunAsync(
            workspace,
            request.Modules,
            request.EntryPoint,
            request.Args,
            tests: null,
            request.TimeoutSeconds,
            cancellationToken,
            request.ContinueOnCompileError,
            injectKitForLogs: true).ConfigureAwait(false);
    }

    private async Task<(VbaRunInfo Run, VbaTestSummary? Summary, VbaPrepareInfo Prepare)> PrepareAndRunAsync(
        VbaWorkspace workspace,
        IReadOnlyList<VbaModuleSpec> modules,
        string? entryPoint,
        IReadOnlyList<object?> args,
        IReadOnlyList<string>? tests,
        int? timeoutSeconds,
        CancellationToken cancellationToken,
        bool continueOnCompileError = false,
        bool injectKitForLogs = false)
    {
        var com = new AccessComAutomation(_options);
        var timeout = Timeout(timeoutSeconds);

        var outcome = await com.RunInAccessWithWatchdogAsync(
            workspace.WorkDatabasePath,
            app =>
            {
                var stopwatch = Stopwatch.StartNew();
                var writes = new List<VbaModuleState>();
                VbaCompileResult? compile = null;
                VbaTestSummary? summary = null;
                IReadOnlyList<(string Module, VbaProcedureInfo Procedure)> discovered = [];

                object? vbe = null;
                object? project = null;
                try
                {
                    try
                    {
                        vbe = ((dynamic)app).VBE;
                        project = VbaProjectEditor.GetProject(vbe!, null);
                    }
                    catch (Exception ex)
                    {
                        workspace.Log("VBE unavailable: " + ex.Message);
                        project = null;
                    }

                    if (project is not null)
                    {
                        foreach (var module in modules)
                        {
                            var code = module.ResolveCode();
                            var write = VbaProjectEditor.EnsureModule(
                                project,
                                workspace.WorkDatabasePath,
                                module.Name,
                                module.Kind,
                                code,
                                createIfMissing: true,
                                allowOverwrite: true,
                                backupDirectory: workspace.Directory);
                            writes.Add(new VbaModuleState(write.ModuleName, module.StableKind, write.Created, write.Lines, write.Sha256, write.BackupPath));
                        }

                        discovered = VbaProjectEditor.DiscoverTests(project);
                        if (tests is { Count: > 0 })
                        {
                            discovered = discovered
                                .Where(t => tests.Any(name => MatchesTest(name, t.Module, t.Procedure.Name)))
                                .ToArray();
                        }
                    }

                    var needRunner = tests is not null || (entryPoint is null && discovered.Count > 0);
                    var kitInjected = false;
                    if (needRunner || injectKitForLogs)
                    {
                        if (project is null)
                        {
                            throw new InvalidOperationException(
                                "The VBE is not accessible, so the test kit/runner cannot be injected. " +
                                "Enable Trust Center \"Trust access to the VBA project object model\".");
                        }

                        var kit = VbaProjectEditor.EnsureModule(
                            project,
                            workspace.WorkDatabasePath,
                            VbaTestKit.ModuleName,
                            "standard",
                            VbaTestKit.Render(workspace.ResultPath),
                            createIfMissing: true,
                            allowOverwrite: true,
                            backupDirectory: workspace.Directory);
                        kitInjected = true;

                        if (needRunner)
                        {
                            var runnerCode = VbaTestKit.GenerateRunner(discovered);
                            VbaProjectEditor.EnsureModule(
                                project,
                                workspace.WorkDatabasePath,
                                VbaTestKit.RunnerModuleName,
                                "standard",
                                runnerCode,
                                createIfMissing: true,
                                allowOverwrite: true,
                                backupDirectory: workspace.Directory);
                        }
                    }

                    // Resolved before the compile on purpose: compiling (DoCmd.RunCommand 126 plus
                    // OpenModule) invalidates the VBProject RCW on some Access builds, and the
                    // qualified "Module.Procedure" check below is the only thing that dereferences it.
                    // Everything it needs (module list, procedures) has already been read above.
                    var procedure = ResolveEntryPoint(project, entryPoint, discovered, needRunner);

                    compile = VbaCompiler.Compile((object)app);
                    workspace.Log($"compile ok={compile.Ok} isCompiled={compile.IsCompiled}");

                    if (!compile.Ok && !continueOnCompileError)
                    {
                        stopwatch.Stop();
                        return new HarnessExecution(
                            null,
                            null,
                            stopwatch.ElapsedMilliseconds,
                            new VbaPrepareInfo(writes, compile, null, false, kitInjected, needRunner ? VbaTestKit.RunnerModuleName : null),
                            new VbaCompileFailure(compile));
                    }

                    workspace.Log($"running {procedure}");

                    try
                    {
                        var value = InvokeRun(app, procedure, args);
                        stopwatch.Stop();
                        return new HarnessExecution(
                            new RunExecution(value, null, stopwatch.ElapsedMilliseconds, procedure),
                            summary,
                            stopwatch.ElapsedMilliseconds,
                            new VbaPrepareInfo(writes, compile, null, false, kitInjected, needRunner ? VbaTestKit.RunnerModuleName : null),
                            null);
                    }
                    catch (Exception ex) when (!AccessComAutomation.IsCrashException(ex))
                    {
                        stopwatch.Stop();
                        return new HarnessExecution(
                            new RunExecution(null, ex, stopwatch.ElapsedMilliseconds, procedure),
                            summary,
                            stopwatch.ElapsedMilliseconds,
                            new VbaPrepareInfo(writes, compile, null, false, kitInjected, needRunner ? VbaTestKit.RunnerModuleName : null),
                            null);
                    }
                }
                finally
                {
                    AccessComAutomation.ReleaseComObject(project);
                    AccessComAutomation.ReleaseComObject(vbe);
                }
            },
            new AccessRunOptions { Timeout = timeout },
            cancellationToken).ConfigureAwait(false);

        var run = MapOutcome(outcome, workspace, outcome.Value?.Run);

        if (outcome.Value?.CompileFailure is { } compileFailure && run.Error is null)
        {
            run = run with { Ok = false, Error = ErrorFromCompile(compileFailure.Result) };
        }

        VbaTestSummary? testSummary = null;
        if (File.Exists(workspace.ResultPath))
        {
            testSummary = NdjsonResults.ParseTests(workspace.ResultPath, _options.VbaMaxResultChars, markIncomplete: !run.Ok);
        }

        if (run.Error is not null && testSummary?.Error is not null)
        {
            run = run with { Error = run.Error with { Description = run.Error.Description + " " + testSummary.Error } };
        }

        var prepare = outcome.Value?.Prepare ?? new VbaPrepareInfo([], outcome.Value?.CompileFailure?.Result, null, false, false, null);
        if (outcome.Exception is not null && prepare.Compile is null)
        {
            workspace.Log("session failed: " + outcome.Exception);
        }

        return (run, testSummary, prepare);
    }

    private async Task<(VbaRunInfo Run, VbaTestSummary? Summary, VbaPrepareInfo Prepare)> PrepareBuildAndRunMdeAsync(
        VbaWorkspace workspace,
        VbaScenarioRequest request,
        CancellationToken cancellationToken)
    {
        var mdePath = workspace.MdePath!;
        var compiled = AccessDbPathPolicy.IsCompiledDatabase(workspace.WorkDatabasePath);

        // A cache key is only meaningful when the runner's test list is knowable without opening the
        // database. Tests supplied through modules are parseable here; tests already inside the
        // source database bypass the cache.
        var suppliedTests = request.Modules
            .SelectMany(m => VbaProjectEditor.FindProcedures(m.ResolveCode()).Where(p => p.IsPublic && p.IsTest).Select(p => p.Name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var access = AccessProcess.Describe();
        var moduleHashes = request.Modules
            .Select(m => new VbaCachedModule(
                m.Name,
                m.StableKind,
                VbaProjectEditor.Sha256Hex(VbaProjectEditor.NormalizeForWrite(m.ResolveCode(), string.Empty)),
                0))
            .ToArray();

        var cacheKey = VbaMdeCache.ComputeKey(
            request.DatabasePath,
            request.Modules,
            VbaTestKit.Sha256,
            request.DecompileFirst || _options.VbaDecompileFirst,
            access.Version ?? "unknown",
            access.Bitness ?? "unknown",
            request.EntryPoint ?? VbaTestKit.RunnerEntryPoint,
            suppliedTests);

        var cache = new VbaMdeCache(_options);
        // rebuild=true bypasses a cache read but still refreshes the entry for the next call.
        var cacheAllowed = _options.VbaCacheEnabled && request.UseCache && suppliedTests.Length > 0;
        var useCache = cacheAllowed && !request.Rebuild;
        var cacheHit = false;
        string? cacheProblem = null;

        if (useCache && cache.TryGet(cacheKey, out var cachedPath, out cacheProblem))
        {
            workspace.Log("MDE cache hit");
            File.Copy(cachedPath, mdePath, overwrite: true);
            cacheHit = true;
        }
        else if (useCache && cacheProblem is not null)
        {
            workspace.Log("MDE cache miss: " + cacheProblem);
        }

        VbaPrepareInfo prepare;
        if (cacheHit)
        {
            prepare = new VbaPrepareInfo([], null, new MdeBuildResult(true, "syscmd", workspace.WorkDatabasePath, mdePath, true, false, null), true, true, null);
        }
        else
        {
            var com = new AccessComAutomation(_options);
            var timeout = Timeout(request.TimeoutSeconds);
            var outcome = await com.RunInAccessWithWatchdogAsync(
                workspace.WorkDatabasePath,
                app =>
                {
                    var writes = new List<VbaModuleState>();
                    VbaCompileResult? compile = null;
                    IReadOnlyList<(string Module, VbaProcedureInfo Procedure)> discovered = [];

                    object? vbe = null;
                    object? project = null;
                    try
                    {
                        vbe = ((dynamic)app).VBE;
                        project = VbaProjectEditor.GetProject(vbe!, null);

                        foreach (var module in request.Modules)
                        {
                            var write = VbaProjectEditor.EnsureModule(
                                project,
                                workspace.WorkDatabasePath,
                                module.Name,
                                module.Kind,
                                module.ResolveCode(),
                                createIfMissing: true,
                                allowOverwrite: true,
                                backupDirectory: workspace.Directory);
                            writes.Add(new VbaModuleState(write.ModuleName, module.StableKind, write.Created, write.Lines, write.Sha256, write.BackupPath));
                        }

                        discovered = VbaProjectEditor.DiscoverTests(project);
                        var needRunner = request.EntryPoint is null || discovered.Count > 0;

                        VbaProjectEditor.EnsureModule(
                            project,
                            workspace.WorkDatabasePath,
                            VbaTestKit.ModuleName,
                            "standard",
                            VbaTestKit.Render(workspace.ResultPath),
                            createIfMissing: true,
                            allowOverwrite: true,
                            backupDirectory: workspace.Directory);

                        if (needRunner)
                        {
                            VbaProjectEditor.EnsureModule(
                                project,
                                workspace.WorkDatabasePath,
                                VbaTestKit.RunnerModuleName,
                                "standard",
                                VbaTestKit.GenerateRunner(discovered),
                                createIfMissing: true,
                                allowOverwrite: true,
                                backupDirectory: workspace.Directory);
                        }

                        compile = VbaCompiler.Compile((object)app);
                        workspace.Log($"MDE source compile ok={compile.Ok} isCompiled={compile.IsCompiled}");
                    }
                    finally
                    {
                        // Holding a VBE project/component reference across CloseCurrentDatabase can
                        // crash Access inside SysCmd(603). Release everything before closing.
                        AccessComAutomation.ReleaseComObject(project);
                        AccessComAutomation.ReleaseComObject(vbe);
                        project = null;
                        vbe = null;
                    }

                    if (compile is null || !compile.Ok)
                    {
                        return new MdePhaseExecution(writes, compile, null, discovered);
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    dynamic dynamicApp = app;
                    dynamicApp.CloseCurrentDatabase();
                    var build = MdeBuilder.BuildWithSysCmd((object)app, workspace.WorkDatabasePath, mdePath);
                    workspace.Log($"syscmd build ok={build.Ok}: {build.Error}");
                    return new MdePhaseExecution(writes, compile, build, discovered);
                },
                new AccessRunOptions { Timeout = timeout },
                cancellationToken).ConfigureAwait(false);

            var phase = outcome.Value;
            var buildResult = phase?.Build ?? new MdeBuildResult(false, "syscmd", workspace.WorkDatabasePath, mdePath, false, true,
                outcome.Exception?.Message ?? phase?.Compile?.Message ?? "The MDE build did not complete.");

            if (buildResult.Ok)
            {
                var manifest = new VbaCacheManifest(
                    cacheKey,
                    AccessDbPathPolicy.GetDatabaseKind(workspace.WorkDatabasePath),
                    Path.GetFullPath(request.DatabasePath),
                    new FileInfo(request.DatabasePath).Length,
                    new FileInfo(request.DatabasePath).LastWriteTimeUtc.Ticks,
                    access.Version ?? "unknown",
                    access.Bitness ?? "unknown",
                    VbaTestKit.Sha256,
                    request.DecompileFirst || _options.VbaDecompileFirst,
                    request.EntryPoint ?? VbaTestKit.RunnerEntryPoint,
                    suppliedTests,
                    moduleHashes,
                    DateTimeOffset.UtcNow);

                if (cacheAllowed)
                {
                    try
                    {
                        cache.Put(cacheKey, mdePath, manifest);
                    }
                    catch (Exception ex)
                    {
                        workspace.Log("cache write failed: " + ex.Message);
                    }
                }
            }

            prepare = new VbaPrepareInfo(
                phase?.Writes ?? [],
                phase?.Compile,
                buildResult,
                false,
                true,
                phase?.Discovered.Count > 0 ? VbaTestKit.RunnerModuleName : null);

            if (phase?.Compile is { Ok: false })
            {
                var runFailure = new VbaRunInfo(false, null, 0, null, ErrorFromCompile(phase.Compile), false, null, false, [], null);
                return (runFailure, null, prepare);
            }

            if (!buildResult.Ok)
            {
                var runFailure = new VbaRunInfo(false, null, 0, null, new VbaErrorInfo(null, buildResult.Error, "MdeBuilder", null, null, null), false, null, false, [], null);
                return (runFailure, null, prepare);
            }
        }

        // Phase B: run against the compiled copy.
        var run = await RunCompiledAsync(workspace, request, cancellationToken).ConfigureAwait(false);
        VbaTestSummary? summary = File.Exists(workspace.ResultPath)
            ? NdjsonResults.ParseTests(workspace.ResultPath, _options.VbaMaxResultChars, markIncomplete: !run.Ok)
            : null;

        return (run, summary, prepare);
    }

    private async Task<VbaRunInfo> RunCompiledAsync(
        VbaWorkspace workspace,
        VbaScenarioRequest request,
        CancellationToken cancellationToken)
    {
        var target = workspace.MdePath is { } mde && File.Exists(mde)
            ? mde
            : workspace.WorkDatabasePath;
        var com = new AccessComAutomation(_options);
        var timeout = Timeout(request.TimeoutSeconds);

        var outcome = await com.RunInAccessWithWatchdogAsync(
            target,
            app =>
            {
                var stopwatch = Stopwatch.StartNew();
                var procedure = ResolveEntryPointStatically(request.EntryPoint);
                try
                {
                    var value = InvokeRun(app, procedure, request.Args);
                    stopwatch.Stop();
                    return new RunExecution(value, null, stopwatch.ElapsedMilliseconds, procedure);
                }
                catch (Exception ex) when (!AccessComAutomation.IsCrashException(ex))
                {
                    stopwatch.Stop();
                    return new RunExecution(null, ex, stopwatch.ElapsedMilliseconds, procedure);
                }
            },
            new AccessRunOptions { Timeout = timeout },
            cancellationToken).ConfigureAwait(false);

        return MapOutcome(outcome, workspace, outcome.Value);
    }

    private async Task<VbaRunInfo> RunInWorkspaceAsync(
        VbaWorkspace workspace,
        bool compiled,
        string? entryPoint,
        IReadOnlyList<object?> args,
        int? timeoutSeconds,
        CancellationToken cancellationToken)
    {
        var com = new AccessComAutomation(_options);
        var timeout = Timeout(timeoutSeconds);

        var outcome = await com.RunInAccessWithWatchdogAsync(
            workspace.WorkDatabasePath,
            app =>
            {
                var stopwatch = Stopwatch.StartNew();
                string? procedure = null;
                object? vbe = null;
                object? project = null;

                try
                {
                    if (!compiled)
                    {
                        try
                        {
                            vbe = ((dynamic)app).VBE;
                            project = VbaProjectEditor.GetProject(vbe!, null);
                            VbaProjectEditor.EnsureModule(
                                project,
                                workspace.WorkDatabasePath,
                                VbaTestKit.ModuleName,
                                "standard",
                                VbaTestKit.Render(workspace.ResultPath),
                                createIfMissing: true,
                                allowOverwrite: true,
                                backupDirectory: workspace.Directory);
                        }
                        catch (Exception ex)
                        {
                            workspace.Log("kit injection skipped: " + ex.Message);
                        }
                    }

                    procedure = ResolveEntryPoint(project, entryPoint, [], needRunner: false);
                    var value = InvokeRun(app, procedure, args);
                    stopwatch.Stop();
                    return new RunExecution(value, null, stopwatch.ElapsedMilliseconds, procedure);
                }
                catch (Exception ex) when (!AccessComAutomation.IsCrashException(ex))
                {
                    stopwatch.Stop();
                    return new RunExecution(null, ex, stopwatch.ElapsedMilliseconds, procedure);
                }
                finally
                {
                    AccessComAutomation.ReleaseComObject(project);
                    AccessComAutomation.ReleaseComObject(vbe);
                }
            },
            new AccessRunOptions { Timeout = timeout },
            cancellationToken).ConfigureAwait(false);

        return MapOutcome(outcome, workspace, outcome.Value);
    }

    private VbaRunInfo MapOutcome<T>(AccessRunOutcome<T> outcome, VbaWorkspace workspace, RunExecution? execution)
    {
        var logs = File.Exists(workspace.ResultPath)
            ? NdjsonResults.ParseLog(workspace.ResultPath, _options.VbaMaxResultChars).Log
            : [];

        if (outcome.TimedOut)
        {
            return new VbaRunInfo(false, null, 0, null,
                new VbaErrorInfo(null, outcome.Exception?.Message ?? "Timed out.", "AccessMCP", null, null, null),
                false, null, true, logs, outcome.DialogText);
        }

        if (outcome.Cancelled)
        {
            return new VbaRunInfo(false, null, 0, null,
                new VbaErrorInfo(null, "The run was cancelled.", "AccessMCP", null, null, null),
                false, null, false, logs, outcome.DialogText);
        }

        if (outcome.Crashed)
        {
            return new VbaRunInfo(false, execution?.Procedure, execution?.ElapsedMs ?? 0, null,
                new VbaErrorInfo(null, outcome.Exception?.Message ?? "Access crashed.", "Access", null, null, outcome.Exception?.StackTrace),
                true, outcome.ExitCode, false, logs, outcome.DialogText);
        }

        if (outcome.Exception is not null)
        {
            return new VbaRunInfo(false, execution?.Procedure, execution?.ElapsedMs ?? 0, null,
                MapError(outcome.Exception), false, null, false, logs, outcome.DialogText);
        }

        if (execution?.Error is not null)
        {
            return new VbaRunInfo(false, execution.Procedure, execution.ElapsedMs, null,
                MapError(execution.Error), false, null, false, logs, outcome.DialogText);
        }

        return new VbaRunInfo(
            true,
            execution?.Procedure,
            execution?.ElapsedMs ?? 0,
            execution is null ? null : ConvertValue(execution.Value),
            null,
            false,
            null,
            false,
            logs,
            outcome.DialogText);
    }

    private static object? InvokeRun(dynamic app, string procedure, IReadOnlyList<object?> args)
    {
        if (args.Count == 0)
        {
            return AccessComAutomation.InvokeMethod(app, "Run", procedure);
        }

        var all = new object?[args.Count + 1];
        all[0] = procedure;
        for (var i = 0; i < args.Count; i++)
        {
            all[i + 1] = args[i];
        }

        return AccessComAutomation.InvokeMethod(app, "Run", all);
    }

    private static string ResolveEntryPoint(
        object? project,
        string? entryPoint,
        IReadOnlyList<(string Module, VbaProcedureInfo Procedure)> discovered,
        bool needRunner)
    {
        if (string.IsNullOrWhiteSpace(entryPoint))
        {
            if (needRunner || discovered.Count > 0)
            {
                return VbaTestKit.RunnerEntryPoint;
            }

            throw new ArgumentException("An entryPoint is required when no '@test procedures were discovered.", nameof(entryPoint));
        }

        var trimmed = entryPoint.Trim();
        var dot = trimmed.LastIndexOf('.');
        var procedure = dot > 0 ? trimmed[(dot + 1)..] : trimmed;
        var module = dot > 0 ? trimmed[..dot] : null;

        if (project is not null && !string.IsNullOrWhiteSpace(module))
        {
            var owner = VbaProjectEditor.FindProcedureModule(project, procedure);
            if (owner is null)
            {
                throw new InvalidOperationException(
                    $"Procedure '{procedure}' was not found in any module. Access's Application.Run takes an " +
                    "unqualified name; the module part is used only to validate the entry point.");
            }
        }

        return procedure;
    }

    private static string ResolveEntryPointStatically(string? entryPoint)
    {
        if (string.IsNullOrWhiteSpace(entryPoint))
        {
            return VbaTestKit.RunnerEntryPoint;
        }

        var trimmed = entryPoint.Trim();
        var dot = trimmed.LastIndexOf('.');
        return dot > 0 ? trimmed[(dot + 1)..] : trimmed;
    }

    private static bool MatchesTest(string requested, string module, string procedure)
        => requested.Equals(procedure, StringComparison.OrdinalIgnoreCase) ||
           requested.Equals(module + "." + procedure, StringComparison.OrdinalIgnoreCase);

    private static VbaErrorInfo ErrorFromCompile(VbaCompileResult compile)
        => new(null, compile.Message, "VbaCompiler", compile.Module, compile.Line, null);

    private static bool IsTransientMdeFailure(MdeBuildResult result)
    {
        var error = result.Error ?? string.Empty;
        return error.Contains("early-bound", StringComparison.OrdinalIgnoreCase) ||
               error.Contains("remote procedure call", StringComparison.OrdinalIgnoreCase) ||
               error.Contains("RPC", StringComparison.OrdinalIgnoreCase) ||
               error.Contains("0x800706BE", StringComparison.OrdinalIgnoreCase);
    }

    private static VbaErrorInfo MapError(Exception exception)
    {
        if (exception is COMException com)
        {
            var number = com.ErrorCode & 0xFFFF;
            var message = com.Message;
            var marker = message.IndexOf(" (Exception from HRESULT", StringComparison.OrdinalIgnoreCase);
            if (marker > 0)
            {
                message = message[..marker];
            }

            return new VbaErrorInfo(number, message, "VBA", null, null, exception.StackTrace);
        }

        return new VbaErrorInfo(null, exception.Message, "AccessMCP", null, null, exception.StackTrace);
    }

    private static VbaValueInfo ConvertValue(object? value)
    {
        if (value is null or DBNull)
        {
            return new VbaValueInfo("Empty", string.Empty, null);
        }

        if (value is string text)
        {
            JsonElement? json = null;
            var trimmed = text.TrimStart();
            if (trimmed.StartsWith('{') || trimmed.StartsWith('['))
            {
                try
                {
                    using var document = JsonDocument.Parse(text);
                    json = document.RootElement.Clone();
                }
                catch
                {
                }
            }

            return new VbaValueInfo("String", text, json);
        }

        if (value is Array array)
        {
            return new VbaValueInfo("Array", $"[Array {array.Length}]", null);
        }

        if (value is DateTime date)
        {
            return new VbaValueInfo("Date", date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), null);
        }

        if (value is bool boolean)
        {
            return new VbaValueInfo("Boolean", boolean ? "True" : "False", null);
        }

        if (Marshal.IsComObject(value))
        {
            return new VbaValueInfo("Object", $"[Object: {value.GetType().Name}]", null);
        }

        if (value is IFormattable formattable)
        {
            return new VbaValueInfo(value.GetType().Name, formattable.ToString(null, CultureInfo.InvariantCulture) ?? string.Empty, null);
        }

        return new VbaValueInfo(value.GetType().Name, value.ToString() ?? string.Empty, null);
    }

    private TimeSpan Timeout(int? seconds)
        => TimeSpan.FromSeconds(Math.Max(1, seconds ?? _options.VbaTimeoutSeconds));

    internal sealed record RunExecution(object? Value, Exception? Error, long ElapsedMs, string? Procedure);

    internal sealed record HarnessExecution(
        RunExecution? Run,
        VbaTestSummary? Summary,
        long ElapsedMs,
        VbaPrepareInfo Prepare,
        VbaCompileFailure? CompileFailure);

    internal sealed record VbaCompileFailure(VbaCompileResult Result);

    internal sealed record MdePhaseExecution(
        IReadOnlyList<VbaModuleState> Writes,
        VbaCompileResult? Compile,
        MdeBuildResult? Build,
        IReadOnlyList<(string Module, VbaProcedureInfo Procedure)> Discovered);

    internal sealed record BuildExecution(VbaCompileResult Compile, MdeBuildResult? Build);
}
