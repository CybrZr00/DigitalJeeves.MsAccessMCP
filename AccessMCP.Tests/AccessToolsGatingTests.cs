using System.Reflection;
using System.Text.Json;
using DigitalJeeves.MsAccessMCP.Core;
using Microsoft.Extensions.Logging.Abstractions;
using ModelContextProtocol.Server;

namespace DigitalJeeves.MsAccessMCP.Tests;

/// <summary>
/// The checks every tool makes before touching a database: ReadOnly (enforced per tool, nothing does it
/// centrally), EnableComAutomation and EnableVbaTestTools. None of these tests start Access.
/// </summary>
public sealed class AccessToolsGatingTests : IDisposable
{
    private readonly TempDirectory _dir = new();

    public void Dispose() => _dir.Dispose();

    /// <summary>
    /// Write tools that are allowed to skip the up-front ReadOnly check, and why. Anything else marked
    /// ReadOnly = false must appear in <see cref="ReadOnlyCases"/>.
    /// </summary>
    private static readonly Dictionary<string, string> ReadOnlyExemptions = new()
    {
        ["access_execute_saved_query"] = "Checks ReadOnly inside the COM session, once it knows the query does not return rows.",
        ["access_vba_run"] = "Runs against a scratch copy; gated by EnableVbaTestTools instead.",
        ["access_vba_eval"] = "Runs against a scratch copy; gated by EnableVbaTestTools instead.",
        ["access_vba_run_tests"] = "Runs against a scratch copy; gated by EnableVbaTestTools instead.",
        ["access_vba_test"] = "Runs against a scratch copy; gated by EnableVbaTestTools instead.",
        ["access_vba_cache_clear"] = "Only deletes the server's own MDE cache, never a caller's database."
    };

    public static TheoryData<string> ReadOnlyCases => [.. ReadOnlyChecks.Keys];

    private static readonly Dictionary<string, Func<AccessTools, AccessVbaTools, string, Task<string>>> ReadOnlyChecks = new()
    {
        ["access_execute"] = (t, _, db) => t.ExecuteAsync("DELETE FROM T", db),
        ["access_set_form_control_value"] = (t, _, db) => t.SetFormControlValueAsync("frm", "ctl", databasePath: db),
        ["access_add_form_control"] = (t, _, db) => t.AddFormControlAsync("frm", "textbox", "ctl", databasePath: db),
        ["access_remove_form_control"] = (t, _, db) => t.RemoveFormControlAsync("frm", "ctl", db),
        ["access_import_changed_modules_from_folder"] = (t, _, db) => t.ImportChangedModulesFromFolderAsync(@"C:\Data\Export", databasePath: db),
        ["access_import_module_code"] = (t, _, db) => t.ImportModuleCodeAsync("Module1", "Sub A()\r\nEnd Sub", databasePath: db),
        ["access_compile_vba"] = (t, _, db) => t.CompileVbaAsync(db),
        ["access_set_object_properties"] = (t, _, db) => t.SetObjectPropertiesAsync("form", "frm", JsonDocument.Parse("{}").RootElement, db),
        ["access_set_control_properties"] = (t, _, db) => t.SetControlPropertiesAsync("form", "frm", "ctl", JsonDocument.Parse("{}").RootElement, db),
        ["access_add_report_control"] = (t, _, db) => t.AddReportControlAsync("rpt", "textbox", "ctl", databasePath: db),
        ["access_remove_report_control"] = (t, _, db) => t.RemoveReportControlAsync("rpt", "ctl", db),
        ["access_run_macro"] = (t, _, db) => t.RunMacroAsync("mcrMain", db),
        ["access_do_cmd"] = (t, _, db) => t.DoCmdAsync("open_form", "frmMain", databasePath: db),
        ["access_vba_ensure_module"] = (_, v, db) => v.EnsureModuleAsync("Module1", "Sub A()\r\nEnd Sub", databasePath: db),
        ["access_vba_remove_module"] = (_, v, db) => v.RemoveModuleAsync("AccessMcpTemp", databasePath: db),
        ["access_vba_import_file"] = (_, v, db) => v.ImportFileAsync(@"C:\Data\Module1.bas", databasePath: db),
        ["access_vba_compile"] = (_, v, db) => v.CompileAsync(db),
        ["access_vba_build_mde"] = (_, v, db) => v.BuildMdeAsync(db)
    };

    [Theory]
    [MemberData(nameof(ReadOnlyCases))]
    public async Task WriteToolRefusesWhenReadOnly(string toolName)
    {
        // Everything else is enabled, so the only reason left to refuse is ReadOnly. The database file is
        // a placeholder: a tool that got past the check would fail on it rather than modify anything.
        var options = new AccessMcpOptions
        {
            ReadOnly = true,
            EnableComAutomation = true,
            EnableVbaTestTools = true
        };
        var db = _dir.Touch("Sample.accdb");

        var message = await RefusalMessage(() => ReadOnlyChecks[toolName](
            new AccessTools(options),
            new AccessVbaTools(options, NullLoggerFactory.Instance),
            db));

        Assert.Contains("read-only", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EveryWriteToolIsCoveredByTheReadOnlyTestsOrExempt()
    {
        var writeTools = ToolMethods()
            .Where(t => !t.Attribute.ReadOnly)
            .Select(t => t.Attribute.Name!)
            .ToHashSet();

        var uncovered = writeTools.Except(ReadOnlyChecks.Keys).Except(ReadOnlyExemptions.Keys).Order().ToArray();
        var stale = ReadOnlyChecks.Keys.Concat(ReadOnlyExemptions.Keys).Except(writeTools).Order().ToArray();

        Assert.True(uncovered.Length == 0, "Write tools with no ReadOnly test or exemption: " + string.Join(", ", uncovered));
        Assert.True(stale.Length == 0, "ReadOnly test entries for tools that are gone or now read-only: " + string.Join(", ", stale));
    }

    [Fact]
    public void ToolNamesAreUniqueAndFollowTheAccessPrefix()
    {
        var names = ToolMethods().Select(t => t.Attribute.Name!).ToArray();

        Assert.NotEmpty(names);
        Assert.Equal(names.Length, names.Distinct(StringComparer.Ordinal).Count());
        Assert.All(names, n => Assert.Matches("^access_[a-z0-9_]+$", n));
    }

    [Fact]
    public void ToolMethodsReturnStringTasksAndAcceptCancellation()
    {
        Assert.All(ToolMethods(), t =>
        {
            Assert.Equal(typeof(Task<string>), t.Method.ReturnType);
            Assert.Contains(t.Method.GetParameters(), p => p.ParameterType == typeof(CancellationToken));
        });
    }

    [Fact]
    public async Task ListVbeProjectsAsync_WhenComAutomationDisabled_Throws()
    {
        var tools = ToolsWithComDisabled(out var dbPath);

        await Assert.ThrowsAsync<InvalidOperationException>(() => tools.ListVbeProjectsAsync(dbPath));
    }

    [Fact]
    public async Task ListVbeModulesAsync_WhenComAutomationDisabled_Throws()
    {
        var tools = ToolsWithComDisabled(out var dbPath);

        await Assert.ThrowsAsync<InvalidOperationException>(() => tools.ListVbeModulesAsync(databasePath: dbPath));
    }

    [Fact]
    public async Task GetVbeModuleCodeAsync_WhenComAutomationDisabled_Throws()
    {
        var tools = ToolsWithComDisabled(out var dbPath);

        await Assert.ThrowsAsync<InvalidOperationException>(() => tools.GetVbeModuleCodeAsync("Module1", databasePath: dbPath));
    }

    [Fact]
    public async Task GetVbeProcedureInfoAsync_WhenComAutomationDisabled_Throws()
    {
        var tools = ToolsWithComDisabled(out var dbPath);

        await Assert.ThrowsAsync<InvalidOperationException>(() => tools.GetVbeProcedureInfoAsync("Module1", databasePath: dbPath));
    }

    [Fact]
    public async Task ListFormsAsync_WhenComAutomationDisabled_Throws()
    {
        var tools = ToolsWithComDisabled(out var dbPath);

        await Assert.ThrowsAsync<InvalidOperationException>(() => tools.ListFormsAsync(dbPath));
    }

    [Fact]
    public async Task DiffVbaModulesAsync_WhenComAutomationDisabled_Throws()
    {
        var tools = new AccessTools(new AccessMcpOptions { EnableComAutomation = false });

        await Assert.ThrowsAsync<InvalidOperationException>(
            () => tools.DiffVbaModulesAsync(_dir.Touch("a.accdb"), _dir.Touch("b.accdb")));
    }

    [Fact]
    public async Task DiffVbaModulesAsync_RequiresBothPaths()
    {
        var tools = new AccessTools(new AccessMcpOptions { EnableComAutomation = false });

        await Assert.ThrowsAsync<ArgumentException>(() => tools.DiffVbaModulesAsync(string.Empty, _dir.Touch("b.accdb")));
        await Assert.ThrowsAsync<ArgumentException>(() => tools.DiffVbaModulesAsync(_dir.Touch("a.accdb"), "   "));
    }

    [Fact]
    public async Task DiffVbaModulesAsync_RejectsNegativeContext()
    {
        var tools = new AccessTools(new AccessMcpOptions { EnableComAutomation = false });

        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(
            () => tools.DiffVbaModulesAsync(_dir.Touch("a.accdb"), _dir.Touch("b.accdb"), contextLines: -1));
    }

    [Fact]
    public async Task VbaExecutionTools_AreGatedByEnableVbaTestTools()
    {
        var db = _dir.Touch("Sample.accdb");
        var options = new AccessMcpOptions
        {
            DefaultDatabasePath = db,
            EnableComAutomation = true,
            EnableVbaTestTools = false
        };
        var tools = new AccessVbaTools(options, NullLoggerFactory.Instance);

        await Assert.ThrowsAsync<InvalidOperationException>(() => tools.RunAsync("Anything", databasePath: db));
        await Assert.ThrowsAsync<InvalidOperationException>(() => tools.EvalAsync("1+1", databasePath: db));
        await Assert.ThrowsAsync<InvalidOperationException>(() => tools.RunTestsAsync(databasePath: db));
        await Assert.ThrowsAsync<InvalidOperationException>(() => tools.TestAsync(db));
        await Assert.ThrowsAsync<InvalidOperationException>(() => tools.BuildMdeAsync(db));
    }

    [Fact]
    public async Task VbaExecutionTools_AlsoNeedComAutomation()
    {
        var db = _dir.Touch("Sample.accdb");
        var options = new AccessMcpOptions
        {
            EnableComAutomation = false,
            EnableVbaTestTools = true
        };
        var tools = new AccessVbaTools(options, NullLoggerFactory.Instance);

        var exception = await Assert.ThrowsAsync<InvalidOperationException>(() => tools.RunAsync("Anything", databasePath: db));
        Assert.Contains("EnableComAutomation", exception.Message);
    }

    [Fact]
    public async Task VbaEditTools_WhenComAutomationDisabled_ReturnStructuredErrors()
    {
        var db = _dir.Touch("Sample.accdb");
        var tools = new AccessVbaTools(new AccessMcpOptions { EnableComAutomation = false }, NullLoggerFactory.Instance);

        foreach (var json in new[]
        {
            await tools.EnsureModuleAsync("Module1", "Sub A()\r\nEnd Sub", databasePath: db),
            await tools.RemoveModuleAsync("AccessMcpTemp", databasePath: db),
            await tools.ImportFileAsync(_dir.Combine("Module1.bas"), databasePath: db)
        })
        {
            using var document = JsonDocument.Parse(json);
            Assert.False(document.RootElement.GetProperty("ok").GetBoolean());
            Assert.Equal("error", document.RootElement.GetProperty("outcome").GetString());
            Assert.Equal(nameof(InvalidOperationException), document.RootElement.GetProperty("exceptionType").GetString());
        }
    }

    [Fact]
    public async Task EnsureModuleAsync_NeedsCodeOrCodeFile()
    {
        var tools = new AccessVbaTools(new AccessMcpOptions { EnableComAutomation = true }, NullLoggerFactory.Instance);

        using var document = JsonDocument.Parse(await tools.EnsureModuleAsync("Module1", databasePath: _dir.Touch("Sample.accdb")));

        Assert.False(document.RootElement.GetProperty("ok").GetBoolean());
        Assert.Contains("code or codeFile", document.RootElement.GetProperty("error").GetString());
    }

    [Theory]
    [InlineData("""{"name":"M","source":"x"}""", "JSON array")]
    [InlineData("""["M"]""", "must be an object")]
    [InlineData("""[{"source":"x"}]""", "needs a name")]
    [InlineData("""[{"name":"M"}]""", "needs source or filePath")]
    public async Task TestAsync_RejectsMalformedModulesBeforeStartingAccess(string modules, string expected)
    {
        var tools = new AccessVbaTools(
            new AccessMcpOptions { EnableComAutomation = true, EnableVbaTestTools = true },
            NullLoggerFactory.Instance);

        var exception = await Assert.ThrowsAsync<ArgumentException>(
            () => tools.TestAsync(_dir.Touch("Sample.accdb"), JsonDocument.Parse(modules).RootElement));

        Assert.Contains(expected, exception.Message);
    }

    [Theory]
    [InlineData("""{"a":1}""", "JSON array")]
    [InlineData("""[{"a":1}]""", "scalar JSON values")]
    public async Task TestAsync_RejectsArgumentsVbaCannotReceive(string args, string expected)
    {
        var tools = new AccessVbaTools(
            new AccessMcpOptions { EnableComAutomation = true, EnableVbaTestTools = true },
            NullLoggerFactory.Instance);

        var exception = await Assert.ThrowsAsync<ArgumentException>(
            () => tools.TestAsync(_dir.Touch("Sample.accdb"), args: JsonDocument.Parse(args).RootElement));

        Assert.Contains(expected, exception.Message);
    }

    [Fact]
    public async Task CacheClearAsync_NeedsAKeyOrAll()
    {
        var tools = new AccessVbaTools(new AccessMcpOptions { VbaCacheDirectory = _dir.Combine("cache") }, NullLoggerFactory.Instance);

        await Assert.ThrowsAsync<ArgumentException>(() => tools.CacheClearAsync());

        using var document = JsonDocument.Parse(await tools.CacheClearAsync(all: true));
        Assert.True(document.RootElement.GetProperty("ok").GetBoolean());
        Assert.Equal(0, document.RootElement.GetProperty("removed").GetInt32());
    }

    [Fact]
    public async Task CacheInfoAsync_ReportsTheConfiguredRoot()
    {
        var root = _dir.Combine("cache");
        var tools = new AccessVbaTools(new AccessMcpOptions { VbaCacheDirectory = root, VbaCacheMaxBytes = 1234 }, NullLoggerFactory.Instance);

        using var document = JsonDocument.Parse(await tools.CacheInfoAsync());

        Assert.Equal(root, document.RootElement.GetProperty("root").GetString());
        Assert.Equal(1234, document.RootElement.GetProperty("maxBytes").GetInt64());
        Assert.Equal(0, document.RootElement.GetProperty("entries").GetArrayLength());
    }

    [Fact]
    public async Task Doctor_WorksWithoutComAutomation()
    {
        var options = new AccessMcpOptions
        {
            EnableComAutomation = false,
            EnableVbaTestTools = false,
            VbaScratchDirectory = _dir.Combine("scratch"),
            VbaCacheDirectory = _dir.Combine("cache")
        };

        var tools = new AccessVbaTools(options, NullLoggerFactory.Instance);
        var json = await tools.DoctorAsync();
        using var document = JsonDocument.Parse(json);

        Assert.True(document.RootElement.TryGetProperty("access", out _));
        Assert.True(document.RootElement.TryGetProperty("capabilities", out var capabilities));
        Assert.False(capabilities.GetProperty("canEditModules").GetBoolean());
        Assert.False(capabilities.GetProperty("canRun").GetBoolean());
    }

    [Fact]
    public async Task Doctor_DoesNotPromiseRunsWhenComAutomationIsOff()
    {
        // EnableVbaTestTools alone is not enough: access_vba_run checks EnableComAutomation first.
        var options = new AccessMcpOptions
        {
            EnableComAutomation = false,
            EnableVbaTestTools = true,
            VbaScratchDirectory = _dir.Combine("scratch"),
            VbaCacheDirectory = _dir.Combine("cache")
        };

        using var document = JsonDocument.Parse(await new AccessVbaTools(options, NullLoggerFactory.Instance).DoctorAsync());
        var capabilities = document.RootElement.GetProperty("capabilities");

        Assert.False(capabilities.GetProperty("canRun").GetBoolean());
        Assert.False(capabilities.GetProperty("canBuildMde").GetBoolean());
        await Assert.ThrowsAsync<InvalidOperationException>(
            () => new AccessVbaTools(options, NullLoggerFactory.Instance).RunAsync("Anything", databasePath: _dir.Touch("Sample.accdb")));
    }

    private AccessTools ToolsWithComDisabled(out string dbPath)
    {
        dbPath = _dir.Touch("Sample.accdb");
        return new AccessTools(new AccessMcpOptions { DefaultDatabasePath = dbPath, EnableComAutomation = false });
    }

    private static IEnumerable<(MethodInfo Method, McpServerToolAttribute Attribute)> ToolMethods()
        => new[] { typeof(AccessTools), typeof(AccessVbaTools) }
            .SelectMany(type => type.GetMethods(BindingFlags.Public | BindingFlags.Instance))
            .Select(method => (Method: method, Attribute: method.GetCustomAttribute<McpServerToolAttribute>()))
            .Where(t => t.Attribute is not null)
            .Select(t => (t.Method, t.Attribute!));

    /// <summary>
    /// The refusal text, whether the tool throws (most tools) or returns the structured error JSON
    /// (the access_vba_* edit tools catch and serialize).
    /// </summary>
    private static async Task<string> RefusalMessage(Func<Task<string>> call)
    {
        string json;
        try
        {
            json = await call();
        }
        catch (InvalidOperationException ex)
        {
            return ex.Message;
        }

        using var document = JsonDocument.Parse(json);
        Assert.False(document.RootElement.GetProperty("ok").GetBoolean());
        return document.RootElement.GetProperty("error").GetString() ?? string.Empty;
    }
}
