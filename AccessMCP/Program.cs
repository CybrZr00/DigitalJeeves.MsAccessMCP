using DigitalJeeves.MsAccessMCP.Core;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using System.Text.Json;

var builder = Host.CreateApplicationBuilder(args);
builder.Logging.ClearProviders();
builder.Logging.AddConsole(o =>
{
    o.LogToStandardErrorThreshold = LogLevel.Trace;
});

var options = LoadOptions(args);

builder.Services.AddSingleton(options);
builder.Services
    .AddMcpServer(o => o.Filters.Request.CallToolFilters.Add(SurfaceToolExceptions))
    .WithStdioServerTransport()
    // AccessTools lives in DigitalJeeves.MsAccessMCP.Core, a separate assembly from this entry project. WithToolsFromAssembly()
    // with no argument only scans the entry assembly, so point it at AccessTools' assembly explicitly.
    .WithToolsFromAssembly(typeof(AccessTools).Assembly);

await builder.Build().RunAsync();

/// <summary>
/// The MCP SDK turns an unhandled tool exception into "An error occurred invoking '&lt;tool&gt;'." and
/// drops the reason, so a missing file, a bad module name and a COM failure all look identical to the
/// caller and, by omission, become expensive to tell apart. Surface the exception instead, in the same
/// JSON shape the tools use for their own failures.
/// </summary>
static McpRequestHandler<CallToolRequestParams, CallToolResult> SurfaceToolExceptions(
    McpRequestHandler<CallToolRequestParams, CallToolResult> next)
    => async (request, cancellationToken) =>
    {
        try
        {
            return await next(request, cancellationToken);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            var error = ex.GetBaseException();
            var text = JsonSerializer.Serialize(new
            {
                ok = false,
                outcome = "error",
                error = error.Message,
                exceptionType = error.GetType().Name
            }, new JsonSerializerOptions(JsonSerializerDefaults.Web));

            return new CallToolResult
            {
                IsError = true,
                Content = [new TextContentBlock { Text = text }]
            };
        }
    };

/// <summary>
/// Loads MCP configuration from --config, environment, or defaults.
/// </summary>
static AccessMcpOptions LoadOptions(string[] args)
{
    var jsonOptions = new JsonSerializerOptions(JsonSerializerDefaults.Web)
    {
        AllowTrailingCommas = true,
        ReadCommentHandling = JsonCommentHandling.Skip
    };

    string? configPath = null;
    for (var i = 0; i < args.Length; i++)
    {
        if (args[i] is "--config" or "-c")
        {
            if (i + 1 >= args.Length)
            {
                throw new ArgumentException("Missing value for --config.");
            }

            configPath = args[i + 1];
        }
    }

    // Precedence:
    // 1) JSON file via --config
    // 2) env ACCESS_MCP_OPTIONS (JSON)
    // 3) config.json beside this exe (defensive: a launcher that forgot --config
    //    still gets a real DefaultDatabasePath instead of empty results)
    // 4) empty defaults
    if (!string.IsNullOrWhiteSpace(configPath))
    {
        var json = File.ReadAllText(configPath);
        return JsonSerializer.Deserialize<AccessMcpOptions>(json, jsonOptions)
            ?? throw new InvalidOperationException("Failed to parse config JSON.");
    }

    var env = Environment.GetEnvironmentVariable("ACCESS_MCP_OPTIONS");
    if (!string.IsNullOrWhiteSpace(env))
    {
        return JsonSerializer.Deserialize<AccessMcpOptions>(env, jsonOptions)
            ?? throw new InvalidOperationException("Failed to parse ACCESS_MCP_OPTIONS JSON.");
    }

    var beside = Path.Combine(AppContext.BaseDirectory, "config.json");
    if (File.Exists(beside))
    {
        var json = File.ReadAllText(beside);
        return JsonSerializer.Deserialize<AccessMcpOptions>(json, jsonOptions)
            ?? throw new InvalidOperationException("Failed to parse config.json beside the exe.");
    }

    return new AccessMcpOptions();
}
