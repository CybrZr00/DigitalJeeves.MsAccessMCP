using DigitalJeeves.MsAccessMCP.Core;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
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
    .AddMcpServer()
    .WithStdioServerTransport()
    // AccessTools lives in DigitalJeeves.MsAccessMCP.Core, a separate assembly from this entry project. WithToolsFromAssembly()
    // with no argument only scans the entry assembly, so point it at AccessTools' assembly explicitly.
    .WithToolsFromAssembly(typeof(AccessTools).Assembly);

await builder.Build().RunAsync();

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
