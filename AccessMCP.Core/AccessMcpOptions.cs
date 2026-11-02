using System.ComponentModel;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>
/// Configuration options for the Access MCP server.
/// </summary>
public sealed record AccessMcpOptions
{
    [Description("Default database path used when a tool call doesn't provide one.")]
    public string? DefaultDatabasePath { get; init; }

    [Description("OLE DB provider name. Typically 'Microsoft.ACE.OLEDB.16.0' or 'Microsoft.ACE.OLEDB.12.0'.")]
    public string OleDbProvider { get; init; } = "Microsoft.ACE.OLEDB.16.0";

    [Description("Command timeout in seconds for Access queries.")]
    public int CommandTimeoutSeconds { get; init; } = 30;

    [Description("If true, only read-only tools are permitted.")]
    public bool ReadOnly { get; init; }

    [Description("If true, enables Access COM automation (forms/controls). Windows only, requires Access installed.")]
    public bool EnableComAutomation { get; init; }

    [Description("If true (the default), COM sessions remove the database's StartUpForm property before opening it and write it back afterwards, so a start-up form can't block automation on a modal dialog. This writes to the database even for read-only tools. Set false to leave the file untouched; a start-up form then runs on every COM session.")]
    public bool SuppressStartupForm { get; init; } = true;

    [Description("Access COM ProgID to instantiate. Default 'Access.Application'.")]
    public string AccessComProgId { get; init; } = "Access.Application";

    [Description("DAO ProgID to instantiate. Default 'DAO.DBEngine.120'.")]
    public string DaoProgId { get; init; } = "DAO.DBEngine.120";

    [Description("Optional path to a .mdw workgroup file for Access user-level security. When set, applied via DBEngine.SystemDB before opening the database, so COM automation logs in as the workgroup's user rather than the default system workgroup.")]
    public string? WorkgroupPath { get; init; }

    [Description("If true, enables the access_vba_run/eval/run_tests/test tools. These execute arbitrary VBA, so the flag is separate from EnableComAutomation and defaults to false.")]
    public bool EnableVbaTestTools { get; init; }

    [Description("Scratch directory for VBA test workspaces. One subdirectory is created per scenario and deleted afterwards unless artifacts are kept. Defaults to %TEMP%\\AccessMCP.VbaTest.")]
    public string? VbaScratchDirectory { get; init; }

    [Description("Directory for the built-MDE cache. Defaults to %LOCALAPPDATA%\\AccessMCP\\VbaTest\\cache.")]
    public string? VbaCacheDirectory { get; init; }

    [Description("Default timeout in seconds for a VBA run/compile. Default 120.")]
    public int VbaTimeoutSeconds { get; init; } = 120;

    [Description("Timeout in seconds for the compile gate. Default 60.")]
    public int VbaCompileTimeoutSeconds { get; init; } = 60;

    [Description("Default MDE build strategy: 'syscmd' (headless SysCmd(603)) or 'prebuilt' (caller supplies an .mde/.accde). Default 'syscmd'.")]
    public string VbaMdeBuildStrategy { get; init; } = "syscmd";

    [Description("If true, run MSACCESS.EXE /compact /decompile on the scratch copy before compiling. Slower and invalidates the MDE cache. Default false.")]
    public bool VbaDecompileFirst { get; init; }

    [Description("If true, cache built MDEs under VbaCacheDirectory. Default true.")]
    public bool VbaCacheEnabled { get; init; } = true;

    [Description("Maximum size of the MDE cache in bytes before least-recently-used entries are removed. Default 2 GiB.")]
    public long VbaCacheMaxBytes { get; init; } = 2L * 1024 * 1024 * 1024;

    [Description("If true, keep the scratch workspace after a successful VBA scenario. Failed scenarios keep it regardless. Default false.")]
    public bool KeepVbaArtifacts { get; init; }

    [Description("Maximum characters of VBA result/log text returned in one MCP response. Default 200000.")]
    public int VbaMaxResultChars { get; init; } = 200_000;

    [Description("How long, in seconds, to wait for Access to close cleanly before killing it. Default 10.")]
    public int VbaCleanupTimeoutSeconds { get; init; } = 10;
}
