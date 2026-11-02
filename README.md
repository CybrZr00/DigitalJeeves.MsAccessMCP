# DigitalJeeves.MsAccessMCP

A Windows-only [Model Context Protocol](https://modelcontextprotocol.io) server that lets an MCP client
(Claude Code, Claude Desktop, VS Code, or anything else that speaks MCP over stdio) work with Microsoft
Access `.mdb` / `.accdb` databases.

It covers three areas:

- **Data**: schema, tables, indexes, relationships, SELECT queries, paged queries, exports, and
  non-query SQL through the ACE OLE DB provider.
- **Design objects**: forms, reports, controls, macros and saved queries through Access COM
  automation.
- **VBA**: read, export, import, diff and compile modules, plus a test harness that runs VBA and
  `'@test` procedures against a scratch copy of the database, optionally as a built `.mde`/`.accde`.

The server process is `DigitalJeeves.MsAccessMCP.exe`. The tool surface itself lives in the `DigitalJeeves.MsAccessMCP.Core` library,
so a .NET host can also load it in-process (see [Embedding](#embedding-the-core-library-in-process)).

> This project was written with heavy AI assistance. Review it before production use, and start
> against copies of your databases.

## Requirements

- Windows
- .NET 10 runtime (SDK to build)
- Microsoft Access Database Engine, for the ACE OLE DB provider. The provider is usually
  `Microsoft.ACE.OLEDB.16.0`; some machines only have `Microsoft.ACE.OLEDB.12.0`.
- Microsoft Access, for every COM, VBE and VBA tool. The server's bitness must match Access: publish
  `win-x64` for 64-bit Office, `win-x86` for 32-bit.
- For the VBE tools: Access Trust Center, "Trust access to the VBA project object model".

## Install

Build and publish from source:

```powershell
dotnet publish AccessMCP/DigitalJeeves.MsAccessMCP.csproj -c Release -r win-x64 --self-contained false
```

The exe lands in `AccessMCP/bin/Release/net10.0-windows10.0.19041.0/win-x64/publish/`. Copy that folder
wherever you like; `config.json` is published next to the exe.

## Configure

Options are read from the first of these that exists:

1. `--config <path>` (or `-c <path>`), a JSON file
2. the `ACCESS_MCP_OPTIONS` environment variable, holding JSON
3. `config.json` beside `DigitalJeeves.MsAccessMCP.exe`
4. built-in defaults

```json
{
  "DefaultDatabasePath": "C:\\Data\\Sample.accdb",
  "OleDbProvider": "Microsoft.ACE.OLEDB.16.0",
  "CommandTimeoutSeconds": 30,
  "ReadOnly": true,
  "EnableComAutomation": true,
  "EnableVbaTestTools": false
}
```

| Option | Default | What it does |
|---|---|---|
| `DefaultDatabasePath` | none | Used when a tool call omits `databasePath`. Environment variables are expanded. |
| `OleDbProvider` | `Microsoft.ACE.OLEDB.16.0` | ACE provider name. |
| `CommandTimeoutSeconds` | `30` | OLE DB command timeout. |
| `ReadOnly` | `false` | Rejects every tool that writes: `access_execute`, form/report/control setters, module imports, compile, macros, VBA harness writers. |
| `EnableComAutomation` | `false` | Required for every form, report, macro, VBE and VBA tool. Launches Access. |
| `SuppressStartupForm` | `true` | COM sessions remove the database's `StartUpForm` property while they work and put it back afterwards, so a start-up form can't block automation. This writes to the file even for read-only tools; set `false` to never touch it. |
| `EnableVbaTestTools` | `false` | Enables `access_vba_run`, `_eval`, `_run_tests`, `_test`, `_build_mde`. These run arbitrary VBA. |
| `AccessComProgId` / `DaoProgId` | `Access.Application` / `DAO.DBEngine.120` | COM ProgIDs. |
| `WorkgroupPath` | none | `.mdw` workgroup file for user-level security. |
| `Vba*` | see `AccessMcpOptions.cs` | Scratch and cache directories, timeouts, MDE build strategy, cache size, result size cap. |

There is no path allow-list: any `.mdb`/`.accdb` the server's account can read can be opened. Run it
under an account that only sees the databases you mean to expose, and use `ReadOnly: true` unless the
client needs to write.

## Connect a client

The server speaks MCP over stdin/stdout, so the client starts it. Paths below are examples.

**Claude Code**

```powershell
claude mcp add access -- "C:\Tools\AccessMCP\DigitalJeeves.MsAccessMCP.exe" --config "C:\Tools\AccessMCP\config.json"
```

**Claude Desktop** (`claude_desktop_config.json`) or a project `.mcp.json`:

```json
{
  "mcpServers": {
    "access": {
      "command": "C:\\Tools\\AccessMCP\\DigitalJeeves.MsAccessMCP.exe",
      "args": ["--config", "C:\\Tools\\AccessMCP\\config.json"]
    }
  }
}
```

**VS Code** (`.vscode/mcp.json`):

```json
{
  "servers": {
    "access": {
      "type": "stdio",
      "command": "C:\\Tools\\AccessMCP\\DigitalJeeves.MsAccessMCP.exe",
      "args": ["--config", "C:\\Tools\\AccessMCP\\config.json"]
    }
  }
}
```

Logs go to stderr, which keeps stdout clean for JSON-RPC.

## Tools

Every tool takes an optional `databasePath`; without it the server uses `DefaultDatabasePath`. Results
are indented JSON.

### Data (OLE DB, no Access process)

| Tool | Purpose |
|---|---|
| `access_list_tables`, `access_list_views`, `access_list_queries` | List user tables, views, stored queries. |
| `access_list_indexes`, `access_list_relationships` | Indexes and relationships, optionally filtered by table. |
| `access_get_schema` | Tables with their columns. |
| `access_get_table_design` | Field and index design metadata via DAO. |
| `access_query`, `access_query_paged`, `access_scalar` | Run a SELECT and get rows, a page of rows, or one value. |
| `access_export_table` | Export a table as JSON or CSV. |
| `access_execute` | INSERT / UPDATE / DELETE / DDL. Blocked by `ReadOnly`. |

OLE DB parameters are positional even when named: pass them in the same order as the placeholders in
the SQL.

### Forms, reports, macros, queries (COM, needs `EnableComAutomation`)

| Tool | Purpose |
|---|---|
| `access_list_forms`, `access_list_reports`, `access_list_macros` | List design objects. |
| `access_get_form_controls`, `access_get_report_controls` | List controls. |
| `access_add_form_control`, `access_remove_form_control`, `access_add_report_control`, `access_remove_report_control` | Change controls. |
| `access_set_form_control_value` | Set a control value on an open form. |
| `access_get_object_properties`, `access_set_object_properties` | Form/report properties. |
| `access_get_control_properties`, `access_set_control_properties` | Control properties. |
| `access_run_macro`, `access_do_cmd` | Run a macro; run a limited set of `DoCmd` actions. |
| `access_execute_saved_query`, `access_get_query_sql` | Run a saved query; read its SQL without running it. |

### VBA source (COM / VBE, needs `EnableComAutomation`)

| Tool | Purpose |
|---|---|
| `access_list_modules`, `access_list_module_procedures`, `access_get_module_code` | Read modules. |
| `access_vbe_list_projects`, `access_vbe_list_modules`, `access_vbe_get_module_code`, `access_vbe_get_proc_info` | Same through the VBE object model. |
| `access_export_module_code`, `access_import_module_code` | Export or replace one module. Import backs up the old code and can refuse if the module changed underneath you. |
| `access_export_all_modules_to_folder` | Export every module (and optionally saved queries) to a folder with a manifest. |
| `access_import_changed_modules_from_folder` | Push back the files edited since that export. |
| `access_check_export_drift`, `access_pull_drifted_modules` | Find, then re-export, modules changed inside Access since the export. |
| `access_diff_vba_modules` | Diff modules (and optionally saved queries) between two databases. |
| `access_compile_vba` | Compile and save, reporting the real `IsCompiled` result and the first error. |

### VBA test harness

These work on a scratch copy and never modify the configured database. Execution tools also need
`EnableVbaTestTools`.

| Tool | Purpose |
|---|---|
| `access_vba_doctor` | Report what will work on this machine: Access version and bitness, VBE trust, provider, PIAs, MDE build support. Start here. |
| `access_vba_ensure_module`, `access_vba_remove_module` | Create/replace or remove a module with a verified write. |
| `access_vba_import_file`, `access_vba_export_file` | `.bas` / `.cls` round trips with headers preserved. |
| `access_vba_compile` | Compile with diagnostics. |
| `access_vba_run`, `access_vba_eval` | Run a public procedure or evaluate an expression. |
| `access_vba_run_tests` | Discover and run `'@test` procedures. |
| `access_vba_build_mde` | Build an `.mde` / `.accde`. |
| `access_vba_test` | End to end: copy, inject, compile, optionally build an MDE, run, report. |
| `access_vba_cache_info`, `access_vba_cache_clear` | Inspect or clear the built-MDE cache. |

[docs/VBA-TEST-HARNESS.md](docs/VBA-TEST-HARNESS.md) has the full reference, workflows and limitations.
Exact compile-error line numbers need the optional x86 VBE add-in in `tools/AccessMCP.VbeAddin`; without
it, compile falls back to best-effort diagnostics.

## Embedding the Core library in-process

A .NET host targeting `net10.0-windows10.0.19041.0` or later can reference the
`DigitalJeeves.MsAccessMCP.Core` package (or project) and skip the subprocess. The namespace is
`DigitalJeeves.MsAccessMCP.Core`:

```csharp
var options = new AccessMcpOptions { DefaultDatabasePath = @"C:\Data\Sample.accdb", ReadOnly = true };

// As an MCP server inside your own host:
services.AddSingleton(options);
services.AddMcpServer()
    .WithStdioServerTransport()
    .WithToolsFromAssembly(typeof(AccessTools).Assembly);

// Or call the tools directly:
var tools = new AccessTools(options);
string json = await tools.ListTablesAsync();
```

COM sessions are serialized per process, so two concurrent COM tool calls queue rather than run in
parallel. All OLE DB work runs on one dedicated STA thread for the same reason.

## Troubleshooting

- **"Provider not registered" or connection failures**: install the Access Database Engine at the same
  bitness as the server, or set `OleDbProvider` to `Microsoft.ACE.OLEDB.12.0`.
- **COM tools fail immediately**: `EnableComAutomation` is off, Access isn't installed, or the server's
  bitness doesn't match Office.
- **"Programmatic access to Visual Basic Project is not trusted"**: enable the Trust Center setting
  listed under Requirements.
- **Mapped drives**: the account running the MCP client must see the drive. UNC paths are more
  reliable than drive letters.

Run `access_vba_doctor` for a capability report before digging further.

## Contributing

See [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md). Record user-visible changes in
[CHANGELOG.md](CHANGELOG.md).

## License

[MIT](LICENSE). Copyright (c) 2026 Digital Jeeves.
