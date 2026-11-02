# Changelog

All notable changes to this project are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses
[Semantic Versioning](https://semver.org/). See [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md#changelog) for
how to add entries.

## [Unreleased]

### Added

- Initial public import of the Access MCP server (`DigitalJeeves.MsAccessMCP.exe`) and its tool
  library (`DigitalJeeves.MsAccessMCP.Core`), with 58 tools covering OLE DB data access, Access COM
  automation, VBA source management and the VBA test harness.
- Optional x86 VBE add-in (`tools/AccessMCP.VbeAddin`) and its build, register and sync scripts.
- Consumer README, developer guide and AI agent guidance (`CLAUDE.md`).
- xUnit test project (`AccessMCP.Tests`), ported from the original AccessMCP tests and extended to cover
  path policy, ReadOnly and COM gating for every write tool, export-folder drift classification, the MDE
  cache, scratch workspaces, the OLE DB STA funnel, and OLE DB/DAO tools against a scratch `.accdb`.
- MIT licence (`LICENSE`), also declared in the package metadata.

### Fixed

- `access_list_indexes` returned nothing when given `tableName` (the table name went into the OLE DB
  index-name restriction) and threw on any database with an index (it read a nonexistent `INDEX_TYPE`
  column; ACE calls it `TYPE`). A multi-column index is now listed once instead of once per column.
- `access_vba_cache_clear` now refuses a `key` that isn't a 64-character hex cache key. Before, a key
  such as `..\..\folder` resolved outside the cache directory and that directory was deleted.
- `access_vba_doctor` capabilities now match what the tools allow. `canRun` and `canBuildMde` were true
  with `EnableComAutomation` off, although every run and build refuses without it; `canBuildMde` also
  ignored `EnableVbaTestTools`; and `canEditModules`/`canCompile` were true on a `ReadOnly` server.

### Removed

- `price_lookup` tool and the organisation-specific schema, data access and pricing code behind it.

### Changed

- `access_do_cmd` is now annotated as a write tool (`readOnlyHint` false) and is refused when
  `ReadOnly` is true. Opening a form or report runs its event code, which can change data.
- Renamed from `AccessMCP` to `DigitalJeeves.MsAccessMCP`. The NuGet package, Core assembly and
  namespace are now `DigitalJeeves.MsAccessMCP.Core` (were `AccessMCP.Core`), and the server exe is
  `DigitalJeeves.MsAccessMCP.exe` (was `AccessMCP.exe`). In-process hosts need
  `using DigitalJeeves.MsAccessMCP.Core;`, and MCP client configs must point at the new exe name.
- `AccessTools` constructor is now `AccessTools(AccessMcpOptions options)`; the unused
  `ILoggerFactory` parameter is gone. In-process hosts that construct it by hand must drop that
  argument. Hosts using `WithToolsFromAssembly` are unaffected.
- Sample `config.json` ships with `DefaultDatabasePath: null`; set it to your own database.
