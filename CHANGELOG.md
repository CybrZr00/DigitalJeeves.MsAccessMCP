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
- MIT licence (`LICENSE`), also declared in the package metadata.

### Removed

- `price_lookup` tool and the organisation-specific schema, data access and pricing code behind it.

### Changed

- Renamed from `AccessMCP` to `DigitalJeeves.MsAccessMCP`. The NuGet package, Core assembly and
  namespace are now `DigitalJeeves.MsAccessMCP.Core` (were `AccessMCP.Core`), and the server exe is
  `DigitalJeeves.MsAccessMCP.exe` (was `AccessMCP.exe`). In-process hosts need
  `using DigitalJeeves.MsAccessMCP.Core;`, and MCP client configs must point at the new exe name.
- `AccessTools` constructor is now `AccessTools(AccessMcpOptions options)`; the unused
  `ILoggerFactory` parameter is gone. In-process hosts that construct it by hand must drop that
  argument. Hosts using `WithToolsFromAssembly` are unaffected.
- Sample `config.json` ships with `DefaultDatabasePath: null`; set it to your own database.
