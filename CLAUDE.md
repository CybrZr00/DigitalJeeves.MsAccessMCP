# CLAUDE.md

Guidance for AI coding agents (Claude Code, Copilot, Cursor and friends) working in this repo. Humans
should read [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) first; the invariants below apply to everyone.

## What this is

A Windows-only MCP stdio server for Microsoft Access databases.

- `AccessMCP/` is the host, `DigitalJeeves.MsAccessMCP.csproj`, building `DigitalJeeves.MsAccessMCP.exe`. `Program.cs` loads
  options and registers every tool in the Core assembly. It should stay thin.
- `AccessMCP.Core/` holds `DigitalJeeves.MsAccessMCP.Core.csproj`, the tool surface and everything behind it:
  - `AccessTools.cs`: data, design-object and VBA source tools
  - `AccessVbaTools.cs`: the `access_vba_*` test harness tools
  - `AccessOleDb.cs` + `OleDbStaThread.cs`: OLE DB, all funnelled through one STA thread
  - `AccessDao.cs`: DAO table design, no Access process
  - `AccessComAutomation.cs`: Access COM sessions
  - `Vba*.cs`, `MdeBuilder.cs`, `ModalDialogWatcher.cs`, `NdjsonResults.cs`: harness internals
  - `AccessMcpOptions.cs`: every config option
  - `AccessDbPathPolicy.cs`: path normalisation and extension checks
- `tools/AccessMCP.VbeAddin/`: optional x86 C++ VBE add-in for exact compile diagnostics.
- `scripts/`: PowerShell for building/registering the add-in and syncing embedded resources.

## Build

```powershell
dotnet build DigitalJeeves.MsAccessMCP.slnx -c Debug -p:Platform=x64
```

There is no test project yet. A build with zero warnings is the bar; don't introduce new ones.

## Rules for every change

1. **Keep CHANGELOG.md current.** Every user-visible change (new or removed tool, renamed parameter,
   changed output shape, new option, changed default, behaviour fix) gets a line under
   `## [Unreleased]` in the right Keep a Changelog section, in the same commit as the change. Internal
   refactors with no visible effect don't need an entry.
2. **No organisation-specific code or data.** This repo is a generic, public-facing MCP. No company
   names, internal hostnames, UNC shares, private NuGet feeds, customer schemas, real database files or
   credentials. Examples use `C:\Data\Sample.accdb`. Secrets never go in `config.json` or any tracked
   file.
3. **Tool names and parameters are a public contract.** Renaming or removing a tool, or a parameter,
   breaks every client and every in-process host. Add new optional parameters instead, and record
   breaking changes in the changelog under a `Changed` or `Removed` heading.
4. **Bump `<Version>` in `DigitalJeeves.MsAccessMCP.Core.csproj`** when releasing a package, or NuGet caches keep
   serving the old build. Package ID, assembly and root namespace are all
   `DigitalJeeves.MsAccessMCP.Core`; only the folder is still called `AccessMCP.Core`.
5. Generated text in code, JSON, config and commit messages is ASCII only.

## Searching the codebase

Prefer the IDE++ embeddings index when it exists and is fresh:

- File: `.ideplusplus/copilot/embeddings.json` (gitignored, local to each clone).
- Shape: `{ "ModelId": "onnx:bge-small-en-v1.5-int8", "Chunks": [ { "FilePath", "StartLine",
  "EndLine", "Text", "Embedding": [384 floats] } ] }`. `FilePath` is absolute.
- **Check freshness before trusting it.** It is only rebuilt when the maintainer remembers to. Treat it
  as stale if its modified time is older than the newest commit (`git log -1 --format=%cI`), or if any
  sampled `FilePath` no longer exists. When stale, say so and fall back to Grep/Glob; never cite a
  chunk's line numbers without opening the file.
- Query embeddings must come from the same model (`BAAI/bge-small-en-v1.5`) for cosine similarity to
  mean anything. Without it, the chunk `Text` and line ranges are still a usable map of the code.

## Invariants (read before touching COM, OLE DB or VBA code)

- **OLE DB parameters are positional.** `AccessOleDb.AddParameters()` adds in dictionary order; callers
  must match placeholder order in the SQL.
- **All OLE DB work goes through `OleDbStaThread`.** ACE/Jet is apartment-threaded COM; running it on
  thread-pool threads corrupts native state. New OLE DB code must use the same funnel.
- **`ReadOnly` is enforced per tool.** Each write tool checks `AccessMcpOptions.ReadOnly` itself. A new
  write tool must add the check; nothing does it centrally.
- **COM runs in an STA via `RunInAccessAsync` / `RunInAccessWithWatchdogAsync`** and must release every
  COM object it touches (`VBProjects`, `VBComponents`, `CodeModule`, controls).
- **Both runners hold one process-wide `AutomationGate` (`SemaphoreSlim(1, 1)`).** Never call one runner
  from inside another's callback: it deadlocks. Open databases one after the other instead, as
  `access_diff_vba_modules` does. Always `WaitAsync`, never a synchronous `Wait`, or UI-thread hosts
  freeze.
- **`StartUpForm` round-trip.** Unless `SuppressStartupForm` is false, both runners delete the target's
  `StartUpForm` property before `OpenCurrentDatabase` and restore it after. `StartUpFormRecovery`
  records the value until the restore succeeds. Keep the delete/restore pair tight. This is the one
  deliberate write to a caller's database.
- **The harness never modifies the caller's database.** It works on `VbaWorkspace` copies and cleans up
  unless artifacts are kept.
- **Module writes are body-only.** Never write `Attribute VB_Name` or `VERSION`/`BEGIN`/`END` into a
  `CodeModule`. Use `VbaProjectEditor.EnsureModule` or `VBComponents.Import`.
- **`Application.Run` takes unqualified names only** and can raise modal dialogs; the generated test
  runner uses direct `Call TestName` statements.
- **`SysCmd(603, source, target)` must run early-bound on the same Access instance** that compiled and
  closed the source, after releasing all VBE references. A fresh instance or late-bound call silently
  does nothing. `MdeBuilder` resolves the Office PIAs copied from the GAC at build time.
- **Compile-error text comes from `Resources/VbeCompileErrors.json`** via `VbeCompileErrorTable`. Never
  hardcode error codes in `VbaCompiler`.
- **JSON output** uses `JsonSerializerDefaults.Web`, indented. Keep new tools consistent with existing
  result shapes.
