# Development guide

## Layout

```
DigitalJeeves.MsAccessMCP.slnx
AccessMCP/                 stdio MCP host (DigitalJeeves.MsAccessMCP.exe), Program.cs + sample config.json
AccessMCP.Core/            DigitalJeeves.MsAccessMCP.Core: tool surface and the OLE DB / DAO / COM / VBA harness behind it
  Resources/               embedded VbaTestKit.bas and VbeCompileErrors.json
AccessMCP.Tests/           DigitalJeeves.MsAccessMCP.Tests: xUnit tests, no Access process needed
tools/AccessMCP.VbeAddin/  optional x86 C++ VBE add-in for exact compile diagnostics
scripts/                   add-in build/register/test, resource sync scripts
docs/                      this guide, the VBA harness reference, spike notes
CHANGELOG.md               user-visible changes, Keep a Changelog format
CLAUDE.md                  AI agent guidance and the code invariants
```

## Prerequisites

- Windows, .NET 10 SDK
- Access Database Engine (ACE OLE DB provider)
- Microsoft Access for anything COM, VBE or VBA. Building on a machine with Office installed also copies
  the `office`, `Microsoft.Office.Interop.Access.Dao` and `Microsoft.Vbe.Interop` PIAs from the GAC into
  the output; without them `access_vba_build_mde` reports unavailable instead of failing the build.
- For the VBE add-in only: Visual Studio C++ build tools (x86).

## Build, run, publish

```powershell
dotnet build DigitalJeeves.MsAccessMCP.slnx -c Debug -p:Platform=x64
```

Run the tests:

```powershell
dotnet test DigitalJeeves.MsAccessMCP.slnx -c Debug -p:Platform=x64
```

Most tests are pure logic. The `[AceFact]` ones in `AccessOleDbIntegrationTests` build a scratch
`.accdb` in `%TEMP%` through ADOX and run the OLE DB and DAO tools against it; they skip when the ACE
provider isn't installed at the test process's bitness (x64 ACE needs the x64 build above). Nothing
starts Access, so COM, VBE and harness runs are still checked by hand.

Run the host against a config file:

```powershell
AccessMCP\bin\x64\Debug\net10.0-windows10.0.19041.0\DigitalJeeves.MsAccessMCP.exe --config C:\Data\config.json
```

It waits silently on stdin for an MCP client. To poke at it by hand, the MCP Inspector works well:

```powershell
npx @modelcontextprotocol/inspector AccessMCP\bin\x64\Debug\net10.0-windows10.0.19041.0\DigitalJeeves.MsAccessMCP.exe --config C:\Data\config.json
```

Publish for distribution:

```powershell
dotnet publish AccessMCP/DigitalJeeves.MsAccessMCP.csproj -c Release -r win-x64 --self-contained false
```

Never commit a real database. `.gitignore` excludes `*.mdb`, `*.accdb` and friends; keep test
databases outside the repo and point `DefaultDatabasePath` at them from a local config.

## Adding or changing a tool

1. Add a method to `AccessTools` (or `AccessVbaTools` for harness work) with `[McpServerTool(Name =
   "access_...", Title = "...", ReadOnly = ..., Idempotent = ...)]` and a `[Description]` on the method
   and each parameter. Descriptions are what the model reads to decide how to call the tool, so be
   specific about inputs, side effects and output.
2. Normalise the path with `AccessDbPathPolicy.NormalizeAndValidate` (or `NormalizeComDatabasePath` for
   tools that accept `.mde`/`.accde`).
3. If it writes, check `_options.ReadOnly` and refuse, and add the tool to the ReadOnly table in
   `AccessMCP.Tests/AccessToolsGatingTests.cs` (a coverage test fails until you do). If it needs
   Access, go through the COM runners.
4. Return indented JSON built with the shared `JsonOptions`.
5. Add tests for any logic that runs without Access.
6. Add a line to `CHANGELOG.md` under `[Unreleased]`.
7. Add the tool to the table in `README.md`.

The invariants in [CLAUDE.md](../CLAUDE.md#invariants-read-before-touching-com-ole-db-or-vba-code)
cover threading, COM lifetime, the automation gate and the VBA harness rules. Read them before touching
those areas; most of them were learned the hard way.

## VBE add-in

Optional, x86 only, registered per user:

```powershell
pwsh ./tools/AccessMCP.VbeAddin/build-vbe-addin.ps1
pwsh ./scripts/register-vbe-addin.ps1
pwsh ./scripts/test-vbe-addin.ps1
```

Embedded resources have sync scripts with a `-Check` mode:

```powershell
pwsh ./scripts/sync-vbe-errors.ps1 -Check     # VbeCompileErrors.json against ~/vbe7
pwsh ./scripts/sync-vba-testkit.ps1 -Check    # VbaTestKit.bas against VbaTestHost
```

Both default to source paths on the original author's machine; pass `-Source` when yours differ.
Details are in [VBA-TEST-HARNESS.md](VBA-TEST-HARNESS.md#6-compile-diagnostics-vbe-add-in).

## Changelog

`CHANGELOG.md` follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and
[Semantic Versioning](https://semver.org/).

- Add entries under `## [Unreleased]` in the same commit as the change, grouped as `Added`, `Changed`,
  `Deprecated`, `Removed`, `Fixed`, `Security`.
- Write for someone calling the tools: name the tool or option and say what's different for them.
- On release, rename `[Unreleased]` to `[x.y.z] - YYYY-MM-DD`, bump `<Version>` in
  `DigitalJeeves.MsAccessMCP.Core.csproj` to match, and start a fresh `[Unreleased]` section.
- Removing or renaming a tool or parameter is a major version bump.

## Codebase search with IDE++ embeddings

The IDE++ extension can index the solution into `.ideplusplus/copilot/embeddings.json` (model
`bge-small-en-v1.5`, int8 ONNX). Its chat uses the index for semantic search, and `CLAUDE.md` tells AI
agents to use it as well, but only when it's fresh.

- Rebuild it after any large change: files added, moved or deleted, or a big refactor. A stale index
  sends searches to code that no longer exists.
- It's gitignored because it stores absolute paths and verbatim code chunks. If you want to share one,
  rebuild it from a clean checkout first and remove `.ideplusplus/` from `.gitignore` on purpose.

## Security notes

- There is no database path allow-list. Whoever controls the MCP client can open any database the
  server's account can read. Deploy it under an account scoped to the databases you mean to expose.
- `EnableVbaTestTools` runs arbitrary VBA with the user's privileges. It defaults to off; keep it that
  way outside development machines.
- `SuppressStartupForm` writes to the database even for read-only tools. Document this for anyone who
  expects `ReadOnly: true` to mean "never touches the file".
