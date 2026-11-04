# VBA test harness for Access

**Status:** implemented (2026-10-18)
**Tools:** `access_vba_*` in `AccessMCP.Core/AccessVbaTools.cs`
**Primary consumer:** VBAStack, especially its compiled-MDE code path

This harness compiles, runs and regression-tests VBA **inside a real Microsoft Access process**,
against a scratch copy of an `.mdb`/`.accdb` and against a compiled `.mde`/`.accde` built by the
harness. It is complementary to `vba-MCP`: that tool runs VBA in a standalone VBE host with no host
object model; this one has the full Access object model and can build MDEs.

The developer's database is not modified by default: every scenario copies it to a scratch workspace
first. `access_vba_run_tests`, `access_vba_build_mde` and `access_vba_test` also take
`inPlace: true`, which works on the database itself so changes persist; see [5. Security](#5-security).

---

## 1. Requirements

| Requirement | Why |
|---|---|
| Microsoft Access **full** (not Runtime), x86 or x64 | The Runtime has no VBE, so modules cannot be edited, compiled or baked into an MDE |
| Trust Center → **Trust access to the VBA project object model** | `app.VBE` is used for module reads/writes and discovery |
| `EnableComAutomation=true` | All `access_vba_*` tools use COM |
| `EnableVbaTestTools=true` | Separate, default-off gate for the tools that execute arbitrary VBA (`run`, `eval`, `run_tests`, `test`, `build_mde`) |
| Office PIAs | `Microsoft.Office.Interop.Access` comes from NuGet; `office.dll`, `Microsoft.Office.Interop.Access.Dao.dll` and `Microsoft.Vbe.Interop.dll` are referenced from the GAC and copied next to the host at build time |
| VBE add-in (optional, exact compile diagnostics) | `tools/AccessMCP.VbeAddin` + `scripts/register-vbe-addin.ps1`; x86 Access only, with fallback when absent |
| Matching bitness for OLE DB | The ACE provider is in-process; a 64-bit host needs the 64-bit ACE provider. Access automation itself is out-of-process and works cross-bitness (verified x64 host → x86 Access) |

Run **`access_vba_doctor`** first. It reports Access version/build/bitness/runtime, VBE trust, the
ACE provider, the scratch/cache directories, PIA availability and `canBuildMde` without throwing.

---

## 2. Configuration

```json
{
  "DefaultDatabasePath": "C:\\data\\app.accdb",
  "EnableComAutomation": true,
  "EnableVbaTestTools": true,
  "ReadOnly": false,

  "VbaScratchDirectory": "%TEMP%\\AccessMCP.VbaTest",
  "VbaCacheDirectory": "%LOCALAPPDATA%\\AccessMCP\\VbaTest\\cache",
  "VbaTimeoutSeconds": 120,
  "VbaCompileTimeoutSeconds": 60,
  "VbaMdeBuildStrategy": "syscmd",
  "VbaDecompileFirst": false,
  "VbaCacheEnabled": true,
  "VbaCacheMaxBytes": 2147483648,
  "KeepVbaArtifacts": false,
  "VbaMaxResultChars": 200000,
  "VbaCleanupTimeoutSeconds": 10
}
```

- `EnableVbaTestTools` defaults to **false**. `access_vba_doctor`, `access_vba_export_file`,
  `ensure_module`, `remove_module`, `import_file` and `compile` work with just
  `EnableComAutomation`; the execution tools refuse until the flag is set.
- `ReadOnly=true` blocks tools that write to the **configured** database. `access_vba_test` and the
  other scenario tools only write to their scratch copy, so they still work.
- `VbaCacheDirectory` is for built MDEs only. The prepared MDB is always materialised per run, so a
  stale fixture can never be served.
- `rebuild=true` on a call forces a build but still refreshes the cached entry.

---

## 3. Tool reference

### `access_vba_doctor(databasePath?)`

Returns `access`, `host`, `database`, `vbe`, `ace`, `scratch`, `cache` and `capabilities`
(`canEditModules`, `canCompile`, `canRun`, `canBuildMde`, `piaAvailable`, `mdeStrategy`). Works with
COM disabled.

The capability flags follow the same switches the tools check. `canEditModules` and `canCompile` need
`EnableComAutomation`, a trusted VBE and a server that isn't `ReadOnly`. `canRun` needs both
`EnableComAutomation` and `EnableVbaTestTools`. `canBuildMde` needs what `canRun` needs plus the PIAs
and a full (non-runtime) Access; `access_vba_build_mde` is still refused when `ReadOnly`, but
`access_vba_test` with `mode="mde"` builds in its scratch copy and works.

### `access_vba_ensure_module`

```
moduleName, code? | codeFile?, kind="standard"|"class",
compile=true, databasePath?, timeoutSeconds?
→ { ok, moduleName, kind, created, lines, sha256, backupPath?, compile }
```

Creates the module with `VBComponents.Add` if missing, names it, then writes the body through a
verified line-level writer. Existing modules are backed up and rolled back on any read-back
mismatch. **Never uses `AddFromString`.** Export-only headers (`Attribute VB_Name`,
`VERSION 1.0 CLASS`, `BEGIN … END`) are stripped before writing; existing `Option` lines are
preserved unless the supplied code provides its own.

> A module body written through `CodeModule` must **not** contain `Attribute VB_Name`; it is a
> syntax error there. Use `access_vba_import_file` if you want the file's headers preserved.

### `access_vba_remove_module`

Refuses form/report code-behind. By default only modules whose names start with `AccessMcp` may be
removed (`onlyHarnessModules=false` to override).

### `access_vba_import_file` / `access_vba_export_file`

`import_file` uses `VBComponents.Import` (`.bas`/`.cls`; forms need a designer and are rejected) and
preserves the file's `Attribute VB_Name`. `export_file` uses `VBComponent.Export`; it only writes a
file, never the database.

### `access_vba_compile`

```
databasePath?, timeoutSeconds?
→ { ok, isCompiled, module?, line?, message?, dialogText?, brokenReferenceFlag, brokenReferences }
```

Compiles with `acCmdCompileAndSaveAllModules` (126) and uses **`Application.IsCompiled`** as the
gate. If false, it opens a code context and retries once (AutoBuilder's nudge). Broken references
are reported from `Application.BrokenReference`; only that flag aborts the compile. The caret is a
best-effort hint: Access does not expose a structured compiler API, and the caret is often parked
past the real error, so `module`/`line` may be null.

`access_compile_vba` (the older tool) now uses the same gate and returns the same diagnostic shape
instead of an unconditional `compiled: true`.

### `access_vba_run`

```
procedure, args?, timeoutSeconds?, databasePath?
→ { ok, procedure, elapsedMs, result: { type, text, json? }, error, crashed, exitCode, timedOut,
    logTail, dialogText, artifactsKept, artifactDirectory }
```

Runs in a scratch copy. Access's `Application.Run` accepts **unqualified** procedure names;
`Module.Procedure` is accepted by the API and the module part is used for validation, but the call
is made unqualified. `ByRef` parameters are inbound-only: COM does not copy changes back.

### `access_vba_eval`

`Application.Eval` in a scratch copy, same result shape.

### `access_vba_run_tests`

```
tests?: string[], timeoutSeconds?, keepArtifacts?, databasePath?, inPlace=false
→ { ok, total, passed, failed, skipped, xfailed, xpassed, tests: [...], logTail, error,
    timedOut, crashed, exitCode, elapsedMs, dialogText, artifactsKept, artifactDirectory }
```

MDB only. Discovers public `'@test` procedures through the VBE, injects `VbaTestKit`, generates
`AccessMcpTestRunner` (direct `Call TestName` statements, one error handler per test), compiles and
runs it. Every test start/outcome is appended to an NDJSON file as it happens, so a crash or timeout
still returns the results written so far.

Supported directives: `'@test`, `'@test-skip`, `'@test-xfail`.
- `fail` and `error` count as failures; `skip` is skipped.
- `xfail` is not a failure; `xpass` (an expected failure that passed) **is** a failure.
- Test procedure names must be unique project-wide (Access resolves calls by name).

The canonical kit lives at `AccessMCP.Core/Resources/VbaTestKit.bas` and is byte-identical to the
`VbaTestHost` template. Sync/check it with:

```powershell
pwsh ./scripts/sync-vba-testkit.ps1          # write
pwsh ./scripts/sync-vba-testkit.ps1 -Check   # CI drift check
```

### `access_vba_build_mde`

```
databasePath, outputPath?, strategy="syscmd"|"prebuilt", decompileFirst?, timeoutSeconds?, inPlace=false
→ { ok, strategy, sourcePath, outputPath, verified, manualRequired, error }
```

Builds from a scratch copy; the source is untouched. `inPlace: true` compiles the database itself
(the output still goes to `outputPath`). `syscmd` compiles the copy, closes it, and calls
`SysCmd(603, source, target)` **on the same Access instance** (see 7). `prebuilt` copies an existing
`.mde`/`.accde` instead. The output path is used verbatim; the default is the source name with
`.mde`/`.accde` beside the source.

### `access_vba_cache_info` / `access_vba_cache_clear`

Inspect or remove built-MDE cache entries.

### `access_vba_test`

```
databasePath,
modules: [{ name, kind, source? | filePath? }],
entryPoint?, args?, mode="mdb"|"mde",
mdeStrategy?, decompileFirst?, useCache=true, rebuild=false,
continueOnCompileError=false, keepArtifacts=false, timeoutSeconds?, inPlace=false
→ { ok, mode, sourceDatabasePath, workDatabasePath, mdePath?, prepare, run, tests, artifacts… }
```

The end-to-end entry point. If `entryPoint` is omitted and `'@test` procedures are found, the
generated runner is used. `mode=mdb` runs the entry point in the prepared MDB; `mode=mde` builds (or
reuses a cached) MDE and runs it. `inPlace=true` prepares the database itself; modules you supply
persist, and the harness removes the kit/runner it injected before returning.

Example:

```json
{
  "databasePath": "C:\\data\\sample_form.mdb",
  "modules": [
    { "name": "VBAStack", "kind": "standard", "filePath": "C:\\src\\VBAStack.bas" },
    { "name": "AccessMcpVbaStackSelfTest", "kind": "standard", "filePath": "C:\\src\\AccessMcpVbaStackSelfTest.bas" }
  ],
  "mode": "mde",
  "mdeStrategy": "syscmd",
  "keepArtifacts": false,
  "timeoutSeconds": 180
}
```

---

## 4. Workflows

### Run a test suite in an MDB

```text
access_vba_test(databasePath=app.accdb,
                modules=[{name: MyTests, kind: standard, filePath: MyTests.bas}],
                mode="mdb")
```

### Run the same suite in a freshly built MDE

Use `mode="mde"`, `mdeStrategy="syscmd"`. The kit and runner are injected into the MDB **before**
the build, so the compiled database contains the runner. The second call with `useCache=true`
reuses the built MDE unless any module, the kit, the Access build or the source identity changed.

### Debug a failed scenario

`keepArtifacts=true` (or a failed run) leaves the workspace with `work.accdb`, `out.mde`/`out.accde`,
`result.ndjson` and `run.log`. The response includes the directory.

### Build an MDE without running tests

`access_vba_build_mde(databasePath, outputPath)`. It compiles the copy first and refuses to build a
project that does not compile.

---

## 5. Security

Executing VBA is equivalent to code execution with the user's privileges. There is no sandbox.

- The execution tools are behind `EnableVbaTestTools`, default **false**, independent of
  `EnableComAutomation`.
- `access_vba_run` and `access_vba_eval` always work on a scratch copy: they persist nothing, so
  running them in place would only mean executing with real side effects. `access_vba_run_tests`,
  `access_vba_build_mde` and `access_vba_test` accept `inPlace: true`, which works on the database
  itself: changes persist, nothing is rolled back, and `access_vba_run_tests` / `access_vba_test` remove their
  injected `AccessMcpTestKit` / `AccessMcpTestRunner` modules afterwards. If that cleanup fails it is
  logged in the artifact directory rather than thrown. The artifact directory is always kept for
  an in-place scenario run, because it holds the backups of any module the run overwrites in your
  database. Use in-place deliberately, not as a default.
- `AutomationSecurity` is set to `msoAutomationSecurityLow` for the scratch session. The database's
  own startup code, `AutoExec` macro and startup form can run; the StartUpForm property is removed
  from the **MDB copy** before compiling/building and restored only in the copy.
- Module writes refuse form/report code-behind; removal defaults to `AccessMcp*` modules only.
- Result and log text is capped by `VbaMaxResultChars`.
- Cached MDEs contain compiled code; treat the cache as executable content.

---

## 6. Compile diagnostics: VBE add-in

Access's automation API exposes only `Application.IsCompiled`; the compiler's own error record is
in-process state in `VBE7.DLL`. AccessMCP therefore ships a small native **x86 VBE add-in** that runs
inside Access, reads the record immediately after the normal compile command, and returns it through
`Application.VBE.AddIns("AccessMCP.VbeProbe").Object`.

### Build, register, test

```powershell
pwsh ./tools/AccessMCP.VbeAddin/build-vbe-addin.ps1   # x86 native DLL
pwsh ./scripts/register-vbe-addin.ps1                 # per-user HKCU, no admin
pwsh ./scripts/test-vbe-addin.ps1                     # deliberate compile error, asserts line/module
```

The add-in is built separately from the .NET solution because it must match **Access's** bitness
(x86 here), not the MCP host. It is not part of the `AccessMCP` publish output; install it once per
user with the registration script.

### How it works

1. `VbeAddinClient.TryConnect` calls `VBE.AddIns.Update()`, finds `AccessMCP.VbeProbe`, sets
   `Connect = true`, and gets `AddIn.Object`.
2. `VbaCompiler` runs `DoCmd.RunCommand(126)` exactly as before. Access sets the VBE's
   `BTSRC_ERR` record and leaves it in place.
3. On the same Access STA, the client calls `ReadErrorRecord()`; the add-in walks
   `rtcErrObj() -> _g_EbThread -> +0x17C BTSRC_ERR` (the same reader as `VbaTestHost`) and returns
   `{ present, errorNumber, line, column, module, project }`.
4. `VbaCompiler` reports `ok=false`, `module`, `line`, `errorNumber` and a message. `IsCompiled`
   remains the success gate; the add-in is only consulted when the compile fails.

Example from a deliberate `Dim y As MissingType`:

```json
{"present":true,"errorNumber":32807,"line":5,"module":"BadMod"}
```

The record holds the compiler's logical line (a line continuation can shift it by one). It does not
always contain the message text. VBE error numbers are Win32 **RT_STRING IDs** in `VBE7INTL.dll`:
`GetErrMsg` → `_BstrWOfParmStrid` → `_XszOfParmStrid` → `LoadStringA(VBE7INTL, errorNumber)`.
AccessMCP embeds the full **156-code** compiler lookup generated in `~/vbe7`; the record's own
message (when present) takes precedence, then the table, then `VBE compile error <number>`. Sync the
embedded copy with `pwsh ./scripts/sync-vbe-errors.ps1 -Check`.

### Compile-error table

The canonical table lives outside this repository, in the VBE7 research repository:

```text
~/vbe7/
  VBE7_Compile_Error_Codes.md/json   # full compiler table with categories and provenance
  VbeCompileErrors.json              # compact code -> message map consumed by AccessMCP
  corpus/compile-corpus.json         # bad-module corpus
  corpus/compile-results.json        # observed codes from VbaTestHost (46 so far)
  scripts/dump-vbe7intl-strings.ps1  # RT_STRING dump
  scripts/collect-compile-errors.ps1 # empirical oracle
  scripts/build-compile-error-table.ps1
  scripts/make-compact-table.ps1
```

The compiler codes are the non-empty strings in `0x9C63`–`0x9D00`, plus `0x8027`, `0x8028`,
`0x802C`–`0x802E` and `0x88C3`–`0x88D2`. `0x802C` and the hidden-module message are text-substitute
errors: the compiler stores the actual name in the error state, so the record message wins over the
table template.

### Registration

The add-in is registered entirely under `HKCU`, following the working VBAIDETabs pattern:

```
HKCU\Software\Microsoft\VBA\VBE\6.0\Addins\AccessMCP.VbeProbe   (Addins64 for x64 Access)
    FriendlyName, Description, LoadBehavior=0, CommandLineSafe=1
HKCU\Software\Classes\CLSID\{EFAA3F97-4E2F-410E-8545-1F4464E9E060}\InprocServer32
    (default) = installed DLL, ThreadingModel = Apartment
HKCU\Software\Classes\AccessMCP.VbeProbe\CLSID = {…}
```

`Software\Classes` is registry-view redirected, so the script writes the COM keys to **both** the
32-bit and 64-bit views; Access x86 reads the 32-bit view. `LoadBehavior=0` keeps the add-in unloaded
in normal Access sessions; the harness connects it on demand.

### Fallback and scope

- Add-in missing, not registered, or x64 Access without an x64 build: `VbaCompiler` falls back to
  `IsCompiled` plus the VBE compile dialog text. Module/line are then best-effort and may be null.
- The reader is x86-only (the `VBE7` offsets are for the 32-bit build). The current Access
  installation is x86.
- Uninstall with `pwsh ./scripts/register-vbe-addin.ps1 -Unregister`.
- The add-in is an in-process, trusted component; it runs with Access's privileges. It is only
  connected for the duration of a compile and is disconnected afterwards.

---

## 7. Implementation notes and limitations

- **`SysCmd(603)` must run on the same Access instance that compiled and closed the source.**
  A fresh instance silently returns 0 and creates no file. The call must be early-bound (a late-bound
  call with integer 603 also no-ops). The harness releases its VBE project/component references
  **before** `CloseCurrentDatabase`, because holding them across the close can crash Access inside
  `SysCmd`.
- The Office PIAs are not self-contained. The NuGet PIA is referenced directly; `office`,
  `Microsoft.Office.Interop.Access.Dao` and `Microsoft.Vbe.Interop` are referenced from the GAC and
  copied next to the host. `MdeBuilder` registers an `AssemblyLoadContext` resolver for them because
  they are not in the app's `.deps.json`.
- **Compile diagnostics.** `Application.IsCompiled` is the reliable success gate. With the VBE
  add-in installed, `DoCmd.RunCommand(126)` leaves the engine record in place and the add-in reads
  the exact module/line/error number. Without the add-in, module/line fall back to the caret and may
  be null. A compile error normally returns in milliseconds without a modal dialog.
- **`Application.Run` takes unqualified names.** `Run("Module.Proc")` fails with 2517. Tests are
  invoked with direct `Call` statements rather than `Application.Run`, which otherwise lets VBA show
  a modal run-time error dialog instead of returning the error.
- A VBE "Save As" dialog (module name edit id 2020) can appear on quit. The Win32 watcher cancels it;
  cleanup is time-bounded and kills the Access process if it does not finish.
- `ByRef` arguments do not round-trip through COM. Results that are COM objects are reported as
  `[Object: TypeName]`; arrays as `[Array n]`; the recommended contract for a test entry point is a
  JSON string (`Public Function RunAll() As String`).
- Forms: `Application.CreateForm` creates `Form1` without a module; save/close/rename/open in design
  view and set `HasModule=True` to create `Form_<name>`. `Application.Run("Form_X.Proc")` fails with
  2517, so use a standard dispatcher. Direct calls to a form public function work in MDB and MDE in
  current VBAStack: `epiModule+0xC` is a FIXUP, and VBAStack resolves it by asking for the module's
  `ITypeInfo` (`epiModule+2`) and calling `GetTypeAttr`. The `Form_Open(Cancel As Integer)` event is
  still a useful independent test; `Forms(name).Name` priming is only needed with older VBAStack
  builds that do not have the FIXUP fix.
- MDEs are build- and bitness-specific. Refuse mismatched Access/database combinations. An x86-built
  MDE is rejected by x64 Access; build natively on x64. The direct-form path is verified on Access 2013
  x64 (15.0.5349, VBE7 7.01.1127) with an x64 MDB and an x64 `SysCmd(603)`-built MDE. Class
  `Function`/`Sub` frames are also verified there.
- `access_vba_run_tests` is MDB-only. Use `access_vba_test` with `mode="mde"` so the runner is baked
  in before the build.
- `access_vba_ensure_module` with `compile=false` writes through the VBE but may not persist the
  module to disk before Access quits; keep the default `compile=true` when the write must survive.

---

## 8. Testing AccessMCP itself

- Unit tests (no Access): path policy, options, procedure/directive parsing, header normalisation,
  runner generation, NDJSON parsing, MDE cache keys and storage, scratch workspaces, execution
  gating, doctor without COM. They live in `AccessMCP.Tests` (`DigitalJeeves.MsAccessMCP.Tests`).
- Integration (manual/opt-in): create a scratch `.accdb`, ensure modules, compile, run, run tests
  (pass and deliberate fail), build an MDE, run the MDE, verify a cache hit and a timeout kill.
- VBE add-in acceptance: `pwsh ./scripts/test-vbe-addin.ps1` asserts the exact line/module for a
  deliberate compile error through `Application.VBE.AddIns`.
  The plan's `ACCESS_MCP_VBA_TESTS=1` / `ACCESS_MCP_TEST_DB` gate is not wired into CI yet.

```powershell
dotnet test AccessMCP.Tests/DigitalJeeves.MsAccessMCP.Tests.csproj -p:Platform=x64
pwsh ./scripts/sync-vba-testkit.ps1 -Check
```

---

## 9. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `VBE` throws "programmatic access … is not trusted" | Enable Trust Center → Trust access to the VBA project object model |
| `access_vba_run` says "VBA execution is disabled" | Set `EnableVbaTestTools=true` |
| "database is open (its .ldb/.laccdb lock file exists)" | Close the database (or let the previous Access instance finish quitting); the copy waits up to 10 s |
| `SysCmd(603)` returns 0 / no file | The call ran on a different instance, or late-bound. Use the harness's `syscmd` strategy on the MDB that was compiled in the same session |
| "Office PIAs must be deployed next to the MCP host" | Build on a machine with Access installed; `office.dll`, `Microsoft.Office.Interop.Access.Dao.dll` and `Microsoft.Vbe.Interop.dll` must be in the output |
| Tests hang | A test opened a modal dialog or loops. The watchdog kills Access after `timeoutSeconds`; use `keepArtifacts` to inspect `result.ndjson` |
| `access_vba_compile` has no module/line | The VBE add-in is not registered: run `scripts/register-vbe-addin.ps1`; diagnostics fall back to the caret/dialog |`n| A form MDE returns no parameters | Legacy VBAStack only called `ITypeLib::GetTypeInfo(0)`, so it resolved the standard module, never the form. Current VBAStack uses `epiModule+2` and `GetTypeAttr` to resolve the module's `epiModule+0xC` FIXUP; direct form calls then work. `Forms(name).Name` is a fallback for older builds. |







