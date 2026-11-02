# VBA harness spikes — results

**Date:** 2026-10-18
**Machine:** Access 365 x86, 16.0.20430.20182 (Click-to-Run), .NET 10 x64 host
**Scope:** spikes S1–S8 from `docs/VBA-TEST-HARNESS-PLAN.md`

These results supersede the assumptions in the proposal where they differ. The most important
correction is **S2**: `SysCmd(603)` does *not* work from a fresh Access instance.

---

## S1 — compile failure and `IsCompiled`

Method: create a scratch `.accdb`, add `BadMod` containing `Dim y As MissingType`, call
`DoCmd.RunCommand 126`, read `Application.IsCompiled`, inspect the active code pane and windows.

Result:

- `RunCommand(126)` returns in **5–30 ms**; it does **not** block on a dialog in this case.
- `Application.IsCompiled` is reliably **False**.
- `DoCmd.OpenModule("BadMod")` on the broken project throws `Cannot Compile Project`.
- The active code pane is set to `BadMod`, but `GetSelection` returned **line 7** for an error on
  line 5 (the caret was parked past the module end). **The caret is not a reliable line number.**
- `Application.Run("Ok")` still executes a good procedure while another module has a compile error;
  `IsCompiled` stays False.
- A compile-error dialog was not shown by automation; the Win32 watcher remains as a fallback.

Conclusion: use `IsCompiled` as the gate; report module/line only when the caret is plausible, and
treat them as hints.

---

## S2 — `SysCmd(603)` MDE build (corrected)

Method: create a compiled `.accdb`, close it, then call `SysCmd(603, source, target)` from:
(a) a fresh Access instance, (b) the same instance after `CloseCurrentDatabase`, (c) late-bound
PowerShell with `int`, `short` and `long` 603, (d) early-bound C# with the PIA.

Result:

| Attempt | Result |
|---|---|
| Fresh instance, late-bound (PowerShell) | returns 0, **no file** |
| Fresh instance, early-bound PIA | returns 0, **no file** |
| Same instance, source still open | returns 0, no file |
| **Same instance, after compile + `CloseCurrentDatabase`** | **returns 1, file created** (286,720 bytes for a small ACCDB) |
| `NewCurrentDatabase("x.mdb")` then build to `.mde` | file is actually ACE/ACCDB format; no output. The extension must match the container |

Additional requirements found:

- The call must be **early-bound**. `dynamic`/PowerShell with integer 603 is a silent no-op (it does
  not always raise 7952).
- The source must be **compiled and closed** in that same instance immediately before the call.
- The output path is used verbatim; Access does not append `.accde`.
- Holding a VBE `VBProject`/`VBComponents` reference across `CloseCurrentDatabase` crashed Access
  inside `SysCmd` (`0x800706BE`). Release all VBE references, then `GC.Collect()`, then close, then
  `SysCmd`.
- The Office PIAs must be deployed: `Microsoft.Office.Interop.Access` (NuGet) plus `office.dll`,
  `Microsoft.Office.Interop.Access.Dao.dll`, `Microsoft.Vbe.Interop.dll` (GAC). A
  `AssemblyLoadContext` resolver is needed because the GAC assemblies are not in `.deps.json`.

Conclusion: the harness's `syscmd` strategy uses the **same session** for compile, close and build;
`MdeBuilder` retries once if the connection drops.

---

## S3 — x64 host → x86 Access VBE

Method: PowerShell 7 (x64) creates `Access.Application` (x86) and enumerates `app.VBE.VBProjects` and
1,128 `VBComponents` on a copy of `sample_form.mdb`.

Result: **works**. No x86 helper process is required. The existing `access_vbe_*` tools are already
x64-host-safe.

---

## S5 — module creation and writes

Method: create standard and class modules with `VBComponents.Add`, set `.Name`, write code with
`InsertLines`/`ReplaceLine`, compile, run.

Result:

- `VBComponents.Add(vbext_ct_StdModule)` creates a module with `Option Compare Database`,
  `Option Explicit` and a blank line. **Do not include `Attribute VB_Name` in the body** — it is a
  syntax error there.
- Write the body after the existing Option lines (or replace them deliberately). The verified
  line-level writer works; `AddFromString` is still avoided (known corruption).
- A class module created with `Add(vbext_ct_ClassModule)` compiles and can be instantiated from a
  standard module. `Application.Run("NewClassName")` is not a valid way to instantiate it.
- `Application.Run("Module.Procedure")` fails with 2517; **unqualified `Run("Procedure")` works**.
- `access_vba_ensure_module` strips export headers (`Attribute`, `VERSION`/`BEGIN`/`END`) and
  preserves the module's Option block.

---

## S6 — automation security and startup forms

- Set `Application.AutomationSecurity = 1` before `OpenCurrentDatabase`.
- `Application.Visible = false` hides the Access window, but the **VBE window is still visible** when
  a module is opened (`wndclass_desked_gsk`). This is cosmetic for a scratch run; do not rely on
  `Visible=false` to guarantee no UI.
- Removing/restoring `StartUpForm` through DAO works on an MDB copy. Do it **before** compiling and
  building the MDE; an MDE must not be edited.
- The VBE "Save As" dialog appears on quit after unsaved module writes. Structure (this build):

```
#32770 "Save As"
  id=2021 Static    "Module &Name:"
  id=2020 RichEdit20W
  id=1    Button    OK
  id=2    Button    Cancel
```

The watcher cancels it (accepting would write stray `.bas` files) and cleanup is bounded by a kill
timer.

---

## S7 - form creation and the compiled form-frame result

The proposal's route is confirmed: `CreateForm` → save/close `Form1` → `Rename` → open in design view
→ `HasModule = True` creates `Form_<name>` → write the module. A standard dispatcher module is the
entry point (`Application.Run("Form_X.Proc")` gives 2517).

MDB: the form frame and parameters are correct.

Compiled DBs: `GetCallstack` works. The old failure was `epiModule+0xC` being an unresolved 8-byte
`FIXUP` (value `0x50000`, flags `0x80000004`), not a private-object pointer. `CheckAddressSafe`
accepted `0x50000`, so the old code skipped the load and read `+0x18=0x40`. The previous
`LoadAndRelease_TypeAttrs` always requested `ITypeLib::GetTypeInfo(0)`, i.e. the standard module, so
its side effect never resolved the form's FIXUP.

Fix: use the module object index at `epiModule+2` for `GetTypeInfo`, then call
`ITypeInfo::GetTypeAttr`. `CEcTypeInfo::GetTypeAttr` reaches `memSection::Pv`, which loads the section
and rewrites the FIXUP to a direct pointer.

Verified with no form open and no `Forms(...).Name` priming:

| Target | Direct `RunFormParams` |
|---|---|
| MDB | `ByRef x As Long : 42`, `ByRef y As String : hello` |
| UI-built ACCDE | same |
| `SysCmd(603)`-built MDE | same |
| x64 MDB (Access 2013 x64) | same |
| x64 `SysCmd(603)`-built MDE | same |
| x64 class `CThing.DoFunc` / `DoSub` | `CThing|DoFunc|...` / `CThing|DoSub|...` in MDB and x64 MDE |

The `Form_Open(Cancel As Integer)` event still returns `ByRef Cancel As Integer : 0` in MDB and MDE.
Both paths pass twice in one process and across two processes.

The x64 test used Access 2013 x64 (15.0.5349, VBE7 7.01.1127). x86-built MDE/ACCDE files were rejected
by x64 Access, so the MDE was rebuilt natively with `SysCmd(603)`. The class fixture (`CThing.DoFunc`
and `CThing.DoSub`, invoked through `C7Dispatch`) also passed in the x64 MDB and the native x64 MDE,
confirming the class-frame hidden-`Me` / hidden-return handling on x64. Automation note: the
`Access.Application` class factory failed on that desktop (`0x80080005`); launching `MSACCESS.EXE "db"`
and attaching with `Marshal.GetActiveObject("Access.Application")` worked, and late-bound
`SysCmd(603, source, target)` succeeded after open → compile → `IsCompiled` → `CloseCurrentDatabase`.

M6 is a **VBAStack fix**, not a test correction.

---

## S8 — cache key and round-trip

Method: run `access_vba_test` in `mode=mde` twice with the same source and modules.

Result:

- First call (cache miss) builds and writes `<cache>\entries\<key>\base.accde` + `manifest.json`.
- Second call with `useCache=true` returns `cacheHit: true` and does not rebuild.
- The key includes the **original source path**, size and last-write ticks, Access version/bitness,
  every module hash, the kit hash, the decompile flag and the test list. Using the per-run workspace
  path instead of the source made every key unique; that was fixed.
- `rebuild=true` bypasses the read but still refreshes the entry.
- No database marker is needed: a cache hit is validated by the scenario's own run; a corrupt entry
  fails the run and can be cleared with `access_vba_cache_clear`.

---

## Additional observations

- Access automation is effectively **single-instance**: `Activator.CreateInstance` may attach to an
  existing `MSACCESS.EXE`. Leftover modal dialogs from a previous run can wedge the next one. The
  harness serialises sessions, watches dialogs and kills by PID on timeout.
- `MSACCESS.EXE` may keep the `.ldb` file for a moment after `Quit`. Waiting for an exclusive read
  (not merely checking for `.ldb`) avoids false "database is open" failures.
- `Application.Run` of a procedure that raises an error from a test can produce a **modal VBA
  run-time error dialog** even with `On Error Resume Next` in the caller. Generated tests therefore
  use direct `Call TestName` statements. This also removed a 180 s timeout in the deliberate-failure
  test.
- `RunCommand(280)` (`acCmdSaveAllModules`) is not available in a hidden automation session; compile
  (126) is the reliable way to persist module writes.
- The Office PIAs are a build-time and run-time dependency. `access_vba_doctor` reports
  `piaAvailable` by checking the deployed files and the loadable types.

---

## S9 — VBE add-in for exact compile diagnostics (2026-10-18)

Result: **works**. The add-in is loaded through `Application.VBE.AddIns` and returns the VBE's
`BTSRC_ERR` record after the normal compile command.

Findings:

- **Registration.** The keys used by VBAIDETabs work:
  `HKCU\Software\Microsoft\VBA\VBE\6.0\Addins\<ProgID>` (x86) / `Addins64` (x64), plus the
  per-user COM class under `HKCU\Software\Classes\CLSID\{GUID}\InprocServer32`.
  `HKCU\Software\Classes` is registry-view redirected: the 32-bit view had to be written separately,
  otherwise Access x86 reported `REGDB_E_CLASSNOTREG`.
- **Loading.** `VBE.AddIns.Update()` then `VBE.AddIns.Item("AccessMCP.VbeProbe")` and
  `AddIn.Connect = true` loads the DLL. `LoadBehavior=0` keeps it unloaded in normal sessions.
- **`AddIn.Object`.** For a native COM add-in the VBE does not populate `Object` automatically; the
  add-in sets the `AddIn.Object` property in `OnConnection`. Without that, `AddIn.Object` is null.
- **DISPIDs.** `IDTExtensibility2` uses dispatch IDs 1–5. Custom methods must not use 1/2; they are
  mapped to 10+ (`ReadErrorRecord` is 12).
- **The compile record survives `DoCmd.RunCommand(126)`.** After the normal compile command, the
  add-in's `ReadErrorRecord()` returned:
  ```json
  {"present":true,"errorNumber":32807,"line":5,"module":"BadMod"}
  ```
  for a deliberate `Dim y As MissingType` on line 5. `Application.IsCompiled` was false. This is the
  same `BTSRC_ERR` record `VbaTestHost` reads, but obtained in-process from Access.
- **The public wrappers do not expose the compiler.** `VBProject` QI for `IVbaProject` and the VBE
  QI for `IVba` both return `E_NOINTERFACE`, even in-process. The project's `ITypeLib` (reached via
  `References` + `0x14` on x86) QIs for `IVBEProject`, but its `CompileProject()` sees only the saved
  type-library snapshot: it returned `S_OK` for an unsaved broken module. Not used.
- **The VBE CommandBars Compile command (578)** shows the engine description but clears the record
  and disconnects the add-in. The normal `126` command is the right trigger.
- **Message text.** The record's message pointer at `+0x08` did not yield a printable string on this
  build; the add-in returns `errorNumber` (32807). AccessMCP maps 32807 to
  "User-defined type not defined" and otherwise reports the number.
- **Bitness.** The reader uses the x86 VBE7 offsets, so the add-in is x86 only (matching Access
  x86). AccessMCP remains x64 and calls the add-in cross-process.

Acceptance test:

```powershell
pwsh ./scripts/test-vbe-addin.ps1
# ReadErrorRecord={"present":true,"errorNumber":32807,"line":5,"module":"BadMod"}
# PASS: exact module and line from the VBE error record.
```

---

## S10 — VBE compile-error message table (2026-10-19)

Result: **complete compiler mapping found**. The `BTSRC_ERR.number` is not an arbitrary code; it is the
Win32 **RT_STRING resource ID** in `VBE7INTL.dll`.

Lookup chain (REA on `VBE7.dll`, PDB symbols available):

```text
GetCompileErrorInfo
  -> GetErrMsg(number)
     -> _BstrWOfParmStrid(number)
        -> _XszOfParmStrid(number)
           -> LoadStringA(_g_hinst /* VBE7INTL.dll */, number)
```

- `VBE7.dll` = `f3ab53c1e4a3c1e8da3ea98769d152f39a382089f6c276e42758371bb8d27d6e`
- `VBE7INTL.dll` = `af8d8033a4f0d699cc7a35ed3f45136f67d26877e67dbb2032fbdfbb6bb85500`
- `VBE7.pdb` = `aae5eb9b49d01096facd0e0532d39012a71b2919c1b1bb65b720a398455bbd79`
- `VbaTestHost.exe` x86 Debug = `13daa587f85ee212d1d70b833b6052bdd380d115a119bfe849015ef7273596d6`

`GetErrMsg` uses `_BstrWOfParmStrid`, which replaces `|` placeholders with arguments; for compile
errors the argument list is empty. `IsTextSubstErr` (`VBE7.dll:0x1022ff14`, 15 DWORDs) lists the codes
that use the compiler's stored text instead of the string table. The compile-relevant entries are
`0x802C` (ambiguous name) and `0x9D09` (compile error in hidden module); their record text wins over
the table template.

### Enumeration

`LoadStringW` on `VBE7INTL.dll` with `LOAD_LIBRARY_AS_DATAFILE` gives the table directly. Ranges:

| Range | Meaning | Non-empty |
|---|---|---:|
| `0x8000-0x8FFF` | project/file/automation errors plus literal constants | 37 |
| `0x9000-0x9FFF` | compiler `0x9C63-0x9D00`, then runtime/UI | 297 |
| `0xDF00-0xDFFF` | UI/format strings | 56 |

Compiler codes are `0x9C63-0x9D00`, plus `0x8027`, `0x8028`, `0x802C-0x802E` and
`0x88C3-0x88D2`: **156 codes**. `0x9CD7-0x9CEE` are the individual `Expected: ...` fragments, kept
separate from standalone compiler codes.

### Empirical validation

`VbaTestHost` x86 was used as the oracle: a 57-snippet corpus produced **46 distinct compiler codes**,
and every description matched the `LoadStringW` text. Examples:

| Code | Message | Corpus sample |
|---|---|---|
| `0x8027` | User-defined type not defined | `Dim y As MissingType` |
| `0x9C70` | Syntax error | `If Then` |
| `0x9C8F` | Expected End Sub | missing `End Sub` |
| `0x9C9F` | Duplicate declaration in current scope | duplicate `Dim x` |
| `0x9CAF` | Variable not defined | `Option Explicit`, `y = 1` |
| `0x9C92` | Block If without End If | `If True Then` without `End If` |
| `0x9CB2` | Constant expression required | `Dim a(n)` with variable `n` |
| `0x9CFC` | User-defined type without members not allowed | empty `Type T` |

Artifacts live in `~/vbe7`:

```text
VBE7_Compile_Error_Codes.json/.md   full table, 156 codes
VbeCompileErrors.json               compact 156-entry map
VBE7INTL_strings.json               all 390 non-empty IDs with classification
corpus/compile-corpus.json          57 snippets
corpus/compile-results.json/.md     observed codes and descriptions
scripts/*.ps1                       dump, collect, build, compact
```

AccessMCP embeds `VbeCompileErrors.json` through `VbeCompileErrorTable`; the record's own message
takes precedence, then the table, then `VBE compile error <number>`. `scripts/sync-vbe-errors.ps1
-Check` guards the embedded copy.
