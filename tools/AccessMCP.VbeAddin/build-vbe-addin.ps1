#requires -Version 7.0
<#
.SYNOPSIS
    Builds the x86 AccessMCP VBE add-in.

.DESCRIPTION
    The add-in is a native x86 COM DLL. It must match the bitness of Access, not the
    AccessMCP host, so it is built separately from the .NET solution with the x86 MSVC
    toolchain. It links the static CRT so the add-in has no VC runtime dependency.

.EXAMPLE
    pwsh ./tools/AccessMCP.VbeAddin/build-vbe-addin.ps1
#>
[CmdletBinding()]
param(
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$out = Join-Path $root "bin\x86\$Configuration"
New-Item -ItemType Directory -Force -Path $out | Out-Null

function Find-VisualStudio {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (Test-Path $vswhere) {
        $path = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
        if ($LASTEXITCODE -eq 0 -and $path) {
            return $path.Trim()
        }
    }

    foreach ($candidate in @('D:\VS2026Enterprise', 'C:\Program Files\Microsoft Visual Studio\2022\Enterprise', 'C:\Program Files\Microsoft Visual Studio\2022\Professional', 'C:\Program Files\Microsoft Visual Studio\2022\Community')) {
        if (Test-Path (Join-Path $candidate 'VC\Auxiliary\Build\vcvars32.bat')) {
            return $candidate
        }
    }

    throw 'Visual Studio with the C++ x86 toolchain was not found.'
}

$vs = Find-VisualStudio
$vcvars = Join-Path $vs 'VC\Auxiliary\Build\vcvars32.bat'
if (-not (Test-Path $vcvars)) {
    throw "vcvars32.bat not found under $vs."
}

$optimization = if ($Configuration -eq 'Debug') { '/Od /MTd /Zi' } else { '/O2 /MT' }

$arguments = @(
    'call', "`"$vcvars`"", '>nul', '&&',
    'cl', '/nologo', '/EHsc', '/std:c++17', '/DUNICODE', '/D_UNICODE', '/LD',
    $optimization,
    "/Fo:`"$out\\`"",
    "/Fe:`"$out\AccessMCP.VbeProbe.dll`"",
    "`"$root\VbeProbe.cpp`"",
    "`"$root\VbeErrorRecord.cpp`"",
    '/link', "/DEF:`"$root\VbeProbe.def`"",
    'ole32.lib', 'oleaut32.lib', 'user32.lib', 'advapi32.lib'
)

Write-Host "Building AccessMCP.VbeProbe.dll (x86, $Configuration)..."
$commandLine = ($arguments -join ' ')
& cmd.exe /d /s /c $commandLine
if ($LASTEXITCODE -ne 0) {
    throw "cl.exe failed with exit code $LASTEXITCODE."
}

$dll = Join-Path $out 'AccessMCP.VbeProbe.dll'
if (-not (Test-Path $dll)) {
    throw "Build succeeded but $dll was not produced."
}

$file = Get-Item $dll
Write-Host ("Built {0} ({1:N0} bytes)" -f $file.FullName, $file.Length)
