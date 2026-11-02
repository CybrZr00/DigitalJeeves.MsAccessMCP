#requires -Version 7.0
<#
.SYNOPSIS
    Registers (or unregisters) the per-user AccessMCP VBE add-in.

.DESCRIPTION
    The add-in is a native x86 COM DLL. Registration is entirely under HKCU, so no
    administrator rights are required. The DLL is copied to a stable per-user folder
    and the registry points at that copy.

    The VBE add-in key used by VBAIDETabs (and confirmed working with Access) is:
        HKCU\Software\Microsoft\VBA\VBE\6.0\Addins\<ProgID>          (x86)
        HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64\<ProgID>        (x64)

    HKCR/HKCU Software\Classes is registry-view redirected. Access x86 reads the
    32-bit view, so the COM keys are written to both views. LoadBehavior=0 keeps the
    add-in unloaded in normal Access sessions; AccessMCP connects it on demand through
    Application.VBE.AddIns.

.PARAMETER DllPath
    The built add-in DLL. Defaults to the Release x86 build in this repository.

.PARAMETER Unregister
    Remove the registry keys and the installed copy.

.EXAMPLE
    pwsh ./scripts/register-vbe-addin.ps1
    pwsh ./scripts/register-vbe-addin.ps1 -Unregister
#>
[CmdletBinding()]
param(
    [string]$DllPath,
    [switch]$Unregister
)

$ErrorActionPreference = 'Stop'

$ProgId = 'AccessMCP.VbeProbe'
$Clsid = '{EFAA3F97-4E2F-410E-8545-1F4464E9E060}'
$FriendlyName = 'AccessMCP VBE Probe'
$Description = 'In-process compile diagnostics for AccessMCP.'
$InstallFolder = Join-Path $env:LOCALAPPDATA 'AccessMCP\VbeAddin'
$InstalledDll = Join-Path $InstallFolder 'AccessMCP.VbeProbe.dll'
$AddInSubKey = "Software\Microsoft\VBA\VBE\6.0\Addins\$ProgId"
$ProgIdSubKey = "Software\Classes\$ProgId"
$ClsidSubKey = "Software\Classes\CLSID\$Clsid"
$Views = @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)

function Get-PEMachine([string]$Path) {
    if (-not (Test-Path $Path)) {
        throw "File not found: $Path"
    }

    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $reader = [System.IO.BinaryReader]::new($stream)
        $stream.Position = 0x3C
        $offset = $reader.ReadInt32()
        $stream.Position = $offset
        if ($reader.ReadUInt32() -ne 0x00004550) {
            throw "Not a PE file: $Path"
        }
        return $reader.ReadUInt16()
    }
    finally {
        $stream.Dispose()
    }
}

function New-RegKey([Microsoft.Win32.RegistryView]$View, [string]$SubKey) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $View)
    try {
        $key = $base.CreateSubKey($SubKey)
        $key.Dispose()
    }
    finally {
        $base.Dispose()
    }
}

function Set-RegValue(
    [Microsoft.Win32.RegistryView]$View,
    [string]$SubKey,
    [string]$Name,
    [object]$Value,
    [Microsoft.Win32.RegistryValueKind]$Kind
) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $View)
    try {
        $key = $base.CreateSubKey($SubKey)
        try {
            $key.SetValue($Name, $Value, $Kind)
        }
        finally {
            $key.Dispose()
        }
    }
    finally {
        $base.Dispose()
    }
}

function Remove-RegKey([Microsoft.Win32.RegistryView]$View, [string]$SubKey) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $View)
    try {
        $base.DeleteSubKeyTree($SubKey, $false)
    }
    catch {
        # Already absent in this view.
    }
    finally {
        $base.Dispose()
    }
}

if ($Unregister) {
    foreach ($view in $Views) {
        Remove-RegKey $view $AddInSubKey
        Remove-RegKey $view $ProgIdSubKey
        Remove-RegKey $view $ClsidSubKey
    }

    if (Test-Path $InstallFolder) {
        Remove-Item -Path $InstallFolder -Recurse -Force
    }

    Write-Host "Unregistered $ProgId and removed $InstallFolder."
    return
}

if (-not $DllPath) {
    $DllPath = Join-Path $PSScriptRoot '..\tools\AccessMCP.VbeAddin\bin\x86\Release\AccessMCP.VbeProbe.dll'
}
$DllPath = [System.IO.Path]::GetFullPath($DllPath)

$machine = Get-PEMachine $DllPath
if ($machine -ne 0x014C) {
    throw ("The DLL at $DllPath is not x86 (PE machine 0x{0:X4}). Access here is x86." -f $machine)
}

New-Item -ItemType Directory -Force -Path $InstallFolder | Out-Null
try {
    Copy-Item -Path $DllPath -Destination $InstalledDll -Force -ErrorAction Stop
}
catch [System.IO.IOException] {
    throw "Cannot replace $InstalledDll because it is in use. Close every Access instance (and any MCP server that has loaded the add-in) and retry."
}

foreach ($view in $Views) {
    # COM class, per-user.
    Set-RegValue $view $ProgIdSubKey '' $FriendlyName ([Microsoft.Win32.RegistryValueKind]::String)
    Set-RegValue $view "$ProgIdSubKey\CLSID" '' $Clsid ([Microsoft.Win32.RegistryValueKind]::String)

    Set-RegValue $view $ClsidSubKey '' $FriendlyName ([Microsoft.Win32.RegistryValueKind]::String)
    Set-RegValue $view "$ClsidSubKey\ProgID" '' $ProgId ([Microsoft.Win32.RegistryValueKind]::String)
    Set-RegValue $view "$ClsidSubKey\InprocServer32" '' $InstalledDll ([Microsoft.Win32.RegistryValueKind]::String)
    Set-RegValue $view "$ClsidSubKey\InprocServer32" 'ThreadingModel' 'Apartment' ([Microsoft.Win32.RegistryValueKind]::String)
    New-RegKey $view "$ClsidSubKey\Programmable"

    # VBE add-in registration. LoadBehavior=0: do not load until AccessMCP connects it.
    Set-RegValue $view $AddInSubKey 'FriendlyName' $FriendlyName ([Microsoft.Win32.RegistryValueKind]::String)
    Set-RegValue $view $AddInSubKey 'Description' $Description ([Microsoft.Win32.RegistryValueKind]::String)
    Set-RegValue $view $AddInSubKey 'LoadBehavior' 0 ([Microsoft.Win32.RegistryValueKind]::DWord)
    Set-RegValue $view $AddInSubKey 'CommandLineSafe' 1 ([Microsoft.Win32.RegistryValueKind]::DWord)
}

Write-Host "Registered $ProgId ($Clsid) from $InstalledDll"
Write-Host "VBE add-in key: HKCU\$AddInSubKey (32-bit and 64-bit views)"
