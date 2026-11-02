#requires -Version 7.0
<#
.SYNOPSIS
    Synchronises the canonical VbaTestKit VBA module into AccessMCP's embedded resource.

.DESCRIPTION
    The assertion kit has one source of truth shared with VbaTestHost. Until the kit is extracted
    into its own repository/submodule, this script reads the raw string template out of
    VbaTestHost/main.cpp (or a canonical VbaTestKit.bas when one exists) and writes it to
    AccessMCP.Core/Resources/VbaTestKit.bas. A hash file is written next to it so a CI check can
    detect drift.

.PARAMETER Source
    Path to main.cpp (or a VbaTestKit.bas file). Defaults to the sibling VBASA checkout.

.PARAMETER Destination
    Destination .bas path. Defaults to AccessMCP.Core/Resources/VbaTestKit.bas.

.PARAMETER Check
    Do not write anything; exit 1 when the destination differs from the source.

.EXAMPLE
    pwsh ./scripts/sync-vba-testkit.ps1 -Check
#>
[CmdletBinding()]
param(
    [string]$Source,
    [string]$Destination,
    [switch]$Check
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $Destination) {
    $Destination = Join-Path $repoRoot 'AccessMCP.Core/Resources/VbaTestKit.bas'
}

if (-not $Source) {
    $candidates = @(
        (Join-Path (Split-Path -Parent $repoRoot) 'VBASA/VbaTestKit/VbaTestKit.bas'),
        (Join-Path (Split-Path -Parent $repoRoot) 'VBASA/VbaTestHost/main.cpp'),
        'D:\Workspace\source\repos\VBASA\VbaTestKit\VbaTestKit.bas',
        'D:\Workspace\source\repos\VBASA\VbaTestHost\main.cpp'
    )

    $Source = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $Source) {
        throw 'Could not find the canonical kit. Pass -Source explicitly.'
    }
}

$text = Get-Content -Path $Source -Raw

if ($Source -like '*.bas') {
    $kit = $text
}
else {
    $match = [regex]::Match($text, 'kSupportTemplate\s*=\s*LR"VBASRC\((?<body>[\s\S]*?)\)VBASRC"')
    if (-not $match.Success) {
        throw "No kSupportTemplate raw string found in $Source"
    }

    $kit = $match.Groups['body'].Value
}

# Normalise to LF for hashing, but keep the file's existing line endings on disk.
$normalised = $kit.Replace("`r`n", "`n")
$hash = [System.Security.Cryptography.SHA256]::HashData([System.Text.Encoding]::UTF8.GetBytes($normalised.TrimEnd("`n")))
$hashHex = [Convert]::ToHexString($hash).ToLowerInvariant()

if ($Check) {
    if (-not (Test-Path $Destination)) {
        Write-Error "Destination $Destination does not exist."
        exit 1
    }

    $current = (Get-Content -Path $Destination -Raw).Replace("`r`n", "`n")
    $currentHash = [System.Security.Cryptography.SHA256]::HashData([System.Text.Encoding]::UTF8.GetBytes($current.TrimEnd("`n")))
    $currentHex = [Convert]::ToHexString($currentHash).ToLowerInvariant()

    if ($currentHex -ne $hashHex) {
        Write-Error "VbaTestKit.bas is out of date (source $hashHex, destination $currentHex). Run scripts/sync-vba-testkit.ps1."
        exit 1
    }

    Write-Host "VbaTestKit.bas is up to date ($hashHex)."
    exit 0
}

$directory = Split-Path -Parent $Destination
New-Item -ItemType Directory -Force -Path $directory | Out-Null
[System.IO.File]::WriteAllText($Destination, $kit, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText("$Destination.sha256", "$hashHex`n", [System.Text.UTF8Encoding]::new($false))

Write-Host "Wrote $Destination ($($kit.Length) chars, sha256 $hashHex)."
