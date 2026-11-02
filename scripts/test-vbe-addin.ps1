#requires -Version 7.0
<#
.SYNOPSIS
    Acceptance test for the AccessMCP VBE add-in.

.DESCRIPTION
    Creates a scratch .accdb with a deliberate compile error, loads the add-in via
    Application.VBE.AddIns, runs the normal compile command (126), and asserts that
    ReadErrorRecord returns the VBE's own module and line.

.EXAMPLE
    pwsh ./scripts/test-vbe-addin.ps1
#>
[CmdletBinding()]
param(
    [switch]$SkipRegister,
    [switch]$KeepArtifacts
)

$ErrorActionPreference = 'Stop'

if (-not $SkipRegister) {
    & (Join-Path $PSScriptRoot 'register-vbe-addin.ps1') | Out-Host
}

$ProgId = 'AccessMCP.VbeProbe'
$dir = Join-Path $env:TEMP ('AccessMcpVbeAddin-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dir | Out-Null
$db = Join-Path $dir 'probe.accdb'

$app = $null
$pid2 = $null
$killer = $null
try {
    $app = New-Object -ComObject Access.Application
    $app.Visible = $false
    $app.UserControl = $false
    $app.AutomationSecurity = 1
    $app.NewCurrentDatabase($db)
    $pid2 = (Get-Process MSACCESS | Sort-Object StartTime | Select-Object -Last 1).Id
    $killer = Start-Job -ArgumentList $pid2 -ScriptBlock {
        param($processId)
        Start-Sleep -Seconds 20
        Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
    }

    $project = $app.VBE.VBProjects.Item(1)

    $good = $project.VBComponents.Add(1)
    $good.Name = 'GoodMod'
    $good.CodeModule.InsertLines($good.CodeModule.CountOfLines + 1,
        "Public Function Ok() As Long`r`n    Ok = 1`r`nEnd Function")

    $bad = $project.VBComponents.Add(1)
    $bad.Name = 'BadMod'
    $bad.CodeModule.InsertLines($bad.CodeModule.CountOfLines + 1,
        "Public Sub Broken()`r`n    Dim y As MissingType`r`nEnd Sub")

    $vbe = $app.VBE
    $vbe.AddIns.Update()

    $addin = $null
    try {
        $addin = $vbe.AddIns.Item($ProgId)
    }
    catch {
        Write-Host 'Add-in not found in VBE.AddIns. Available add-ins:'
        for ($i = 1; $i -le $vbe.AddIns.Count; $i++) {
            $item = $vbe.AddIns.Item($i)
            Write-Host ("  {0}  {1}" -f $item.ProgId, $item.Description)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($item) | Out-Null
        }
        throw "Add-in '$ProgId' is not registered. Run scripts/register-vbe-addin.ps1."
    }

    Write-Host "Add-in: ProgId=$($addin.ProgId) Guid=$($addin.Guid)"
    $addin.Connect = $true
    Start-Sleep -Milliseconds 300
    if (-not $addin.Connect) { throw 'The add-in did not connect.' }

    $probe = $addin.Object
    if ($null -eq $probe) { throw 'AddIn.Object is null.' }

    # The normal compile command. It does not show a dialog via automation; it sets
    # the VBE's BTSRC_ERR record, which the add-in reads in-process.
    $app.DoCmd.RunCommand(126)
    Write-Host "IsCompiled=$($app.IsCompiled)"

    $json = $probe.ReadErrorRecord()
    Write-Host "ReadErrorRecord=$json"

    $parsed = $json | ConvertFrom-Json
    if ($parsed.present -ne $true) { throw 'The VBE error record is not present.' }
    if ($parsed.line -ne 5) { throw "Expected line=5, got $($parsed.line)." }
    if ($parsed.module -ne 'BadMod') { throw "Expected module=BadMod, got $($parsed.module)." }
    if (-not $parsed.errorNumber) { throw 'Expected an error number.' }

    Write-Host 'PASS: exact module and line from the VBE error record.'

    $addin.Connect = $false
    Start-Sleep -Milliseconds 200
    Write-Host "Disconnected: $(-not $addin.Connect)"
}
catch {
    Write-Host "FAILED: $($_.Exception.Message)"
    throw
}
finally {
    if ($null -ne $killer) {
        Stop-Job $killer -ErrorAction SilentlyContinue
        Remove-Job $killer -Force -ErrorAction SilentlyContinue
    }

    if ($null -ne $app) {
        try { $app.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
    }

    if ($pid2) {
        Get-Process -Id $pid2 -ErrorAction SilentlyContinue | ForEach-Object {
            try { $_.Kill() } catch { }
        }
    }

    if (-not $KeepArtifacts) {
        Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue
    }
    else {
        Write-Host "Artifacts: $dir"
    }
}
