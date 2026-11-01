<#
.SYNOPSIS
    Compiles SqlPlanViewer.ps1 into SqlPlanViewer.exe (with icon, no console window).

.DESCRIPTION
    Uses the open-source ps2exe module (https://github.com/MScholtes/PS2EXE),
    installed for the current user from the PowerShell Gallery if missing.
    Run it in Windows PowerShell 5.1 (or just double-click Build-SqlPlanViewerExe.cmd),
    from the folder that holds SqlPlanViewer.ps1 and SqlPlanViewer.ico.

    The window stays open at the end and everything is logged to build.log,
    so any error can be read.

    The exe is NOT standalone native code: it embeds the script (readable by
    anyone with the exe) and runs it on the Windows PowerShell 5.1 that ships
    with Windows 10/11.

.PARAMETER Version
    Version number stamped into the exe.

.PARAMETER CertThumbprint
    Optional. Thumbprint of a code-signing certificate (CurrentUser\My or
    LocalMachine\My) to sign the exe with.

.PARAMETER NoPause
    Do not wait for Enter at the end.
#>
[CmdletBinding()]
param(
    [string]$Source,
    [string]$Icon,
    [string]$Output,
    [string]$Version = '1.0.0.0',
    [string]$CertThumbprint,
    [switch]$NoPause
)

$ErrorActionPreference = 'Stop'

# $PSScriptRoot can be empty in parameter defaults on Windows PowerShell 5.1, so resolve paths here.
$here = $PSScriptRoot
if (-not $here -and $MyInvocation.MyCommand.Path) { $here = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $here) { $here = (Get-Location).Path }
if (-not $Source) { $Source = Join-Path $here 'SqlPlanViewer.ps1' }
if (-not $Icon)   { $Icon   = Join-Path $here 'SqlPlanViewer.ico' }
if (-not $Output) { $Output = Join-Path $here 'SqlPlanViewer.exe' }
$log = Join-Path $here 'build.log'
$ok = $false
try { Start-Transcript -Path $log -Force | Out-Null } catch { }

try {
    Write-Host "PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
    if ($PSVersionTable.PSEdition -eq 'Core') {
        Write-Warning 'You are in PowerShell 7. Building from Windows PowerShell 5.1 (powershell.exe) is more reliable.'
    }

    foreach ($f in $Source, $Icon) {
        if (-not (Test-Path -LiteralPath $f)) { throw "Required file not found: $f" }
    }

    # older Windows PowerShell needs TLS 1.2 to talk to the PowerShell Gallery
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    if (-not (Get-Module -ListAvailable -Name ps2exe)) {
        Write-Host 'Installing the ps2exe module (current user) from the PowerShell Gallery...'
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [version]'2.8.5.201' })) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
        }
        Install-Module -Name ps2exe -Scope CurrentUser -Force
    }
    Import-Module ps2exe

    if (Test-Path -LiteralPath $Output) { Remove-Item -LiteralPath $Output -Force }

    Write-Host "Compiling $Source ..."
    Invoke-ps2exe -inputFile $Source -outputFile $Output -iconFile $Icon `
        -noConsole -STA `
        -title 'SQL Plan Viewer' -description 'SQL Server execution plan viewer' `
        -product 'SQL Plan Viewer' -company 'SQL Plan Viewer' -copyright 'SQL Plan Viewer' `
        -version $Version

    if (-not (Test-Path -LiteralPath $Output)) { throw 'ps2exe finished but no exe was produced (see messages above).' }

    if ($CertThumbprint) {
        $cert = Get-ChildItem -Path Cert:\CurrentUser\My, Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $CertThumbprint } | Select-Object -First 1
        if (-not $cert) { throw "Certificate with thumbprint $CertThumbprint not found." }
        $sig = Set-AuthenticodeSignature -FilePath $Output -Certificate $cert -TimestampServer 'http://timestamp.digicert.com'
        if ($sig.Status -ne 'Valid') { Write-Warning "Signing finished with status: $($sig.Status)" }
    }

    Write-Host ('SUCCESS: built {0} ({1:N0} KB)' -f $Output, ((Get-Item -LiteralPath $Output).Length / 1KB)) -ForegroundColor Green
    $ok = $true
}
catch {
    Write-Host ''
    Write-Host 'BUILD FAILED' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    if ($_.InvocationInfo) { Write-Host $_.InvocationInfo.PositionMessage -ForegroundColor DarkRed }
    Write-Host "Details were written to: $log"
}
finally {
    try { Stop-Transcript | Out-Null } catch { }
    if (-not $NoPause) { Write-Host ''; [void](Read-Host 'Press Enter to close') }
}
if (-not $ok) { exit 1 }
