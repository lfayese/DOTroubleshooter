<#
.SYNOPSIS
  Launches Delivery Optimization Update Check tool using portable PowerShell 7.4.0 and runs the main troubleshooter script.

.DESCRIPTION
  - Looks for .\pwsh\pwsh.exe (portable PS)
  - Launches .\DOTUpdateCheck.ps1 with passed args
  - Designed for ps2exe conversion into a single .exe launcher
  
  .EXAMPLE
    .\Run-DOTUpdateCheck.ps1 -DiagnosticsZip "C:\Temp\DiagnosticsData.zip"

     DOTUpdateCheck.exe -DiagnosticsZip "C:\Temp\DiagnosticsData.zip"
    
  .EXAMPLE
    .\Run-DOTUpdateCheck.ps1 -Show -OutputPath "C:\Reports" -Verbose

     DOTUpdateCheck.exe -Show -OutputPath "C:\Reports" -Verbose
  
  .EXAMPLE
    .\Run-DOTUpdateCheck.ps1 -DiagnosticsZip "C:\Temp\DiagnosticsData.zip" -Show -OutputPath "C:\Reports" -Verbose

     DOTUpdateCheck.exe -DiagnosticsZip "C:\Temp\DiagnosticsData.zip" -Show -OutputPath "C:\Reports" -Verbose
  
#>
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ArgsPassed
)
function Get-RootPath {
    if ($PSScriptRoot) {
        # Running as script using PSScriptRoot
        return $PSScriptRoot
    } elseif ($MyInvocation.MyCommand.Path) {
        # Fallback to MyInvocation if available
        return Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        # Running as exe: use current directory
        return (Get-Location).Path
    }
}
# Get base directory for execution
$Root = Get-RootPath
$PwshPath   = Join-Path -Path $Root -ChildPath "pwsh\pwsh.exe"
$MainScript = Join-Path -Path $Root -ChildPath "DOTUpdateCheck.ps1"
# Logging debug info
Write-Host "[DEBUG] Current execution path: $Root" -ForegroundColor Yellow
Write-Host "[DEBUG] PwshPath: $PwshPath" -ForegroundColor Yellow
Write-Host "[DEBUG] MainScript: $MainScript" -ForegroundColor Yellow
# Validate existence of executable and main script
if (-not (Test-Path $PwshPath)) {
    Write-Host "[ERROR] Could not find pwsh.exe at: $PwshPath" -ForegroundColor Red
    Start-Sleep -Seconds 5
    exit 1
}
if (-not (Test-Path $MainScript)) {
    Write-Host "[ERROR] Could not find DOTUpdateCheck.ps1 at: $MainScript" -ForegroundColor Red
    Start-Sleep -Seconds 5
    exit 1
}
Write-Host "[INFO] Launching DOTUpdateCheck with PowerShell 7.4.0..." -ForegroundColor Cyan
Write-Host "       Script: $MainScript"
if ($ArgsPassed.Count) {
    Write-Host "       Args:   $($ArgsPassed -join ' ')"
}
# Build the argument list as an array for robustness
$psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $MainScript)
if ($ArgsPassed) {
    $psArgs += $ArgsPassed
}

Start-Process -FilePath $PwshPath -ArgumentList $psArgs -WorkingDirectory $Root -WindowStyle Normal -Wait