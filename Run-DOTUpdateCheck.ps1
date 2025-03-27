<#
.SYNOPSIS
  Launches DORefinedTool using portable PowerShell 7.4.0 and runs the main troubleshooter script.

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

param (
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ArgsPassed
)

# Get paths
$Root       = Split-Path -Parent $MyInvocation.MyCommand.Path
$PwshPath   = Join-Path -Path $Root -ChildPath "pwsh\pwsh.exe"
$MainScript = Join-Path -Path $Root -ChildPath "Invoke-DORefinedTroubleshooter.ps1"

# Validate
if (-not (Test-Path $PwshPath)) {
    Write-Host "[ERROR] Could not find pwsh.exe at: $PwshPath" -ForegroundColor Red
    Start-Sleep -Seconds 5
    exit 1
}

if (-not (Test-Path $MainScript)) {
    Write-Host "[ERROR] Could not find Invoke-DORefinedTroubleshooter.ps1 at: $MainScript" -ForegroundColor Red
    Start-Sleep -Seconds 5
    exit 1
}

# Convert args to string
$argLine = $ArgsPassed -join ' '

Write-Host "[INFO] Launching DORefinedTool with PowerShell 7.4.0..." -ForegroundColor Cyan
Write-Host "       Script: $MainScript"
Write-Host "       Args:   $argLine"

Start-Process -FilePath $PwshPath `
              -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$MainScript`" $argLine" `
              -WorkingDirectory $Root `
              -WindowStyle Normal
