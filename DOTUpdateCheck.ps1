<# 
.SYNOPSIS 
  Delivery Optimization Troubleshooter 
.DESCRIPTION 
  - Parses Intune diagnostic ZIPs or live system 
  - Converts ETL to WindowsUpdate.log and extracts DO entries 
  - Checks peer ports (7680/3544), DO service health, registry 
  - Builds full Excel report with recommendations and health status 
.PARAMETER DiagnosticsZip 
  Optional ZIP file to extract and analyze 
.PARAMETER OutputPath 
  Path where Excel report will be saved (default: Desktop) 
.PARAMETER Show 
  Optional switch to open the Excel report after generation 
.EXAMPLE 
  .\DOTUpdateCheck.ps1 -DiagnosticsZip "C:\Diag\IntuneLog.zip" -Show #Opens the Excel report after generation 
#>
param (
    [string]$OutputPath = "$env:USERPROFILE\Desktop",
    [switch]$Show,
    [string]$DiagnosticsZip
)
# === Path Management ===
if ($PSScriptRoot) {
    $ScriptRoot = $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptRoot = (Get-Location).Path
}
# === Module Import ===
$modulePath = Join-Path -Path $ScriptRoot -ChildPath "Modules\ImportExcel"
Write-Host "[DEBUG] Module path: $modulePath" -ForegroundColor Yellow
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    try {
        Import-Module -Name $modulePath -Force
    } catch {
        Write-Host "[ERROR] Failed to import ImportExcel module: $_" -ForegroundColor Red
        Write-Host "[INFO] Will try to use the module if it's installed globally" -ForegroundColor Cyan
    }
} else {
    Import-Module ImportExcel
}
# === Buffers ===
$Buffers = @{
    Summary         = [System.Collections.ArrayList]::new()
    Recommendations = [System.Collections.ArrayList]::new()
    DOLogErrors     = [System.Collections.ArrayList]::new()
    DOHealth        = [System.Collections.ArrayList]::new()
    PeerTests       = [System.Collections.ArrayList]::new()
    WindowsUpdate   = [System.Collections.ArrayList]::new()
}
# === Logging ===
function Write-Log {
    param ([string]$Message, [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]$Level = "INFO")
    $color = switch ($Level) {
        "ERROR"   { "Red" }
        "WARN"    { "Yellow" }
        "SUCCESS" { "Green" }
        default   { "Cyan" }
    }
    Write-Host "[$Level] $Message" -ForegroundColor $color
}
function Add-Recommendation {
    param (
        [string]$Area,
        [string]$Recommendation,
        [ValidateSet("Critical", "Important", "Informational")]$Severity = "Informational"
    )
    $Buffers.Recommendations.Add([PSCustomObject]@{
        Area           = $Area
        Recommendation = $Recommendation
        Severity       = $Severity
    }) | Out-Null
}
function Get-DOErrorsTable {
    @(
        @{ Code = "0x80D01001"; Description = "Service error"; Recommendation = "Restart Delivery Optimization service." },
        @{ Code = "0x80D02002"; Description = "Timeout"; Recommendation = "Check internet connection or proxy." },
        @{ Code = "0x80D02004"; Description = "Empty job"; Recommendation = "Ensure content is available." }
    ) | ForEach-Object {
        [PSCustomObject]@{
            ErrorCode      = $_.Code
            Description    = $_.Description
            Recommendation = $_.Recommendation
        }
    }
}
function Remove-Jobs {
    Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
}
# === ZIP/ETL Functions ===
function Expand-Zip {
    param ([string]$Path)
    # Use a unique temporary folder
    $Script:ExtractPath = Join-Path $env:TEMP "DOExtract_$([guid]::NewGuid())"
    if (Test-Path $Script:ExtractPath) {
        Remove-Item -Path $Script:ExtractPath -Recurse -Force
    }
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($Path, $Script:ExtractPath)
        Write-Log "Extracted ZIP to: $Script:ExtractPath" "SUCCESS"
        return $Script:ExtractPath
    } catch {
        Write-Log "Failed to extract ZIP: $_" "ERROR"
        return $null
    }
}
function Convert-WindowsUpdateLog {
    param ($BaseFolder)
    if (-not $BaseFolder -or -not (Test-Path $BaseFolder)) {
        Write-Log "Invalid folder path for ETL extraction." "ERROR"
        return $null
    }
    # Use a filter to narrow down folders (assumes the folder name starts with 'windir_Logs_WindowsUpdate_etl')
    $etlFolder = Get-ChildItem -Path $BaseFolder -Directory -Recurse -Filter "windir_Logs_WindowsUpdate_etl*" | Select-Object -First 1
    if ($etlFolder) {
        $outLog = Join-Path $env:TEMP "WU_Converted.log"
        try {
            Get-WindowsUpdateLog -ETLPath $etlFolder.FullName -LogPath $outLog -ErrorAction Stop
            Write-Log "Converted WindowsUpdate ETL logs." "SUCCESS"
            return $outLog
        } catch {
            Write-Log "Failed to convert WindowsUpdate logs: $_" "ERROR"
            return $null
        }
    } else {
        Write-Log "ETL folder not found in ZIP." "WARN"
        return $null
    }
}
# === Optimized WindowsUpdate Log Reading ===
function Read-WindowsUpdateLog {
    param ([string]$WUFilePath)
    if (-not (Test-Path $WUFilePath)) {
        Write-Log "WindowsUpdate log file not found: $WUFilePath" "ERROR"
        return
    }
    # Use batch processing to avoid reading the entire file into memory at once.
    $doCollected = New-Object System.Collections.Generic.List[string]
    $fullCollected = New-Object System.Collections.Generic.List[string]
    # Get-Content with a higher ReadCount to read lines in chunks
    Get-Content $WUFilePath -ReadCount 1000 | ForEach-Object {
        foreach ($line in $_) {
            $fullCollected.Add($line)
            if ($line -match "DeliveryOptimization|DO_|BITS") {
                $doCollected.Add($line)
            }
        }
    }
    if ($doCollected.Count -gt 0) {
        $Buffers.WindowsUpdate.Add([PSCustomObject]@{ LogLine = "=== DO-Related Entries ===" })
        foreach ($doline in $doCollected) {
            $Buffers.WindowsUpdate.Add([PSCustomObject]@{ LogLine = $doline })
        }
        $Buffers.WindowsUpdate.Add([PSCustomObject]@{ LogLine = "=== Full Log ===" })
    }
    foreach ($line in $fullCollected) {
        $Buffers.WindowsUpdate.Add([PSCustomObject]@{ LogLine = $line })
    }
    Write-Log "Parsed WindowsUpdate.log with $($doCollected.Count) DO-related entries." "INFO"
}
# === Diagnostics Functions ===
function Get-DOHealthStatus {
    try {
        $status = Get-DeliveryOptimizationStatus
        $desc = switch ($status.DODownloadMode) {
            0 { "HTTP only" }
            1 { "LAN only" }
            2 { "LAN + Internet" }
            3 { "LAN + Internet + Group" }
            99 { "Fallback mode" }
            default { "Unknown" }
        }
        $Buffers.DOHealth.Add([PSCustomObject]@{
            Mode               = $status.DODownloadMode
            Description        = $desc
            Peers              = $status.NumberOfPeers
            MaxCacheSizeMB     = $status.MaxCacheSize
            CurrentCacheSizeMB = $status.CurrentCacheSize
            PeerCachingAllowed = $status.PeerCachingAllowed
        })
        if ($status.DODownloadMode -eq 0) {
            Add-Recommendation -Area "Config" -Recommendation "Enable peer caching for bandwidth savings." -Severity "Important"
        } elseif ($status.DODownloadMode -eq 99) {
            Add-Recommendation -Area "Service" -Recommendation "Restart DO service - it is in fallback mode." -Severity "Critical"
        }
    } catch {
        Write-Log "Failed to get DO status: $_" "ERROR"
        Add-Recommendation -Area "Service" -Recommendation "Ensure Delivery Optimization service is running." -Severity "Critical"
    }
}
function Test-DOConnectivity {
    $ports = @(7680, 3544)
    $targets = @("127.0.0.1", "$env:COMPUTERNAME")
    foreach ($port in $ports) {
        foreach ($target in $targets) {
            try {
                $result = Test-NetConnection -ComputerName $target -Port $port -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                $Buffers.PeerTests.Add([PSCustomObject]@{
                    Target  = $target
                    Port    = $port
                    TCP     = $result.TcpTestSucceeded
                    Ping    = $result.PingSucceeded
                    Status  = if ($result.TcpTestSucceeded) { "PASS" } else { "FAIL" }
                    Impact  = if (-not $result.TcpTestSucceeded) { "Peer sharing may not work." } else { "OK" }
                })
                if (-not $result.TcpTestSucceeded) {
                    Add-Recommendation -Area "Firewall" -Recommendation "Open TCP port $port for DO peer sharing." -Severity "Important"
                }
            } catch {
                Write-Log "Failed to test connectivity to $target $port - $_" "ERROR"
                $Buffers.PeerTests.Add([PSCustomObject]@{
                    Target = $target
                    Port   = $port
                    TCP    = $false
                    Ping   = $false
                    Status = "ERROR"
                    Impact = "Connection test failed."
                })
            }
        }
    }
}
# === Excel Report ===
function Export-DOReport {
    try {
        if (-not (Test-Path $OutputPath)) {
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        }
        $excelPath = Join-Path $OutputPath "DO_Report_$(Get-Date -Format yyyyMMdd_HHmmss).xlsx"
        $Buffers.DOHealth        | Export-Excel -Path $excelPath -WorksheetName "DO_Health" -AutoSize
        $Buffers.PeerTests       | Export-Excel -Path $excelPath -WorksheetName "Peer_Ports" -AutoSize -Append
        $Buffers.Recommendations | Export-Excel -Path $excelPath -WorksheetName "Recommendations" -AutoSize -Append
        if ($Buffers.WindowsUpdate.Count -gt 0) {
            $Buffers.WindowsUpdate | Export-Excel -Path $excelPath -WorksheetName "WindowsUpdate.log" -AutoSize -Append
        }
        Write-Host "`n📊 Excel Report saved to: $excelPath" -ForegroundColor Cyan
        if ($Show -and (Test-Path $excelPath)) { 
            Invoke-Item $excelPath 
        }
    } catch {
        Write-Log "Failed to generate Excel report: $_" "ERROR"
        Write-Host "Please ensure the ImportExcel module is properly installed." -ForegroundColor Yellow
    }
}
# === MAIN ===
Write-Host "`n🟦 Running Delivery Optimization Troubleshooter..." -ForegroundColor Cyan
if ($DiagnosticsZip -and (Test-Path $DiagnosticsZip)) {
    $Extracted = Expand-Zip -Path $DiagnosticsZip
    if ($Extracted) {
        $wuLogPath = Convert-WindowsUpdateLog -BaseFolder $Extracted
        if ($wuLogPath) { 
            Read-WindowsUpdateLog -WUFilePath $wuLogPath 
        }
    }
}
# Run diagnostics
Get-DOHealthStatus
Test-DOConnectivity
# Generate report
Export-DOReport
Remove-Jobs
Write-Host "`n✅ Troubleshooting Complete!" -ForegroundColor Green