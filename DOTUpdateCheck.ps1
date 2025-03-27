<# 
.SYNOPSIS 
  Delivery Optimization Troubleshooter 
.DESCRIPTION 
  - Parses Intune diagnostic ZIPs or live system 
  - Converts ETL to WindowsUpdate.log and extracts DO entries 
  - Checks peer ports (7680/3544), DO service health, registry 
  - Runs the official Microsoft DeliveryOptimizationTroubleshooter
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
    Summary                 = [System.Collections.ArrayList]::new()
    Recommendations         = [System.Collections.ArrayList]::new()
    DOLogErrors             = [System.Collections.ArrayList]::new()
    DOHealth                = [System.Collections.ArrayList]::new()
    PeerTests               = [System.Collections.ArrayList]::new()
    WindowsUpdate           = [System.Collections.ArrayList]::new()
    OfficialTroubleshooter  = [System.Collections.ArrayList]::new()
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
    
    # Search for all .etl files using a filter before recursing
    $etlFiles = Get-ChildItem -Path $BaseFolder -Recurse -Filter "*.etl" -File -ErrorAction SilentlyContinue
    
    if ($etlFiles -and $etlFiles.Count -gt 0) {
        Write-Log "Found $($etlFiles.Count) ETL file(s) in the extracted ZIP." "INFO"
        
        # Concatenate all ETL file paths separated by commas
        $etlPaths = $etlFiles.FullName -join ','
        $outLog = Join-Path $env:TEMP "WU_Converted.log"
        
        try {
            Write-Log "Converting ETL files to WindowsUpdate.log..." "INFO"
            Get-WindowsUpdateLog -ETLPath $etlPaths -LogPath $outLog -ErrorAction Stop
            Write-Log "Converted WindowsUpdate ETL logs." "SUCCESS"
            return $outLog
        }
        catch {
            Write-Log "Failed to convert WindowsUpdate logs: $_" "ERROR"
            
            # Fallback to older method if multiple files failed
            $etlFolder = Get-ChildItem -Path $BaseFolder -Directory -Recurse -Filter "windir_Logs_WindowsUpdate_etl*" | Select-Object -First 1
            if ($etlFolder) {
                try {
                    Write-Log "Trying alternate method with folder path..." "INFO"
                    Get-WindowsUpdateLog -ETLPath $etlFolder.FullName -LogPath $outLog -ErrorAction Stop
                    Write-Log "Converted WindowsUpdate ETL logs using folder path." "SUCCESS"
                    return $outLog
                }
                catch {
                    Write-Log "All conversion attempts failed: $_" "ERROR"
                    return $null
                }
            }
            else {
                return $null
            }
        }
    }
    else {
        # Fallback to folder-based search if no .etl files are found
        $etlFolder = Get-ChildItem -Path $BaseFolder -Directory -Recurse -Filter "windir_Logs_WindowsUpdate_etl*" | Select-Object -First 1
        if ($etlFolder) {
            $outLog = Join-Path $env:TEMP "WU_Converted.log"
            try {
                Get-WindowsUpdateLog -ETLPath $etlFolder.FullName -LogPath $outLog -ErrorAction Stop
                Write-Log "Converted WindowsUpdate ETL logs using folder path." "SUCCESS"
                return $outLog
            }
            catch {
                Write-Log "Failed to convert WindowsUpdate logs: $_" "ERROR"
                return $null
            }
        }
        else {
            Write-Log "No ETL files found in ZIP." "WARN"
            return $null
        }
    }
}

function Read-WindowsUpdateLog {
    param ([string]$WUFilePath)
    if (-not (Test-Path $WUFilePath)) {
        Write-Log "WindowsUpdate log file not found: $WUFilePath" "ERROR"
        return
    }
    # Use separate lists to store matched and full lines.
    $doBuffer = New-Object System.Collections.Generic.List[string]
    $fullBuffer = New-Object System.Collections.Generic.List[string]
    
    # Read the log in chunks (batch size 1000 lines)
    Get-Content $WUFilePath -ReadCount 1000 | ForEach-Object {
        foreach ($line in $_) {
            $fullBuffer.Add($line)
            if ($line -match "DeliveryOptimization|DO_|BITS") {
                $doBuffer.Add($line)
            }
        }
    }
    
    # Clear existing buffer (if any) and add headers and collected data
    if ($doBuffer.Count -gt 0) {
        $Buffers.WindowsUpdate.Add([PSCustomObject]@{ LogLine = "=== DO-Related Entries ===" }) | Out-Null
        foreach ($doline in $doBuffer) {
            $Buffers.WindowsUpdate.Add([PSCustomObject]@{ LogLine = $doline }) | Out-Null
        }
        $Buffers.WindowsUpdate.Add([PSCustomObject]@{ LogLine = "=== Full Log ===" }) | Out-Null
    }
    
    foreach ($line in $fullBuffer) {
        $Buffers.WindowsUpdate.Add([PSCustomObject]@{ LogLine = $line }) | Out-Null
    }
    
    Write-Log "Parsed WindowsUpdate.log with $($doBuffer.Count) DO-related entries." "INFO"
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

# === Microsoft Official Troubleshooter Integration ===
function Invoke-OfficialDOTroubleshooter {
    Write-Log "Running Delivery Optimization Troubleshooter..." "INFO"
    
    # First try to use the local script from the Scripts folder
    $localTroubleshooterPath = Join-Path -Path $ScriptRoot -ChildPath "Scripts\DeliveryOptimizationTroubleshooter.ps1"
    
    if (Test-Path $localTroubleshooterPath) {
        Write-Log "Using local DeliveryOptimizationTroubleshooter script" "INFO"
        $tempOutputFile = Join-Path $env:TEMP "DOTroubleshooterOutput_$(Get-Date -Format yyyyMMddHHmmss).txt"
        
        try {
            $scriptBlock = {
                param($scriptPath, $outputFile)
                $output = & $scriptPath
                $output | Out-File -FilePath $outputFile -Encoding utf8
            }
            $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $localTroubleshooterPath, $tempOutputFile
            Write-Log "Waiting for troubleshooter to complete..." "INFO"
            $null = Wait-Job -Job $job -Timeout 300
            
            if ($job.State -eq 'Running') {
                Write-Log "Troubleshooter is taking too long, stopping job..." "WARN"
                Stop-Job -Job $job
                Remove-Job -Job $job -Force
                Add-Recommendation -Area "Troubleshooter" -Recommendation "The troubleshooter timed out. Try running it manually." -Severity "Important"
                return
            }
            
            Receive-Job -Job $job | Out-Null
            Remove-Job -Job $job -Force
            
            if (Test-Path $tempOutputFile) {
                # Read the file only once
                $allLines = Get-Content -Path $tempOutputFile
                $troubleshooterOutput = $allLines -join "`n"
                
                foreach ($line in $allLines) {
                    $Buffers.OfficialTroubleshooter.Add([PSCustomObject]@{
                        Output = $line
                    }) | Out-Null
                }
                
                if ($troubleshooterOutput -match "Error|Warning|Failed|Issue") {
                    Add-Recommendation -Area "Troubleshooter" -Recommendation "The troubleshooter found issues. See the Troubleshooter tab for details." -Severity "Important"
                }
                
                Write-Log "Troubleshooter completed and results captured" "SUCCESS"
                Remove-Item -Path $tempOutputFile -Force -ErrorAction SilentlyContinue
            }
            else {
                Write-Log "Troubleshooter output file not found" "WARN"
                Add-Recommendation -Area "Troubleshooter" -Recommendation "Troubleshooter ran but produced no output. Try running it manually." -Severity "Important"
            }
        }
        catch {
            Write-Log "Error running local troubleshooter: $_" "ERROR"
            Add-Recommendation -Area "Troubleshooter" -Recommendation "Failed to run local troubleshooter. Check script permissions." -Severity "Important"
        }
    }
    else {
        Write-Log "Local troubleshooter not found, checking PowerShell Gallery version..." "INFO"
        $troubleshooterInstalled = Get-InstalledScript -Name "DeliveryOptimizationTroubleshooter" -ErrorAction SilentlyContinue
        
        if (-not $troubleshooterInstalled) {
            try {
                Write-Log "Installing DeliveryOptimizationTroubleshooter from PowerShell Gallery..." "INFO"
                Install-Script -Name DeliveryOptimizationTroubleshooter -Force -Scope CurrentUser -ErrorAction Stop
                Write-Log "Successfully installed DeliveryOptimizationTroubleshooter" "SUCCESS"
            }
            catch {
                Write-Log "Failed to install DeliveryOptimizationTroubleshooter: $_" "ERROR"
                Add-Recommendation -Area "Tools" -Recommendation "Manually install DeliveryOptimizationTroubleshooter from PowerShell Gallery" -Severity "Important"
                return
            }
        }
        
        $tempOutputFile = Join-Path $env:TEMP "DOTroubleshooterOutput_$(Get-Date -Format yyyyMMddHHmmss).txt"
        
        try {
            $scriptBlock = {
                param($outputFile)
                $output = & DeliveryOptimizationTroubleshooter
                $output | Out-File -FilePath $outputFile -Encoding utf8
            }
            $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $tempOutputFile
            Write-Log "Waiting for troubleshooter to complete..." "INFO"
            $null = Wait-Job -Job $job -Timeout 300
            
            if ($job.State -eq 'Running') {
                Write-Log "Troubleshooter is taking too long, stopping job..." "WARN"
                Stop-Job -Job $job
                Remove-Job -Job $job -Force
                Add-Recommendation -Area "Troubleshooter" -Recommendation "The official troubleshooter timed out. Try running it manually." -Severity "Important"
                return
            }
            
            Receive-Job -Job $job | Out-Null
            Remove-Job -Job $job -Force
            
            if (Test-Path $tempOutputFile) {
                # Read the output only once into memory
                $allLines = Get-Content -Path $tempOutputFile
                $troubleshooterOutput = $allLines -join "`n"
                
                foreach ($line in $allLines) {
                    $Buffers.OfficialTroubleshooter.Add([PSCustomObject]@{
                        Output = $line
                    }) | Out-Null
                }
                
                if ($troubleshooterOutput -match "Error|Warning|Failed|Issue") {
                    Add-Recommendation -Area "Troubleshooter" -Recommendation "The troubleshooter found issues. See the Troubleshooter tab for details." -Severity "Important"
                }
                
                Write-Log "Troubleshooter completed and results captured" "SUCCESS"
                Remove-Item -Path $tempOutputFile -Force -ErrorAction SilentlyContinue
            }
            else {
                Write-Log "Troubleshooter output file not found" "WARN"
                Add-Recommendation -Area "Troubleshooter" -Recommendation "Troubleshooter ran but produced no output. Try running it manually." -Severity "Important"
            }
        }
        catch {
            Write-Log "Error running official troubleshooter: $_" "ERROR"
            Add-Recommendation -Area "Troubleshooter" -Recommendation "Failed to run official troubleshooter. Try running it manually." -Severity "Important"
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
        
        # Create Excel package
        $excel = $Buffers.DOHealth | Export-Excel -Path $excelPath -WorksheetName "DO_Health" -AutoSize -PassThru
        $Buffers.PeerTests | Export-Excel -Path $excelPath -WorksheetName "Peer_Ports" -AutoSize -ExcelPackage $excel -PassThru
        $Buffers.Recommendations | Export-Excel -Path $excelPath -WorksheetName "Recommendations" -AutoSize -ExcelPackage $excel -PassThru
        
        if ($Buffers.WindowsUpdate.Count -gt 0) {
            $Buffers.WindowsUpdate | Export-Excel -Path $excelPath -WorksheetName "WindowsUpdate.log" -AutoSize -ExcelPackage $excel -PassThru
        }
        
        # Add Troubleshooter worksheet if we have data
        if ($Buffers.OfficialTroubleshooter.Count -gt 0) {
            $worksheet = $excel.Workbook.Worksheets.Add("Troubleshooter")
            $row = 1
            
            # Add header
            $worksheet.Cells["A$row"].Value = "Delivery Optimization Troubleshooter Results"
            $worksheet.Cells["A$row"].Style.Font.Bold = $true
            $worksheet.Cells["A$row"].Style.Font.Size = 14
            $row++
            $row++
            
            # Add output lines
            foreach ($line in $Buffers.OfficialTroubleshooter) {
                $worksheet.Cells["A$row"].Value = $line.Output
                
                # Highlight issues in red
                if ($line.Output -match "Error|Warning|Failed|Issue|Critical") {
                    $worksheet.Cells["A$row"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                }
                # Highlight success in green
                elseif ($line.Output -match "Success|Passed|Healthy|OK") {
                    $worksheet.Cells["A$row"].Style.Font.Color.SetColor([System.Drawing.Color]::Green)
                }
                
                $row++
            }
            
            # Format worksheet
            $worksheet.Column(1).Width = 120
            $worksheet.View.FreezePanes(3, 1)
        }
        
        # Save and close the Excel package
        $excel.Save()
        $excel.Dispose()
        
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
Invoke-OfficialDOTroubleshooter

# Generate report
Export-DOReport
Remove-Jobs
Write-Host "`n✅ Troubleshooting Complete!" -ForegroundColor Green