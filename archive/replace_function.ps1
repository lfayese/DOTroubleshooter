Write-Host "Implementing parallel processing for diagnostic files in Invoke-DoTroubleshooter.ps1"
$filePath = "D:\do\DOTroubleshooterWin32\Invoke-DoTroubleshooter.ps1"

# Create a backup
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupPath = "$filePath.$timestamp.bak"
Copy-Item -Path $filePath -Destination $backupPath
Write-Host "Created backup at: $backupPath"

# Read the file content
$fileContent = Get-Content -Path $filePath -Raw

# Create a completely new version of the Invoke-DiagnosticsDataProcessing function
$newFunction = @'
function Invoke-DiagnosticsDataProcessing {
  param (
    [string]$DiagnosticsPath
  )
  Write-Log "Processing diagnostics data from extracted folder..." "INFO"
  Show-Progress -Activity "Processing diagnostics data" -PercentComplete 15

  try {
    # Look for DO-related files in the diagnostics folder
    $doFiles = Get-ChildItem -Path $DiagnosticsPath -Recurse -File | Where-Object {
      $_.Name -match "DeliveryOptimization|DOSvc|DO_" -or 
      $_.FullName -match "\\Windows\\Logs\\DISM\\|\\WindowsUpdate\\|\\Microsoft-Windows-DeliveryOptimization"
    }

    if ($doFiles.Count -eq 0) {
      Write-Log "No Delivery Optimization related files found in diagnostics package" "WARN"
      return
    }

    Write-Log "Found $($doFiles.Count) DO-related files in diagnostics package" "SUCCESS"

    # Process files in parallel for better performance
    $doFiles | ForEach-Object -Parallel {
      $file = $_
      $Buffers = $using:Buffers
      
      $fileInfo = [PSCustomObject]@{
        FileName = $file.Name
        FilePath = $file.FullName
        FileSize = "{0:N2} KB" -f ($file.Length / 1KB)
        LastWriteTime = $file.LastWriteTime
        Category = "Unknown"
      }

      # Check file size before processing to prevent memory issues
      if ($file.Length -gt 10MB) {
        $fileInfo | Add-Member -MemberType NoteProperty -Name "SampleContent" -Value "File too large to display sample content"
        $fileInfo.Category = "Large File"
        $Buffers.DiagnosticsData.Add($fileInfo) | Out-Null
        return
      }

      # Categorize the file
      if ($file.Name -match "\.etl$") {
        $fileInfo.Category = "ETL Log"
        # Process ETL files if needed
      }
      elseif ($file.Name -match "\.log$|\.txt$") {
        $fileInfo.Category = "Text Log"
        
        # Sample the first few lines to add context
        try {
          $sampleContent = Get-Content -Path $file.FullName -TotalCount 10 -ErrorAction Stop
          $relevantEntries = $sampleContent | Where-Object { $_ -match "DeliveryOptimization|BITS|WindowsUpdate|error|warning|fail" }
          
          if ($relevantEntries) {
            $fileInfo | Add-Member -MemberType NoteProperty -Name "SampleContent" -Value ($relevantEntries -join "`n")
          }
        }
        catch {
          # Write-Log not available in parallel context
          $fileInfo | Add-Member -MemberType NoteProperty -Name "Error" -Value "Could not read content from file"
        }
      }
      elseif ($file.Name -match "\.xml$|\.json$") {
        $fileInfo.Category = "Configuration"
        # Process configuration files if needed
      }
      elseif ($file.Name -match "\.cab$|\.zip$") {
        $fileInfo.Category = "Archive"
        # Process nested archives if needed
      }
      
      # Add to diagnostics buffer
      $Buffers.DiagnosticsData.Add($fileInfo) | Out-Null
    } -ThrottleLimit 4

    # Look for Windows Update ETL logs that can be converted
    $wuEtlFiles = Get-ChildItem -Path $DiagnosticsPath -Recurse -File | Where-Object {
      $_.Name -match "\.etl$" -and $_.FullName -match "\\WindowsUpdate\\"
    }

    if ($wuEtlFiles.Count -gt 0) {
      Write-Log "Found $($wuEtlFiles.Count) Windows Update ETL logs" "INFO"
      
      # Create a temporary folder for converted logs
      $tempWULogFolder = Join-Path -Path $env:TEMP -ChildPath "WULogs_$([Guid]::NewGuid().ToString())"
      New-Item -Path $tempWULogFolder -ItemType Directory -Force | Out-Null
      
      # Try to convert the ETL files to text logs if Get-WindowsUpdateLog cmdlet is available
      if (Get-Command -Name Get-WindowsUpdateLog -ErrorAction SilentlyContinue) {
        $etlFolder = Split-Path -Path $wuEtlFiles[0].FullName -Parent
        $outputLog = Join-Path -Path $tempWULogFolder -ChildPath "WindowsUpdate.log"
        
        try {
          Get-WindowsUpdateLog -EtlPath $etlFolder -LogPath $outputLog -ErrorAction Stop
          Write-Log "Successfully converted Windows Update ETL logs to $outputLog" "SUCCESS"
          
          # Process converted log
          $wuLogEntries = Get-Content -Path $outputLog | Where-Object { 
            $_ -match "DeliveryOptimization|DO_|BITS" 
          }
          
          if ($wuLogEntries) {
            Write-Log "Found $($wuLogEntries.Count) DO-related entries in Windows Update logs" "INFO"
            
            foreach ($entry in $wuLogEntries) {
              $Buffers.DiagnosticsData.Add([PSCustomObject]@{
                FileName = "WindowsUpdate.log"
                Category = "Windows Update Log"
                Content = $entry
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
              }) | Out-Null
            }
          }
        }
        catch {
          Write-Log "Failed to convert Windows Update ETL logs: $_" "WARN"
        }
      }
      else {
        Write-Log "Get-WindowsUpdateLog cmdlet not available - cannot convert ETL files" "WARN"
      }
    }

    # Add summary of diagnostics data to Summary buffer
    $Buffers.Summary.Add([PSCustomObject]@{
      Test = "Diagnostics Data"
      Result = "$($doFiles.Count) DO-related files found"
      Status = if ($doFiles.Count -gt 0) { "PASS" } else { "WARN" }
      Impact = if ($doFiles.Count -gt 0) { "Additional diagnostic information available" } else { "Limited diagnostic information" }
    }) | Out-Null
  }
  catch {
    Write-Log "Error processing diagnostics data: $_" "ERROR"
  }
}
'@

# Find and replace the entire Invoke-DiagnosticsDataProcessing function
$pattern = "function Invoke-DiagnosticsDataProcessing \{(?:.|\n)+?^}"
$updatedContent = [System.Text.RegularExpressions.Regex]::Replace($fileContent, $pattern, $newFunction, [System.Text.RegularExpressions.RegexOptions]::Multiline)

# Write the updated content back to the file
Set-Content -Path $filePath -Value $updatedContent -Force

Write-Host "Successfully implemented parallel processing for diagnostic files."