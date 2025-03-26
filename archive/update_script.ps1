# Create backup of the original file
$filePath = 'D:\do\DOTroubleshooterWin32\Invoke-DoTroubleshooter.ps1'
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupPath = "${filePath}.${timestamp}.bak"
Copy-Item -Path $filePath -Destination $backupPath
Write-Output "Created backup at: $backupPath"

# Read the file content
$content = Get-Content -Path $filePath -Raw

# Add function definitions after the global buffers section
$functionsToAdd = @'

# Function to validate diagnostics zip path
function Test-DiagnosticsZipPath {
  param (
    [string]$Path
  )
  
  # Check if path is valid and doesn't contain invalid characters
  if ([string]::IsNullOrEmpty($Path) -or $Path -match '[<>:"|?*]' -or !(Test-Path -Path $Path -PathType Leaf)) {
    Write-Log "Invalid diagnostics zip path: $Path" "ERROR"
    return $false
  }
  
  # Check if the file is actually a zip file
  if (-not ($Path -match '\.(zip|cab)$')) {
    Write-Log "File is not a valid zip or cabinet file: $Path" "ERROR"
    return $false
  }
  
  return $true
}

# Function to clean up temporary files
function Remove-TemporaryFiles {
  param (
    [string]$Path
  )
  
  if (Test-Path -Path $Path) {
    try {
      Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
      Write-Log "Cleaned up temporary files at $Path" "INFO"
    } catch {
      Write-Log "Failed to clean up temporary files: $_" "WARN"
    }
  }
}
'@

$buffersSectionEnd = '# ───── LOGGING FUNCTION ─────'
$content = $content -replace [regex]::Escape($buffersSectionEnd), "$functionsToAdd`n`n$buffersSectionEnd"

# Update diagnostics zip handling
$oldDiagnosticsSection = @'
# Extract and process diagnostics data if provided
if ($DiagnosticsZip -and (Test-Path -Path $DiagnosticsZip)) {
  Write-Log "Diagnostics zip file provided: $DiagnosticsZip" "INFO"
  
  # Create a unique temporary folder for extraction
  $extractPath = Join-Path -Path $env:TEMP -ChildPath "DODiagnostics_$([Guid]::NewGuid().ToString())"
  
  # Extract the zip file
  $extractSuccess = Expand-DiagnosticsZip -ZipPath $DiagnosticsZip -ExtractPath $extractPath
  
  if ($extractSuccess) {
    # Process the extracted data
    Invoke-DiagnosticsDataProcessing -DiagnosticsPath $extractPath
  }
}
else {
  if ($DiagnosticsZip) {
    Write-Log "Specified DiagnosticsZip file not found: $DiagnosticsZip" "WARN"
  }
}
'@

$newDiagnosticsSection = @'
# Extract and process diagnostics data if provided
if ($DiagnosticsZip) {
  # Validate the diagnostics zip path
  if (Test-DiagnosticsZipPath -Path $DiagnosticsZip) {
    Write-Log "Diagnostics zip file provided: $DiagnosticsZip" "INFO"
    
    # Create a unique temporary folder for extraction
    $extractPath = Join-Path -Path $env:TEMP -ChildPath "DODiagnostics_$([Guid]::NewGuid().ToString())"
    
    # Extract the zip file
    $extractSuccess = Expand-DiagnosticsZip -ZipPath $DiagnosticsZip -ExtractPath $extractPath
    
    if ($extractSuccess) {
      # Process the extracted data
      Invoke-DiagnosticsDataProcessing -DiagnosticsPath $extractPath
    }
  }
} else {
  Write-Log "No diagnostics zip file provided" "INFO"
}
'@

$content = $content -replace [regex]::Escape($oldDiagnosticsSection), $newDiagnosticsSection

# Update the parallel processing section
$oldProcessingSection = @'
    # Process each file based on type
    foreach ($file in $doFiles) {
      $fileInfo = [PSCustomObject]@{
        FileName = $file.Name
        FilePath = $file.FullName
        FileSize = "{0:N2} KB" -f ($file.Length / 1KB)
        LastWriteTime = $file.LastWriteTime
        Category = "Unknown"
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
          Write-Log "Could not read content from $($file.Name): $_" "WARN"
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
    }
'@

$newProcessingSection = @'
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
'@

$content = $content -replace [regex]::Escape($oldProcessingSection), $newProcessingSection

# Update the cleanup section
$oldCleanupSection = @'
Write-Host "`n🟦 Delivery Optimization Troubleshooter Script Complete 🟦" -ForegroundColor Cyan
'@

$newCleanupSection = @'
# Clean up temporary files if they exist
if ($extractPath -and (Test-Path -Path $extractPath)) {
  Remove-TemporaryFiles -Path $extractPath
}

Write-Host "`n🟦 Delivery Optimization Troubleshooter Script Complete 🟦" -ForegroundColor Cyan
'@

$content = $content -replace [regex]::Escape($oldCleanupSection), $newCleanupSection

# Save the updated content
Set-Content -Path $filePath -Value $content -Encoding UTF8
Write-Output "Successfully updated the script with the requested changes."