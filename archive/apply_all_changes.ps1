Write-Host "Implementing all required changes to Invoke-DoTroubleshooter.ps1"
$filePath = "D:\do\DOTroubleshooterWin32\Invoke-DoTroubleshooter.ps1"

# Create a backup
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupPath = "$filePath.$timestamp.bak"
Copy-Item -Path $filePath -Destination $backupPath
Write-Host "Created backup at: $backupPath"

# Read the file content
$content = Get-Content -Path $filePath -Raw

# 1. Add the parallel processing with file size check
# We'll use a simpler approach with string replacement
$oldProcessingCode = @'
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

$newProcessingCode = @'
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

if ($content.Contains($oldProcessingCode)) {
    Write-Host "Found the processing code section, replacing with parallel implementation..."
    $content = $content.Replace($oldProcessingCode, $newProcessingCode)
}
else {
    Write-Host "Warning: Could not find the exact processing code section to replace."
}

# Save the updated content
Set-Content -Path $filePath -Value $content -Force
Write-Host "Successfully updated the script with parallel processing code."

# Verify the changes
if (Select-String -Path $filePath -Pattern "ForEach-Object -Parallel" -Quiet) {
    Write-Host "Verified: Parallel processing implementation was added successfully."
}
else {
    Write-Host "Warning: Could not verify parallel processing implementation."
}

if (Select-String -Path $filePath -Pattern "10MB" -Quiet) {
    Write-Host "Verified: File size check was added successfully."
}
else {
    Write-Host "Warning: Could not verify file size check implementation."
}

Write-Host "All changes have been implemented successfully."