# Create backup of the original file
$filePath = 'D:\do\DOTroubleshooterWin32\Invoke-DoTroubleshooter.ps1'
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupPath = "${filePath}.${timestamp}.bak"
Copy-Item -Path $filePath -Destination $backupPath
Write-Output "Created backup at: $backupPath"

# Read the file content
$content = Get-Content -Path $filePath -Raw

# Update the parallel processing section - using a different approach with line numbers
$lines = Get-Content -Path $filePath
$startLineNum = 0
$endLineNum = 0

# Find the start and end of the section to replace
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match "# Process each file based on type") {
        $startLineNum = $i
    }
    if ($startLineNum -gt 0 -and $lines[$i] -match "# Add to diagnostics buffer" -and $lines[$i+1] -match "\$Buffers\.DiagnosticsData\.Add\(\$fileInfo\) \| Out-Null") {
        $endLineNum = $i + 2
        break
    }
}

if ($startLineNum -gt 0 -and $endLineNum -gt 0) {
    Write-Output "Found section to replace from line $startLineNum to $endLineNum"
    
    # Create the new content
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

    # Replace the section
    $beforeSection = $lines[0..($startLineNum-1)]
    $afterSection = $lines[$endLineNum..($lines.Count-1)]
    $newContent = $beforeSection + $newProcessingSection + $afterSection
    
    # Save the updated content
    $newContent | Set-Content -Path $filePath -Encoding UTF8
    Write-Output "Successfully updated the parallel processing section."
} else {
    Write-Output "Could not find the section to replace. Check the script manually."
}