# Create a PowerShell script to directly modify the file using line numbers
$filePath = 'D:\do\DOTroubleshooterWin32\Invoke-DoTroubleshooter.ps1'
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupPath = "${filePath}.${timestamp}.bak"

# Create a backup
Copy-Item -Path $filePath -Destination $backupPath
Write-Output "Created backup at: $backupPath"

# Read the file content as an array of lines
$lines = Get-Content -Path $filePath

# Find the line that starts with "    # Process each file based on type"
$startLine = 0
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i].Trim() -eq "    # Process each file based on type") {
        $startLine = $i
        break
    }
}

if ($startLine -eq 0) {
    Write-Output "Could not find the starting line for the section to replace."
    exit
}

# Find the end of the foreach block
$endLine = 0
$braceCount = 0
$foundFirstBrace = $false
for ($i = $startLine; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]
    
    # Count opening braces
    $openBraces = ($line | Select-String -Pattern "{" -AllMatches).Matches.Count
    # Count closing braces
    $closeBraces = ($line | Select-String -Pattern "}" -AllMatches).Matches.Count
    
    if (!$foundFirstBrace -and $line.Contains("{")) {
        $foundFirstBrace = $true
    }
    
    if ($foundFirstBrace) {
        $braceCount += $openBraces - $closeBraces
        
        if ($braceCount -eq 0 -and $foundFirstBrace) {
            $endLine = $i
            break
        }
    }
}

if ($endLine -eq 0) {
    Write-Output "Could not find the ending line for the section to replace."
    exit
}

Write-Output "Found section to replace from line $startLine to $endLine"

# Create the new section with parallel processing
$newSection = @"
    # Process files in parallel for better performance
    `$doFiles | ForEach-Object -Parallel {
      `$file = `$_
      `$Buffers = `$using:Buffers
      
      `$fileInfo = [PSCustomObject]@{
        FileName = `$file.Name
        FilePath = `$file.FullName
        FileSize = "{0:N2} KB" -f (`$file.Length / 1KB)
        LastWriteTime = `$file.LastWriteTime
        Category = "Unknown"
      }

      # Check file size before processing to prevent memory issues
      if (`$file.Length -gt 10MB) {
        `$fileInfo | Add-Member -MemberType NoteProperty -Name "SampleContent" -Value "File too large to display sample content"
        `$fileInfo.Category = "Large File"
        `$Buffers.DiagnosticsData.Add(`$fileInfo) | Out-Null
        return
      }

      # Categorize the file
      if (`$file.Name -match "\.etl`$") {
        `$fileInfo.Category = "ETL Log"
        # Process ETL files if needed
      }
      elseif (`$file.Name -match "\.log`$|\.txt`$") {
        `$fileInfo.Category = "Text Log"
        
        # Sample the first few lines to add context
        try {
          `$sampleContent = Get-Content -Path `$file.FullName -TotalCount 10 -ErrorAction Stop
          `$relevantEntries = `$sampleContent | Where-Object { `$_ -match "DeliveryOptimization|BITS|WindowsUpdate|error|warning|fail" }
          
          if (`$relevantEntries) {
            `$fileInfo | Add-Member -MemberType NoteProperty -Name "SampleContent" -Value (`$relevantEntries -join "`n")
          }
        }
        catch {
          # Write-Log not available in parallel context
          `$fileInfo | Add-Member -MemberType NoteProperty -Name "Error" -Value "Could not read content from file"
        }
      }
      elseif (`$file.Name -match "\.xml`$|\.json`$") {
        `$fileInfo.Category = "Configuration"
        # Process configuration files if needed
      }
      elseif (`$file.Name -match "\.cab`$|\.zip`$") {
        `$fileInfo.Category = "Archive"
        # Process nested archives if needed
      }
      
      # Add to diagnostics buffer
      `$Buffers.DiagnosticsData.Add(`$fileInfo) | Out-Null
    } -ThrottleLimit 4
"@

# Replace the old section with the new one
$newContent = $lines[0..($startLine-1)] + $newSection + $lines[($endLine+1)..($lines.Count-1)]

# Write the updated content back to the file
$newContent | Set-Content -Path $filePath -Force

Write-Output "Successfully updated the script with parallel processing."