Write-Host "Implementing cleanup code in Invoke-DoTroubleshooter.ps1"
$filePath = "D:\do\DOTroubleshooterWin32\Invoke-DoTroubleshooter.ps1"

# Create a backup
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupPath = "$filePath.$timestamp.bak"
Copy-Item -Path $filePath -Destination $backupPath
Write-Host "Created backup at: $backupPath"

# Read the file content
$content = Get-Content -Path $filePath -Raw

# Find the last line of the script and add cleanup code before it
$lastLine = 'Write-Host "`n🟦 Delivery Optimization Troubleshooter Script Complete 🟦" -ForegroundColor Cyan'
$cleanupCode = @'
# Clean up temporary files if they exist
if ($extractPath -and (Test-Path -Path $extractPath)) {
  Remove-TemporaryFiles -Path $extractPath
}

'@

# Replace the last line with cleanup code + last line
$updatedContent = $content.Replace($lastLine, "$cleanupCode$lastLine")

# Write the updated content back to the file
Set-Content -Path $filePath -Value $updatedContent -Force

Write-Host "Successfully added cleanup code to the script."