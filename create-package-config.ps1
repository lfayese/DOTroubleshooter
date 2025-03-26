

# Now use this config directly with Merge-Script
try {
    Import-Module PowerShellProTools -ErrorAction Stop
    
    # Create output directory if it doesn't exist
    if (-not (Test-Path $packageConfig.OutputPath)) {
        New-Item -Path $packageConfig.OutputPath -ItemType Directory -Force | Out-Null
    }
    
    # Build the executable
    Write-Host "Building executable..." -ForegroundColor Cyan
    Merge-Script -Config $packageConfig
    
    # Check if build was successful
    $exePath = Join-Path $packageConfig.OutputPath "$($packageConfig.Package.OutputName).exe"
    if (Test-Path $exePath) {
        Write-Host "Build successful! Executable created at: $exePath" -ForegroundColor Green
    } else {
        Write-Host "Build failed. Executable not found at expected location: $exePath" -ForegroundColor Red
    }
}
catch {
    Write-Error "Build failed: $_"
    exit 1
}
