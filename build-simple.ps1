[CmdletBinding()]
param(
    [switch]$SkipSigning
)

# Import the PowerShell Pro Tools module
try {
    Import-Module PowerShellProTools -ErrorAction Stop
}
catch {
    Write-Error "Failed to import PowerShellProTools module: $_"
    exit 1
}

# Import the configuration
$configPath = Join-Path $PSScriptRoot "package.psd1"
try {
    $config = Import-PowerShellDataFile -Path $configPath
    
    # Convert relative paths to absolute
    $config.Root = Join-Path $PSScriptRoot $config.Root.Replace(".\", "")
    $config.OutputPath = Join-Path $PSScriptRoot $config.OutputPath.Replace(".\", "")
    
    # Fix resource paths - ensure they are strings
    $resources = [string[]]@()
    foreach ($resource in $config.Package.Resources) {
        $resources += [string](Join-Path $PSScriptRoot $resource.Replace(".\", ""))
    }
    $config.Package.Resources = [string[]]$resources
    
    # Fix icon path
    if ($config.Package.Icon) {
        $config.Package.Icon = [string](Join-Path $PSScriptRoot $config.Package.Icon.Replace(".\", ""))
    }
    
    # Fix certificate path
    if ($config.Signing.CertificatePath) {
        $config.Signing.CertificatePath = [string](Join-Path $PSScriptRoot $config.Signing.CertificatePath.Replace(".\", ""))
    }
    
    # Disable signing if requested
    if ($SkipSigning) {
        $config.Signing.Enabled = $false
    }
    
    # Create output directory if it doesn't exist
    if (-not (Test-Path $config.OutputPath)) {
        New-Item -Path $config.OutputPath -ItemType Directory -Force | Out-Null
    }
    
    # Output the config for debugging
    Write-Host "Configuration:" -ForegroundColor Cyan
    Write-Host "Root: $($config.Root)" -ForegroundColor Gray
    Write-Host "OutputPath: $($config.OutputPath)" -ForegroundColor Gray
    Write-Host "Resources count: $($config.Package.Resources.Count)" -ForegroundColor Gray
    Write-Host "Resources type: $($config.Package.Resources.GetType().FullName)" -ForegroundColor Gray
    
    # Build the executable
    Write-Host "Building executable..." -ForegroundColor Cyan
    Merge-Script -Config $config
    
    # Check if build was successful
    $exePath = Join-Path $config.OutputPath "$($config.Package.OutputName).exe"
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
