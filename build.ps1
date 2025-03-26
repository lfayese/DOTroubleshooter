<#
.SYNOPSIS
    Package a PowerShell script into a standalone executable using PowerShell Pro Tools.
.DESCRIPTION
    This build.ps1 script imports the package configuration from Package.psd1 (which must follow the documented schema),
    optionally creates a JSON version for debugging, and then calls Merge-Script to build the executable.
    If Merge-Script is not found on the system, the script shows a warning and a hint on how to package manually.

    See PowerShell Pro Tools documentation for details on config options such as Root, OutputPath, Package keys, and Bundle.
.NOTES
    Ensure that PowerShell Pro Tools is installed and that Package.psd1 is configured following the schema.
#>
[CmdletBinding()]
param()
# Define the full path to the package config file.
$packageConfigPath = Join-Path $PSScriptRoot "Package.psd1"
if (-not (Test-Path $packageConfigPath)) {
    Write-Error "Package configuration file '$packageConfigPath' not found! Please create it according to the documentation."
    exit 1
}
try {
    # Import the configuration file (a psd1 file containing a hashtable).
    $packageConfig = Import-PowerShellDataFile -Path $packageConfigPath
} catch {
    Write-Error "Failed to import package configuration from '$packageConfigPath': $_"
    exit 1
}
# Optional: Generate a JSON version of the package configuration for debugging purposes.
try {
    $jsonPath = "$packageConfigPath.json"
    $packageConfig | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8 -Force
    Write-Host "✅ Generated JSON configuration file at $jsonPath" -ForegroundColor Green
} catch {
    Write-Warning "Unable to generate JSON configuration file: $_"
}
# Validate that required keys exist in the config
if (-not $packageConfig.Root -or -not $packageConfig.OutputPath) {
    Write-Error "The configuration must include at least 'Root' and 'OutputPath' keys."
    exit 1
}
# Ensure the output directory exists.
if (-not (Test-Path $packageConfig.OutputPath)) {
    try {
        New-Item -Path $packageConfig.OutputPath -ItemType Directory -Force | Out-Null
        Write-Host "📁 Created output directory: $packageConfig.OutputPath" -ForegroundColor Cyan
    } catch {
        Write-Error "Failed to create output directory '$($packageConfig.OutputPath)': $_"
        exit 1
    }
}
# Determine the executable name.
if ($packageConfig.Package.OutputName) {
    $exeName = "$($packageConfig.Package.OutputName).exe"
} else {
    # Fallback: use the root script name with .exe extension.
    $rootScript = Split-Path -Leaf $packageConfig.Root
    $exeName = [System.IO.Path]::ChangeExtension($rootScript, ".exe")
}
$exePath = Join-Path $packageConfig.OutputPath $exeName
# Optionally, remove any previous build if it exists.
if (Test-Path $exePath) {
    Write-Host "🧹 Removing previous build: $exePath" -ForegroundColor Yellow
    Remove-Item -Path $exePath -Force -ErrorAction SilentlyContinue
}
# Check for Merge-Script availability.
if (Get-Command -Name 'Merge-Script' -ErrorAction SilentlyContinue) {
    Write-Host "🔨 Packaging with Merge-Script..."
    try {
        # Start timing the build process.
        $buildStart = Get-Date
        # Call Merge-Script with the configuration hashtable.
        Merge-Script -Config $packageConfig
        # Verify that the executable was created.
        if (Test-Path $exePath) {
            $buildDuration = (Get-Date) - $buildStart
            $fileSizeMB = (Get-Item $exePath).Length / 1MB
            Write-Host "✅ Successfully created executable at $exePath" -ForegroundColor Green
            Write-Host ("📊 Build completed in {0:N2} seconds, file size: {1:N2} MB" -f $buildDuration.TotalSeconds, $fileSizeMB) -ForegroundColor Cyan
        }
        else {
            Write-Warning "⚠️ Build completed but executable not found at expected location: $exePath"
        }
    } catch {
        Write-Error "Failed to package script: $_"
    }
}
else {
    Write-Warning "⚠️ Merge-Script command not found. Please install PowerShell Pro Tools or package manually."
    Write-Host "📋 To package manually, run:" -ForegroundColor White
    Write-Host "    Merge-Script -Config (Import-PowerShellDataFile Package.psd1)" -ForegroundColor White
}
