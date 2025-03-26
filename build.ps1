<#
.SYNOPSIS
    Package a PowerShell script into a signed executable using PowerShell Pro Tools.
.DESCRIPTION
    This build script handles the packaging and code signing of PowerShell scripts into standalone executables.
    It securely manages certificate passwords and provides detailed logging of the build process.

    .\build.ps1
.\build.ps1 -SkipSigning

$securePassword = ConvertTo-SecureString "test123" -AsPlainText -Force
.\build.ps1 -CertPath ".\CodeSigning\CodeSigningCert.pfx" -CertPassword $securePassword


.NOTES
    Version: 2.0
    Author: Claude
    Requires: PowerShell 5.1 or higher, PowerShell Pro Tools
#>

[CmdletBinding()]
param(
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$NoCleanup,
    
    [Parameter()]
    [switch]$Verbose,
    
    [Parameter()]
    [switch]$SkipSigning,
    
    [Parameter()]
    [string]$CertPath,
    
    [Parameter()]
    [securestring]$CertPassword
)

# Script Variables
$script:ErrorActionPreference = 'Stop'
$script:VerbosePreference = if ($Verbose) { 'Continue' } else { 'SilentlyContinue' }

# Build Configuration
$BuildConfig = @{
    MinimumPSVersion = '5.1'
    RequiredModules = @('PowerShellProTools')
    TempPath = Join-Path $env:TEMP "PSBuild_$(Get-Date -Format 'yyyyMMddHHmmss')"
    LogPath = Join-Path $PSScriptRoot 'build.log'
    DefaultCertPath = Join-Path $PSScriptRoot "CodeSigning\CodeSigningCert.pfx"
}

# Helper Functions
function Write-BuildLog {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output with colors
    $colors = @{
        Info = 'Cyan'
        Warning = 'Yellow'
        Error = 'Red'
        Success = 'Green'
    }
    
    Write-Host $logMessage -ForegroundColor $colors[$Level]
    
    # Append to log file
    $logMessage | Add-Content -Path $BuildConfig.LogPath -Encoding UTF8
}

function Test-BuildPrerequisites {
    Write-BuildLog "Checking build prerequisites..." -Level Info
    
    # Check PowerShell version
    if ($PSVersionTable.PSVersion -lt [Version]$BuildConfig.MinimumPSVersion) {
        throw "PowerShell version $($BuildConfig.MinimumPSVersion) or higher is required"
    }
    
    # Verify required modules
    foreach ($module in $BuildConfig.RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            throw "Required module '$module' is not installed"
        }
    }
    
    # Create temp directory if needed
    if (-not (Test-Path $BuildConfig.TempPath)) {
        New-Item -Path $BuildConfig.TempPath -ItemType Directory -Force | Out-Null
    }
    
    Write-BuildLog "Prerequisites check completed" -Level Success
}

function Import-PackageConfig {
    param([string]$ConfigPath)
    
    Write-BuildLog "Importing package configuration from $ConfigPath" -Level Info
    
    if (-not (Test-Path $ConfigPath)) {
        throw "Package configuration file not found at '$ConfigPath'"
    }
    
    try {
        $config = Import-PowerShellDataFile -Path $ConfigPath
        
        # Validate required configuration keys
        $requiredKeys = @('Root', 'OutputPath')
        $missingKeys = $requiredKeys.Where({ -not $config.ContainsKey($_) })
        
        if ($missingKeys) {
            throw "Missing required configuration keys: $($missingKeys -join ', ')"
        }
        
        # Expand any relative paths
        $config.Root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot $config.Root))
        $config.OutputPath = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot $config.OutputPath))
        
        # Generate JSON version for debugging
        $jsonPath = "$ConfigPath.json"
        $config | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8 -Force
        Write-BuildLog "Generated JSON configuration at $jsonPath" -Level Info
        
        return $config
    }
    catch {
        throw "Failed to import package configuration: $_"
    }
}

function Get-SecureCertificatePassword {
    param(
        [securestring]$ProvidedPassword
    )
    
    if ($ProvidedPassword) {
        Write-BuildLog "Using provided certificate password" -Level Info
        return $ProvidedPassword
    }
    
    # Check for password in environment variable (CI/CD scenario)
    $envPassword = $env:CODE_SIGNING_PASSWORD
    if ($envPassword) {
        Write-BuildLog "Using certificate password from environment variable" -Level Info
        return ConvertTo-SecureString -String $envPassword -AsPlainText -Force
    }
    
    # Prompt for password interactively
    Write-BuildLog "Prompting for certificate password..." -Level Info
    $securePassword = Read-Host -AsSecureString -Prompt "Enter certificate password"
    
    return $securePassword
}

function Update-ConfigWithSigningInfo {
    param(
        [hashtable]$Config,
        [string]$CertificatePath,
        [securestring]$CertificatePassword
    )
    
    # Ensure the Signing section exists
    if (-not $Config.ContainsKey('Signing')) {
        $Config.Signing = @{}
    }
    
    # Update certificate path
    if ($CertificatePath) {
        $Config.Signing.CertificatePath = $CertificatePath
    }
    
    # Convert secure string to plain text for the config
    if ($CertificatePassword) {
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($CertificatePassword)
        try {
            $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            $Config.Signing.CertificatePassword = $plainPassword
        }
        finally {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        }
    }
    
    # Ensure timestamp server is set
    if (-not $Config.Signing.TimeStampServer) {
        $Config.Signing.TimeStampServer = "http://timestamp.digicert.com"
    }
    
    # Enable signing
    $Config.Signing.Enabled = $true
    
    return $Config
}

function Test-Certificate {
    param(
        [string]$CertificatePath,
        [securestring]$Password
    )
    
    Write-BuildLog "Testing certificate at $CertificatePath" -Level Info
    
    if (-not (Test-Path $CertificatePath)) {
        throw "Certificate file not found at $CertificatePath"
    }
    
    try {
        # Try to load the certificate to verify it's valid
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($CertificatePath, $Password, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet)
        
        # Verify it's a code signing certificate
        $hasCodeSigningEKU = $cert.EnhancedKeyUsageList | Where-Object { $_.ObjectId -eq "1.3.6.1.5.5.7.3.3" }
        if (-not $hasCodeSigningEKU) {
            Write-BuildLog "WARNING: Certificate does not appear to be a code signing certificate" -Level Warning
        }
        
        # Check expiration
        $now = Get-Date
        if ($cert.NotBefore -gt $now) {
            Write-BuildLog "WARNING: Certificate is not yet valid (valid from $($cert.NotBefore))" -Level Warning
        }
        if ($cert.NotAfter -lt $now) {
            Write-BuildLog "WARNING: Certificate has expired (expired on $($cert.NotAfter))" -Level Warning
        }
        
        Write-BuildLog "Certificate is valid: Subject=$($cert.Subject), Expires=$($cert.NotAfter)" -Level Success
        return $true
    }
    catch {
        Write-BuildLog "Certificate test failed: $_" -Level Error
        return $false
    }
}

function New-BuildArtifact {
    param(
        [hashtable]$Config,
        [string]$OutputPath
    )
    
    Write-BuildLog "Starting build process..." -Level Info
    
    try {
        # Ensure output directory exists
        if (-not (Test-Path $Config.OutputPath)) {
            New-Item -Path $Config.OutputPath -ItemType Directory -Force | Out-Null
            Write-BuildLog "Created output directory: $($Config.OutputPath)" -Level Info
        }
        
        # Determine output executable name
        $exeName = if ($Config.Package.OutputName) {
            "$($Config.Package.OutputName).exe"
        }
        else {
            [System.IO.Path]::ChangeExtension((Split-Path -Leaf $Config.Root), ".exe")
        }
        
        $exePath = Join-Path $Config.OutputPath $exeName
        
        # Remove existing build if present
        if (Test-Path $exePath) {
            if (-not $Force) {
                $choice = Read-Host "Existing build found at $exePath. Overwrite? (Y/N)"
                if ($choice -ne 'Y') {
                    throw "Build cancelled by user"
                }
            }
            Remove-Item -Path $exePath -Force
            Write-BuildLog "Removed existing build" -Level Info
        }
        
        # Perform the build
        $buildStart = Get-Date
        
        if (Get-Command -Name 'Merge-Script' -ErrorAction SilentlyContinue) {
            Write-BuildLog "Building with PowerShell Pro Tools..." -Level Info
            Merge-Script -Config $Config
        }
        else {
            Write-BuildLog "PowerShell Pro Tools not found, attempting alternative build method..." -Level Warning
            # Implement alternative build method here if needed
            throw "No suitable build method available"
        }
        
        # Verify build
        if (Test-Path $exePath) {
            $buildDuration = (Get-Date) - $buildStart
            $fileSize = (Get-Item $exePath).Length / 1MB
            
            Write-BuildLog "Build completed successfully:" -Level Success
            Write-BuildLog "- Location: $exePath" -Level Success
            Write-BuildLog "- Duration: $($buildDuration.TotalSeconds.ToString('0.00')) seconds" -Level Success
            Write-BuildLog "- Size: $($fileSize.ToString('0.00')) MB" -Level Success
            
            # Verify signature if signing was enabled
            if ($Config.Signing.Enabled) {
                $signature = Get-AuthenticodeSignature -FilePath $exePath
                if ($signature.Status -eq 'Valid') {
                    Write-BuildLog "- Signature: Valid (Signed by $($signature.SignerCertificate.Subject))" -Level Success
                }
                else {
                    Write-BuildLog "- Signature: $($signature.Status) - Signing may have failed" -Level Warning
                }
            }
            
            return $exePath
        }
        else {
            throw "Build failed - executable not found at expected location"
        }
    }
    catch {
        Write-BuildLog "Build failed: $_" -Level Error
        throw
    }
}

function Remove-BuildArtifacts {
    if (-not $NoCleanup) {
        Write-BuildLog "Cleaning up build artifacts..." -Level Info
        
        if (Test-Path $BuildConfig.TempPath) {
            Remove-Item -Path $BuildConfig.TempPath -Recurse -Force
        }
        
        Write-BuildLog "Cleanup completed" -Level Success
    }
}

# Main Execution
try {
    Write-BuildLog "Build process started" -Level Info
    
    # Initialize build environment
    Test-BuildPrerequisites
    
    # Import and validate configuration
    $packageConfig = Import-PackageConfig -ConfigPath (Join-Path $PSScriptRoot "Package.psd1")
    
    # Handle code signing
    if (-not $SkipSigning) {
        # Determine certificate path
        $certificatePath = if ($CertPath) { 
            $CertPath 
        } elseif (Test-Path $BuildConfig.DefaultCertPath) { 
            $BuildConfig.DefaultCertPath 
        } else {
            Write-BuildLog "No certificate path provided and default certificate not found" -Level Warning
            $promptForCert = Read-Host "Do you want to specify a certificate path? (Y/N)"
            if ($promptForCert -eq 'Y') {
                Read-Host "Enter the full path to your code signing certificate"
            } else {
                Write-BuildLog "Proceeding without code signing" -Level Warning
                $SkipSigning = $true
                $null
            }
        }
        
        if (-not $SkipSigning -and $certificatePath) {
            # Get certificate password
            $certPassword = Get-SecureCertificatePassword -ProvidedPassword $CertPassword
            
            # Test the certificate
            $certValid = Test-Certificate -CertificatePath $certificatePath -Password $certPassword
            
            if ($certValid) {
                # Update config with signing information
                $packageConfig = Update-ConfigWithSigningInfo -Config $packageConfig -CertificatePath $certificatePath -CertificatePassword $certPassword
                Write-BuildLog "Code signing enabled with certificate: $certificatePath" -Level Success
            } else {
                Write-BuildLog "Certificate validation failed, proceeding without signing" -Level Warning
                $packageConfig.Signing.Enabled = $false
            }
        } else {
            $packageConfig.Signing.Enabled = $false
        }
    } else {
        $packageConfig.Signing.Enabled = $false
        Write-BuildLog "Code signing skipped as requested" -Level Info
    }
    
    # Create build artifact
    $builtExecutable = New-BuildArtifact -Config $packageConfig
    
    Write-BuildLog "Build process completed successfully" -Level Success
    
    # Show final status
    if ($packageConfig.Signing.Enabled) {
        Write-Host "`n✅ Signed executable created at: $builtExecutable" -ForegroundColor Green
    } else {
        Write-Host "`n✅ Unsigned executable created at: $builtExecutable" -ForegroundColor Green
        Write-Host "   Note: The executable is not digitally signed" -ForegroundColor Yellow
    }
}
catch {
    Write-BuildLog "Build process failed: $_" -Level Error
    exit 1
}
finally {
    # Clean up sensitive information
    if ($packageConfig -and $packageConfig.Signing -and $packageConfig.Signing.CertificatePassword) {
        $packageConfig.Signing.CertificatePassword = $null
    }
    
    Remove-BuildArtifacts
}
