<#
.SYNOPSIS
Deploys and runs the DO Troubleshooter with embedded resources and PowerShell 7.

.DESCRIPTION
This script is designed to deploy and execute the DO Troubleshooter. It handles resource extraction, PowerShell 7 detection, and elevation to administrator privileges if required. The script can run in both script mode and executable mode, supporting embedded resources or file-based resources.

.PARAMETER Show
Displays a message box indicating that the DO Troubleshooter is launching.

.PARAMETER OutputPath
Specifies the directory where the output reports and logs will be saved. Defaults to "$env:USERPROFILE\DOReports".

.PARAMETER DiagnosticsZip
Specifies the path to a diagnostics zip file to be used by the troubleshooter.

.EXAMPLE
.\Deploy-Do.ps1 -Show

Launches the DO Troubleshooter with a message box and uses the default output path.

.EXAMPLE
.\Deploy-Do.ps1 -OutputPath "C:\CustomReports"

Runs the DO Troubleshooter and saves the output to the specified custom directory.

.EXAMPLE
.\Deploy-Do.ps1 -DiagnosticsZip "C:\Diagnostics\Logs.zip"

Runs the DO Troubleshooter with the specified diagnostics zip file.

.EXAMPLE
.\Deploy-Do.ps1 -Show -OutputPath "C:\CustomReports" -DiagnosticsZip "C:\Diagnostics\Logs.zip"

Launches the DO Troubleshooter with a message box, saves the output to the specified custom directory, and uses the specified diagnostics zip file.

.NOTES
- The script requires administrator privileges to execute certain operations.
- If PowerShell 7 is not installed, the script attempts to use an embedded version or locate it in standard installation paths.
- The script cleans up temporary resources after execution.

#>
[CmdletBinding()]
param (
    [switch]$Show,
    [string]$OutputPath = "$env:USERPROFILE\DOReports",
    [string]$DiagnosticsZip
)

# Initialize variables
$ExeDirectory = $null
if ($MyInvocation.MyCommand.Path) {
    $ExeDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
} elseif ($env:EXEPATH) {
    # If running from an executable built with PowerShell Pro Tools
    $ExeDirectory = Split-Path -Parent $env:EXEPATH
} else {
    $ExeDirectory = $PSScriptRoot
}

$ExtractDir = Join-Path $env:TEMP "DODeploy_$([guid]::NewGuid().ToString())"
$LogFile = Join-Path $OutputPath "DODeployment.log"

# Create output directory
New-Item -ItemType Directory -Path $OutputPath -Force -ErrorAction SilentlyContinue | Out-Null

# Helper functions
function Write-Log {
    param ([string]$Message)
    $timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    "$timestamp - $Message" | Tee-Object -FilePath $LogFile -Append
}

function Test-IsAdministrator {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Expand-Resources {
    Write-Log "Extracting resources to $ExtractDir..."
    New-Item -Path $ExtractDir -ItemType Directory -Force | Out-Null
    
    try {
        # Handle both script mode and executable mode for resource extraction
        if ($null -ne [Embedded] -and (Get-Member -InputObject ([Embedded]) -Static -Name GetNames)) {
            # Running from executable with embedded resources
            $resources = [Embedded]::GetNames()
            foreach ($resource in $resources) {
                $targetPath = Join-Path $ExtractDir $resource
                $targetDir = Split-Path $targetPath -Parent
                
                if (-not (Test-Path $targetDir)) {
                    New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
                }
                
                [System.IO.File]::WriteAllBytes($targetPath, [Embedded]::Get($resource))
            }
        } else {
            # Running as script - copy resources from script directory
            Write-Log "Embedded class not found, using file copy mode..."
            
            # Copy required files from script directory
            $resourcesToExtract = @(
                "Invoke-DoTroubleshooter.ps1",
                "PowerShell-7.5.0-win-x64.zip",
                "PSTools"
            )
            
            foreach ($resource in $resourcesToExtract) {
                $sourcePath = Join-Path $ExeDirectory $resource
                $targetPath = Join-Path $ExtractDir $resource
                
                if (Test-Path -Path $sourcePath -PathType Container) {
                    Copy-Item -Path $sourcePath -Destination $ExtractDir -Recurse -Force
                    Write-Log "Copied directory: $resource"
                } elseif (Test-Path -Path $sourcePath -PathType Leaf) {
                    $targetDir = Split-Path $targetPath -Parent
                    if (-not (Test-Path $targetDir)) {
                        New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
                    }
                    Copy-Item -Path $sourcePath -Destination $targetPath -Force
                    Write-Log "Copied file: $resource"
                } else {
                    Write-Log "Resource not found: $sourcePath"
                }
            }
            
            # Extract Modules directory if it exists
            $modulesSource = Join-Path $ExeDirectory "Modules"
            if (Test-Path $modulesSource) {
                Copy-Item -Path $modulesSource -Destination $ExtractDir -Recurse -Force
                Write-Log "Copied Modules directory"
            }
            
            # Extract Scripts directory if it exists
            $scriptsSource = Join-Path $ExeDirectory "Scripts"
            if (Test-Path $scriptsSource) {
                Copy-Item -Path $scriptsSource -Destination $ExtractDir -Recurse -Force
                Write-Log "Copied Scripts directory"
            }
        }
        
        Write-Log "Resource extraction complete"
    }
    catch {
        Write-Log "ERROR during extraction: $_"
        throw
    }
}

function Get-PowerShell7 {
    # Try system-installed PowerShell 7
    $pwsh = Get-Command -Name pwsh -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
    if ($pwsh -and (& $pwsh -NoProfile -Command { $PSVersionTable.PSVersion.Major -ge 7 })) {
        Write-Log "Using system PowerShell 7: $pwsh"
        return $pwsh
    }
    
    # Try embedded PowerShell 7.5.0 first (newer version)
    $localPwsh75 = Join-Path $ExtractDir "PowerShell-7.5.0-win-x64\pwsh.exe"
    if (Test-Path -Path $localPwsh75) {
        Write-Log "Using embedded PowerShell 7.5.0: $localPwsh75"
        return $localPwsh75
    }
    
    # Fall back to 7.4.0 if needed
    $localPwsh74 = Join-Path $ExtractDir "PowerShell-7.4.0-win-x64\pwsh.exe"
    if (Test-Path -Path $localPwsh74) {
        Write-Log "Using embedded PowerShell 7.4.0: $localPwsh74"
        return $localPwsh74
    }
    
    # If embedded PowerShell zip exists, extract it
    $ps75zip = Join-Path $ExtractDir "PowerShell-7.5.0-win-x64.zip"
    if (Test-Path -Path $ps75zip) {
        $extractPath = Join-Path $ExtractDir "PowerShell-7.5.0-win-x64"
        Write-Log "Extracting PowerShell 7.5.0 from zip..."
        
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::ExtractToDirectory($ps75zip, $extractPath)
            
            if (Test-Path -Path $localPwsh75) {
                Write-Log "Using freshly extracted PowerShell 7.5.0: $localPwsh75"
                return $localPwsh75
            }
        } catch {
            Write-Log "Error extracting PowerShell 7.5.0: $_"
        }
    }
    
    # Last resort - look for PowerShell 7 in standard installation paths
    $possiblePaths = @(
        "${env:ProgramFiles}\PowerShell\7\pwsh.exe",
        "${env:ProgramFiles(x86)}\PowerShell\7\pwsh.exe",
        "$env:LocalAppData\Microsoft\PowerShell\7\pwsh.exe"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            Write-Log "Found alternative PowerShell 7 path: $path"
            return $path
        }
    }
    
    throw "PowerShell 7 not found. Please install PowerShell 7 or ensure the embedded PowerShell package is included."
}

# Main script execution
try {
    Add-Type -AssemblyName System.Windows.Forms
    
    if ($Show) {
        [System.Windows.Forms.MessageBox]::Show("Launching DO Troubleshooter...", "DO Launcher", "OK", "Information")
    }
    
    # Extract embedded resources
    Expand-Resources
    
    # Get PowerShell 7 path
    $pwshPath = Get-PowerShell7
    
    # Elevate if not running as admin
    if (-not (Test-IsAdministrator)) {
        Write-Log "Elevating script as administrator..."
        
        # Determine the path to use for elevation
        $elevationPath = $null
        if ($null -ne $env:EXEPATH) {
            # If running from executable
            $elevationPath = $env:EXEPATH
            $arguments = "-OutputPath `"$OutputPath`""
            if ($Show) { $arguments += " -Show" }
            if ($DiagnosticsZip) { $arguments += " -DiagnosticsZip `"$DiagnosticsZip`"" }
        } else {
            # If running as script
            $elevationPath = $pwshPath
            $arguments = "-ExecutionPolicy Bypass -NoProfile -File `"$($MyInvocation.MyCommand.Definition)`" -OutputPath `"$OutputPath`""
            if ($Show) { $arguments += " -Show" }
            if ($DiagnosticsZip) { $arguments += " -DiagnosticsZip `"$DiagnosticsZip`"" }
        }
        
        Start-Process -FilePath $elevationPath -ArgumentList $arguments -Verb RunAs
        exit
    }
    
    # Run troubleshooter
    $scriptPath = Join-Path $ExtractDir "Invoke-DoTroubleshooter.ps1"
    $psExec = Join-Path $ExtractDir "PSTools\PsExec64.exe"
    $cmdArgs = "& '$scriptPath' -OutputPath '$OutputPath'"
    if ($Show) { $cmdArgs += " -Show" }
    if ($DiagnosticsZip) { $cmdArgs += " -DiagnosticsZip '$DiagnosticsZip'" }
    
    if (Test-Path $psExec) {
        Write-Log "Running troubleshooter via PsExec..."
        & $psExec -accepteula -s -i "$pwshPath" -ExecutionPolicy Bypass -NoProfile -Command $cmdArgs
    } else {
        Write-Log "Running troubleshooter as Admin (PsExec not found)..."
        & "$pwshPath" -ExecutionPolicy Bypass -NoProfile -Command $cmdArgs
    }
    
    if ($LASTEXITCODE -eq 0) {
        Write-Log "Troubleshooter completed successfully"
    } else {
        Write-Log "Troubleshooter failed with exit code $LASTEXITCODE"
    }
}
catch {
    Write-Log "FATAL ERROR: $_"
    [System.Windows.Forms.MessageBox]::Show("Error: $_", "DO Troubleshooter Error", "OK", "Error")
}
finally {
    # Clean up
    Write-Log "Cleaning up resources..."
    Remove-Item -Path $ExtractDir -Recurse -Force -ErrorAction SilentlyContinue
}
