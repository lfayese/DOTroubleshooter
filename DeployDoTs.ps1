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
.\InvokeDoTs.ps1 -Show
Launches the DO Troubleshooter with a message box and uses the default output path.
.EXAMPLE
.\InvokeDoTs.ps1 -OutputPath "C:\CustomReports"
Runs the DO Troubleshooter and saves the output to the specified custom directory.
.EXAMPLE
.\InvokeDoTs.ps1 -DiagnosticsZip "C:\Diagnostics\Logs.zip"
Runs the DO Troubleshooter with the specified diagnostics zip file.
.EXAMPLE
.\InvokeDoTs.ps1 -Show -OutputPath "C:\CustomReports" -DiagnosticsZip "C:\Diagnostics\Logs.zip"
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
# Determine the executable directory
$ExeDirectory = if ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} elseif ($env:EXEPATH) {
    Split-Path -Parent $env:EXEPATH
} else {
    $PSScriptRoot
}
# Setup temporary extraction directory and log file
$ExtractDir = Join-Path $env:TEMP "DODeploy_$([guid]::NewGuid().ToString())"
$LogFile = Join-Path $OutputPath "DODeployment.log"
# Create output directory
New-Item -ItemType Directory -Path $OutputPath -Force -ErrorAction SilentlyContinue | Out-Null
# Write-Log helper function to log messages with timestamps
function Write-Log {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Tee-Object -FilePath $LogFile -Append
}
# Checks if current process is running with administrator privileges
function Test-IsAdministrator {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltinRole]::Administrator)
}
# Expand-Resources: handles resource extraction from either embedded class or file copy mode.
function Expand-Resources {
    Write-Log "Extracting resources to $ExtractDir..."
    New-Item -Path $ExtractDir -ItemType Directory -Force | Out-Null
    $hasEmbedded = $false
    
    # Check for the Embedded class once and cache the result
    try {
        if ($null -ne [Embedded]::GetNames()) {
            $hasEmbedded = $true
        }
    } catch {
        $hasEmbedded = $false
    }
    
    if ($hasEmbedded) {
        Write-Log "Using embedded resources..."
        $resources = [Embedded]::GetNames()
        foreach ($resource in $resources) {
            $targetPath = Join-Path $ExtractDir $resource
            $targetDir  = Split-Path $targetPath -Parent
            if (-not (Test-Path $targetDir)) {
                New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
            }
            [System.IO.File]::WriteAllBytes($targetPath, [Embedded]::Get($resource))
        }
    } else {
        Write-Log "Embedded class not found. Using file copy mode..."
        $resourcesToExtract = @(
            "InvokeDoTs.ps1",
            "PowerShell-7.4.0-win-x64.zip",
            "PSTools"
        )
        foreach ($resource in $resourcesToExtract) {
            $sourcePath = Join-Path $ExeDirectory $resource
            $targetPath = Join-Path $ExtractDir $resource
            if (Test-Path -Path $sourcePath -PathType Container) {
                # Copy entire directory
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
        # Copy additional directories if they exist.
        foreach ($dir in @("Modules", "Scripts")) {
            $sourceDir = Join-Path $ExeDirectory $dir
            if (Test-Path $sourceDir) {
                Copy-Item -Path $sourceDir -Destination $ExtractDir -Recurse -Force
                Write-Log "Copied $dir directory"
            }
        }
    }
    Write-Log "Resource extraction complete"
}
# Get-PowerShell7: finds a valid PowerShell 7 executable (system-installed, embedded, or in common install locations)
function Get-PowerShell7 {
    # Try to detect a system-installed pwsh
    $systemPwsh = Get-Command -Name pwsh -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue
    if ($systemPwsh) {
        try {
            $ver = & $systemPwsh -NoProfile -Command { $PSVersionTable.PSVersion.Major }
            if ($ver -ge 7) {
                Write-Log "Using system PowerShell 7: $systemPwsh"
                return $systemPwsh
            }
        } catch { }
    }
    # Check for embedded versions
    $localPwsh74 = Join-Path $ExtractDir "PowerShell-7.4.0-win-x64\pwsh.exe"
    if (Test-Path -Path $localPwsh74) {
        Write-Log "Using embedded PowerShell 7.4.0: $localPwsh74"
        return $localPwsh74
    }
    # If compressed version exists, try to extract it
    $ps74zip = Join-Path $ExtractDir "PowerShell-7.4.0-win-x64.zip"
    if (Test-Path -Path $ps74zip) {
        $extractPath = Join-Path $ExtractDir "PowerShell-7.4.0-win-x64"
        Write-Log "Extracting PowerShell 7.4.0 from zip..."
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::ExtractToDirectory($ps74zip, $extractPath)
            if (Test-Path -Path $localPwsh74) {
                Write-Log "Using freshly extracted PowerShell 7.4.0: $localPwsh74"
                return $localPwsh74
            }
        } catch {
            Write-Log "Error extracting PowerShell 7.4.0: $_"
        }
    }
    # Lastly, check standard installation paths
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
# Main script execution block
try {
    # Load Windows Forms for message boxes
    Add-Type -AssemblyName System.Windows.Forms
    if ($Show) {
        [System.Windows.Forms.MessageBox]::Show("Launching DO Troubleshooter...", "DO Launcher", "OK", "Information")
    }
    # Extract resources (embedded or via file copy)
    Expand-Resources
    # Determine the path for PowerShell 7 executable
    $pwshPath = Get-PowerShell7
    # Elevate privileges if not running as administrator
    if (-not (Test-IsAdministrator)) {
        Write-Log "Elevating script as administrator..."
        $arguments = @()
        if ($env:EXEPATH) {
            # Running as an executable
            $elevationPath = $env:EXEPATH
            $arguments += "-OutputPath `"$OutputPath`""
            if ($Show)     { $arguments += "-Show" }
            if ($DiagnosticsZip) { $arguments += "-DiagnosticsZip `"$DiagnosticsZip`"" }
        } else {
            # Running as script: re-launch using pwsh
            $elevationPath = $pwshPath
            $scriptPath = $MyInvocation.MyCommand.Definition
            $arguments += "-ExecutionPolicy Bypass -NoProfile -File `"$scriptPath`" -OutputPath `"$OutputPath`""
            if ($Show)     { $arguments += "-Show" }
            if ($DiagnosticsZip) { $arguments += "-DiagnosticsZip `"$DiagnosticsZip`"" }
        }
        Start-Process -FilePath $elevationPath -ArgumentList $arguments -Verb RunAs
        exit
    }
    # Build command arguments to run the troubleshooter
    $scriptToRun = Join-Path $ExtractDir "Invoke-DoTroubleshooter.ps1"
    $cmdArgs = "& '$scriptToRun' -OutputPath '$OutputPath'"
    if ($Show)         { $cmdArgs += " -Show" }
    if ($DiagnosticsZip){ $cmdArgs += " -DiagnosticsZip '$DiagnosticsZip'" }
    # Execute via PsExec if available, else directly run with pwsh
    $psExecPath = Join-Path $ExtractDir "PSTools\PsExec64.exe"
    if (Test-Path $psExecPath) {
        Write-Log "Running troubleshooter via PsExec..."
        & $psExecPath -accepteula -s -i "$pwshPath" -ExecutionPolicy Bypass -NoProfile -Command $cmdArgs
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
finally
{
    Write-Log "Cleaning up resources..."
    Remove-Item -Path $ExtractDir -Recurse -Force -ErrorAction SilentlyContinue
}
