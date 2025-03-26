[CmdletBinding()]
param (
    [switch]$Show,
    [string]$OutputPath = "$env:USERPROFILE\DOReports",
    [string]$DiagnosticsZip
)

# Initialize variables
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
        $resources = [Embedded]::GetNames()
        foreach ($resource in $resources) {
            $targetPath = Join-Path $ExtractDir $resource
            $targetDir = Split-Path $targetPath -Parent
            
            if (-not (Test-Path $targetDir)) {
                New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
            }
            
            [System.IO.File]::WriteAllBytes($targetPath, [Embedded]::Get($resource))
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
    
    # Use embedded PowerShell
    $localPwsh = Join-Path $ExtractDir "PowerShell-7.4.0-win-x64\pwsh.exe"
    if (Test-Path -Path $localPwsh) {
        Write-Log "Using embedded PowerShell 7: $localPwsh"
        return $localPwsh
    }
    
    throw "PowerShell 7 not found"
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
        $arguments = "-ExecutionPolicy Bypass -NoProfile -File `"$($MyInvocation.MyCommand.Definition)`" -OutputPath `"$OutputPath`""
        if ($Show) { $arguments += " -Show" }
        if ($DiagnosticsZip) { $arguments += " -DiagnosticsZip `"$DiagnosticsZip`"" }
        
        Start-Process -FilePath $pwshPath -ArgumentList $arguments -Verb RunAs
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
        Write-Log "Running troubleshooter as Admin..."
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
}
finally {
    # Clean up
    Write-Log "Cleaning up resources..."
    Remove-Item -Path $ExtractDir -Recurse -Force -ErrorAction SilentlyContinue
}
