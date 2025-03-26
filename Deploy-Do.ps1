[CmdletBinding()]
param (
    [switch]$Show,
    [string]$OutputPath = "$env:USERPROFILE\DOReports"
)
Add-Type -AssemblyName System.Windows.Forms
if ($Show) {
    [System.Windows.Forms.MessageBox]::Show("Launching DO Troubleshooter...", "DO Launcher", "OK", "Information")
}
#region Configuration
$ExtractDir = Join-Path $env:TEMP "DODeploy_$([guid]::NewGuid().ToString())"
$LogFile    = Join-Path $OutputPath "DODeployment.log"
$PwshPath   = ""
#endregion
function Write-Log {
    param ([string]$Message)
    $timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    Add-Content -Path $LogFile -Value "$timestamp - $Message"
    Write-Host $Message
}
function Expand-EmbeddedResourcesParallel {
    Write-Log "Extracting embedded resources to $ExtractDir (parallel)..."
    New-Item -Path $ExtractDir -ItemType Directory -Force | Out-Null
    $resources = [Embedded]::GetNames()
    # Use ForEach-Object -Parallel with a throttle limit to limit concurrency
    $results = $resources | ForEach-Object -Parallel {
        # Using $using:ExtractDir to reference the variable from the parent scope
            try {
            $targetPath = Join-Path $using:ExtractDir $_
                $targetDir = Split-Path $targetPath
                if (-not (Test-Path $targetDir)) {
                    New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
                }
            $bytes = [Embedded]::Get($_)
                [System.IO.File]::WriteAllBytes($targetPath, $bytes)
            return "Extracted: $_"
            }
            catch {
            return "FAILED: $_ - $($_.Exception.Message)"
            }
    } -ThrottleLimit 4   # Adjust throttle limit as appropriate
    # Log each extraction result
    $results | ForEach-Object { Write-Log $_ }
    }
function Test-IsAdministrator {
    $identity  = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltinRole]::Administrator)
}
function Get-PwshExecutable {
    try {
        $pwsh = Get-Command -Name pwsh -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
        if ($pwsh) {
            $version = & $pwsh -NoProfile -Command { $PSVersionTable.PSVersion.Major }
            if ($version -ge 7) {
                Write-Log "Using system-installed PowerShell 7: $pwsh"
                return $pwsh
            }
        }
        $localPwsh = Join-Path $ExtractDir "PowerShell-7.4.0-win-x64\pwsh.exe"
        if (Test-Path -Path $localPwsh) {
            Write-Log "Using embedded PowerShell 7: $localPwsh"
            return $localPwsh
        }
        throw "PowerShell 7 not found."
    }
    catch {
        Write-Log "ERROR: $_"
        exit 1
    }
}
function Invoke-Elevation {
    if (-not (Test-IsAdministrator)) {
        Write-Log "Elevating script as administrator..."
        $pwshExe = Get-PwshExecutable
        Start-Process -FilePath $pwshExe -ArgumentList "-ExecutionPolicy Bypass -NoProfile -File `"$($MyInvocation.MyCommand.Definition)`" -Show:$Show -OutputPath:`"$OutputPath`"" -Verb RunAs
        exit
    }
}
function Invoke-Troubleshooter {
    try {
        $scriptPath = Join-Path $ExtractDir "Invoke-DoTroubleshooter.ps1"
        $psExec     = Join-Path $ExtractDir "PSTools\PsExec64.exe"
        $cmdArgs    = "& '$scriptPath' -OutputPath '$OutputPath' -Show:$Show"
        if (Test-Path $psExec) {
            Write-Log "Running troubleshooter via PsExec..."
            & $psExec -accepteula -s -i "$PwshPath" -ExecutionPolicy Bypass -NoProfile -Command $cmdArgs
        }
        else {
            Write-Log "Running troubleshooter as Admin..."
            & "$PwshPath" -ExecutionPolicy Bypass -NoProfile -Command $cmdArgs
        }
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Troubleshooter completed successfully."
        } else {
            Write-Log "Troubleshooter failed with exit code $LASTEXITCODE"
        }
    }
    catch {
        Write-Log "ERROR during troubleshooter run: $_"
    }
}
function Remove-ExtractedResources {
    Write-Log "Cleaning up extracted files..."
    Remove-Item -Path $ExtractDir -Recurse -Force -ErrorAction SilentlyContinue
}
# --- Main Execution ---
try {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Expand-EmbeddedResourcesParallel
    $PwshPath = Get-PwshExecutable
    Invoke-Elevation
    Invoke-Troubleshooter
}
catch {
    Write-Log "FATAL ERROR: $_"
}
finally {
    Remove-ExtractedResources
}
