<#
.SYNOPSIS
  Generates a comprehensive Delivery Optimization troubleshooting report in an Excel workbook.

.DESCRIPTION
  This script analyzes Delivery Optimization (DO) on the local system and creates an Excel report.
  It examines:
    • DO service health and status
    • Peer-to-peer connectivity on ports 7680 and 3544
    • Microsoft Connected Cache configuration
    • DNS-SD peer discovery settings
    • Group ID configuration
    • Critical endpoint connectivity
    • Non-peerable Teams content
    • Detailed connection statistics

  Results are compiled into an Excel workbook with recommendations for identified issues.

.PARAMETER OutputPath
  Directory for saving the report. Defaults to user's Desktop.

.PARAMETER Show
  Opens the report automatically when completed.

.PARAMETER DiagnosticsZip
  Optional path to a diagnostics ZIP file containing DO logs for analysis.

.EXAMPLE
  .\Invoke-DoTroubleshooter.ps1 -OutputPath "C:\DOReports" -Show

.EXAMPLE
  .\Invoke-DoTroubleshooter.ps1 -DiagnosticsZip "C:\Path\To\DiagnosticsFile.zip" -Show

.NOTES
  Requires PowerShell 5+ and administrative privileges for full functionality.
  Uses the ImportExcel module (bundled if not installed).

.LINK
  https://learn.microsoft.com/en-us/windows/deployment/optimization/delivery-optimization
#>

param (
    [string]$OutputPath = [Environment]::GetFolderPath("Desktop"),
    [switch]$Show,
    [string]$DiagnosticsZip
)

# Setup ImportExcel module
$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\ImportExcel"
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "[INFO] ImportExcel not found, importing from local module path..." -ForegroundColor Yellow
    Import-Module -Name $modulePath -Force
} else {
    Import-Module ImportExcel
}

# Initialize data collection buffers
$Buffers = @{
    Log             = [System.Collections.ArrayList]::new()
    HealthCheck     = [System.Collections.ArrayList]::new()
    P2P             = [System.Collections.ArrayList]::new()
    MCC             = [System.Collections.ArrayList]::new()
    SysInfo         = [System.Collections.ArrayList]::new()
    Errors          = [System.Collections.ArrayList]::new()
    Stats           = [System.Collections.ArrayList]::new()
    DOConfig        = [System.Collections.ArrayList]::new()
    Peers           = [System.Collections.ArrayList]::new()
    Summary         = [System.Collections.ArrayList]::new()
    Recommendations = [System.Collections.ArrayList]::new()
    DiagnosticsData = [System.Collections.ArrayList]::new()
}

# ───── HELPER FUNCTIONS ─────

function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]$Level = "INFO",
        [string]$Category = "Log"
    )
    $entry = [PSCustomObject]@{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Level     = $Level
        Message   = $Message
    }
    $Buffers[$Category].Add($entry) | Out-Null

    switch ($Level) {
        "ERROR" {
            $Buffers.Errors.Add($entry) | Out-Null
            Write-Host "[ERROR] $Message" -ForegroundColor Red
        }
        "WARN" {
            Write-Host "[WARN]  $Message" -ForegroundColor Yellow
        }
        "SUCCESS" {
            Write-Host "[SUCCESS] $Message" -ForegroundColor Green
        }
        default {
            Write-Host "[INFO]  $Message" -ForegroundColor Cyan
        }
    }
}

function Add-Recommendation {
    param (
        [string]$Area,
        [string]$Recommendation,
        [ValidateSet("Critical", "Important", "Informational")]$Severity = "Informational",
        [string]$Reference = ""
    )

    $entry = [PSCustomObject]@{
        Area           = $Area
        Recommendation = $Recommendation
        Severity       = $Severity
        Reference      = $Reference
    }
    $Buffers.Recommendations.Add($entry) | Out-Null
}

function Show-Progress {
    param (
        [string]$Activity,
        [int]$PercentComplete
    )
    Write-Progress -Activity "DO Troubleshooter" -Status $Activity -PercentComplete $PercentComplete
}

function Test-DiagnosticsZipPath {
    param ([string]$Path)
    
    if ([string]::IsNullOrEmpty($Path) -or $Path -match '[<>:"|?*]' -or !(Test-Path -Path $Path -PathType Leaf)) {
        Write-Log "Invalid diagnostics zip path: $Path" "ERROR"
        return $false
    }
    
    if (-not ($Path -match '\.(zip|cab)$')) {
        Write-Log "File is not a valid zip or cabinet file: $Path" "ERROR"
        return $false
    }
    
    return $true
}

function Remove-TemporaryFiles {
    param ([string]$Path)
    
    if (Test-Path -Path $Path) {
        try {
            Remove-Item -Path $Path -Recurse -Force
            Write-Log "Cleaned up temporary files at $Path" "INFO"
        } catch {
            Write-Log "Failed to clean up temporary files: $_" "WARN"
        }
    }
}

function Get-DOErrorsTable {
    $errorsObj = @'
[
  { "ErrorCode": "0x80D01001", "Description": "Delivery Optimization was unable to provide the service.", "Recommendation": "Check if the DO service is running and properly configured." },
  { "ErrorCode": "0x80D02002", "Description": "Download of a file saw no progress within the defined period.", "Recommendation": "Check network connectivity and available bandwidth." },
  { "ErrorCode": "0x80D02003", "Description": "Job was not found.", "Recommendation": "The download job may have been cancelled or timed out." },
  { "ErrorCode": "0x80D02005", "Description": "There were no files in the job.", "Recommendation": "Verify that the content is available for download." },
  { "ErrorCode": "0x80D0200B", "Description": "Memory stream transfer is not supported.", "Recommendation": "Use file-based transfers instead of memory streams." },
  { "ErrorCode": "0x80D0200C", "Description": "Job has neither completed nor has it been cancelled prior to reaching the max age threshold.", "Recommendation": "Wait for the job to complete or manually cancel it." },
  { "ErrorCode": "0x80D0200D", "Description": "There is no local file path specified for this download.", "Recommendation": "Provide a valid local file path for the download." },
  { "ErrorCode": "0x80D02010", "Description": "No file is available because no URL generated an error.", "Recommendation": "Verify that the download URLs are valid and accessible." },
  { "ErrorCode": "0x80D02011", "Description": "SetProperty() or GetProperty() called with an unknown property ID.", "Recommendation": "Use valid property IDs with the DO API." },
  { "ErrorCode": "0x80D02012", "Description": "Unable to call SetProperty() on a read-only property.", "Recommendation": "Only attempt to modify writable properties." },
  { "ErrorCode": "0x80D02013", "Description": "The requested action is not allowed in the current job state.", "Recommendation": "Check the job state before performing the action." },
  { "ErrorCode": "0x80D02015", "Description": "Unable to call GetProperty() on a write-only property.", "Recommendation": "Only attempt to read readable properties." },
  { "ErrorCode": "0x80D02016", "Description": "Download job is marked as requiring integrity checking but integrity checking info was not specified.", "Recommendation": "Provide integrity checking information when required." },
  { "ErrorCode": "0x80D02017", "Description": "Download job is marked as requiring integrity checking but integrity checking info could not be retrieved.", "Recommendation": "Verify integrity checking information is accessible." },
  { "ErrorCode": "0x80D02018", "Description": "Unable to start a download because no download sink (either local file or stream interface) was specified.", "Recommendation": "Specify a download destination before starting the download." }
]
'@ | ConvertFrom-Json

    return $errorsObj | ForEach-Object {
        [PSCustomObject]@{
            ErrorCode = $_.ErrorCode
            Description = $_.Description
            Recommendation = $_.Recommendation
        }
    }
}

# ───── SYSTEM INFORMATION AND NETWORK TESTS ─────

function Get-SystemInfo {
    Write-Log "Collecting system information..." "INFO"
    Show-Progress -Activity "Collecting system information" -PercentComplete 10
    
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem
        $nic = Get-NetAdapter | Where-Object Status -eq "Up" | Select-Object -First 1
        
        $sysInfo = [PSCustomObject]@{
            ComputerName    = $env:COMPUTERNAME
            OSVersion       = $os.Caption
            OSBuild         = $os.BuildNumber
            Manufacturer    = $cs.Manufacturer
            Model           = $cs.Model
            Memory          = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
            NetworkAdapter  = $nic.Name
            IPAddress       = (Get-NetIPAddress -InterfaceIndex $nic.ifIndex -AddressFamily IPv4).IPAddress
            UserName        = "$env:USERDOMAIN\$env:USERNAME"
            CollectionTime  = Get-Date
        }
        
        $Buffers.Summary.Add($sysInfo) | Out-Null
        Write-Log "System information collected successfully" "SUCCESS"
    }
    catch {
        Write-Log "Error collecting system information: $_" "ERROR"
    }
}

# ───── DO CONFIGURATION TESTS ─────

function Test-DNSDPeerConfig {
    Write-Log "Checking DNS-SD DO Peer Discovery setting..." "INFO"
    Show-Progress -Activity "Checking DNS-SD configuration" -PercentComplete 30
    
    $key = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization"
    $name = "DORestrictPeerSelectionBy"
    
    try {
        $val = Get-ItemProperty -Path $key -Name $name -ErrorAction Stop | Select-Object -ExpandProperty $name
        
        if ($val -eq 2) {
            Write-Log "DNS-SD enabled for peer discovery (value=2)." "SUCCESS"
            $config = [PSCustomObject]@{
                Setting     = "DORestrictPeerSelectionBy"
                Value       = $val
                Status      = "Optimal"
                Description = "DNS-SD peer discovery enabled"
                Impact      = "Improved peer discovery across subnets"
            }
            
            $Buffers.Summary.Add([PSCustomObject]@{
                Test = "DNS-SD Configuration"
                Result = "Enabled (value=2)"
                Status = "PASS"
                Impact = "Optimal peer discovery"
            }) | Out-Null
        } else {
            Write-Log "DORestrictPeerSelectionBy not set to 2. DNS-SD likely disabled (value=$val)." "WARN"
            $config = [PSCustomObject]@{
                Setting     = "DORestrictPeerSelectionBy"
                Value       = $val
                Status      = "Sub-optimal"
                Description = "DNS-SD peer discovery not enabled"
                Impact      = "Limited peer discovery across subnets"
            }
            
            $Buffers.Summary.Add([PSCustomObject]@{
                Test = "DNS-SD Configuration"
                Result = "Not enabled (value=$val)"
                Status = "WARN"
                Impact = "Reduced peer discovery capabilities"
            }) | Out-Null
            
            Add-Recommendation -Area "Configuration" -Recommendation "Set DORestrictPeerSelectionBy registry value to 2 for optimal peer discovery" -Severity "Important" -Reference "https://learn.microsoft.com/en-us/windows/deployment/optimization/delivery-optimization-reference#delivery-optimization-options"
        }
        
        $Buffers.DOConfig.Add($config) | Out-Null
    } catch {
        Write-Log "DNS-SD config missing or inaccessible." "WARN"
        $config = [PSCustomObject]@{
            Setting     = "DORestrictPeerSelectionBy"
            Value       = "N/A"
            Status      = "Missing"
            Description = "DNS-SD peer discovery setting not found"
            Impact      = "Potential peer discovery limitations"
        }
        
        $Buffers.Summary.Add([PSCustomObject]@{
            Test = "DNS-SD Configuration"
            Result = "Setting not found"
            Status = "WARN"
            Impact = "Potential peer discovery limitations"
        }) | Out-Null
        
        Add-Recommendation -Area "Configuration" -Recommendation "Configure DNS-SD peer discovery by setting DORestrictPeerSelectionBy registry value to 2" -Severity "Important" -Reference "https://learn.microsoft.com/en-us/windows/deployment/optimization/delivery-optimization-reference#delivery-optimization-options"
    }
}

function Test-DOEndpoints {
    Write-Log "Testing critical DO endpoints..." "INFO"
    Show-Progress -Activity "Testing DO endpoints" -PercentComplete 20
    
    $doEndpoints = @(
        [PSCustomObject]@{
            Url = "http://delivery.mp.microsoft.com"
            Description = "Primary delivery content endpoint for Windows Update"
            Required = $true
        },
        [PSCustomObject]@{
            Url = "http://emdl.ws.microsoft.com"
            Description = "Electronic Software Delivery endpoint for Windows Store"
            Required = $true
        },
        [PSCustomObject]@{
            Url = "http://download.windowsupdate.com"
            Description = "Windows Update download endpoint"
            Required = $true
        },
        [PSCustomObject]@{
            Url = "https://tsfe.trafficshaping.dsp.mp.microsoft.com"
            Description = "Traffic shaping endpoint for Delivery Optimization"
            Required = $false
        }
    )
    
    # Start parallel jobs for each endpoint
    $jobs = $doEndpoints | ForEach-Object {
        $endpoint = $_
        Start-Job -ScriptBlock {
            param($url, $description, $required)
            
            try {
                $request = Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing -TimeoutSec 10
                
                return [PSCustomObject]@{
                    Url = $url
                    Description = $description
                    Required = $required
                    Success = $true
                    Status = $request.StatusCode
                    Error = $null
                }
            } catch {
                return [PSCustomObject]@{
                    Url = $url
                    Description = $description
                    Required = $required
                    Success = $false
                    Status = 0
                    Error = $_.Exception.Message
                }
            }
        } -ArgumentList $endpoint.Url, $endpoint.Description, $endpoint.Required
    }
    
    # Process results
    $results = @()
    foreach ($job in $jobs) {
        $result = Receive-Job -Job $job -Wait
        $results += $result
        
        if ($result.Success) {
            Write-Log "Success: $($result.Url) reachable" "SUCCESS"
        } else {
            $severity = if ($result.Required) { "ERROR" } else { "WARN" }
            Write-Log "Failed to reach $($result.Url): $($result.Error)" $severity
            
            if ($result.Required) {
                Add-Recommendation -Area "Network" -Recommendation "Ensure connectivity to $($result.Url) - $($result.Description)" -Severity "Critical" -Reference "https://learn.microsoft.com/en-us/windows/deployment/optimization/delivery-optimization-reference#requirements"
            }
        }
        
        Remove-Job -Job $job
    }
    
    # Add results to buffer
    $results | ForEach-Object { $Buffers.DOConfig.Add($_) | Out-Null }
    
    # Add summary
    $successCount = ($results | Where-Object { $_.Success }).Count
    $totalCount = $results.Count
    $requiredFailures = ($results | Where-Object { -not $_.Success -and $_.Required }).Count
    
    $Buffers.Summary.Add([PSCustomObject]@{
        Test = "DO Endpoints"
        Result = "$successCount of $totalCount endpoints reachable"
        Status = if ($requiredFailures -eq 0) { "PASS" } else { "FAIL" }
        Impact = if ($requiredFailures -eq 0) { "None" } else { "Critical - DO functionality will be impaired" }
    }) | Out-Null
    
    return $results
}

function Test-DOGroupId {
    Write-Log "Checking DO Group ID configuration..." "INFO"
    Show-Progress -Activity "Checking DO Group ID" -PercentComplete 40
    
    try {
        $groupId = ([Windows.Management.Policies.NamedPolicy]::GetPolicyFromPath("DeliveryOptimization", "DOGroupId")).GetString()
        
        if ([string]::IsNullOrEmpty($groupId)) {
            Write-Log "DOGroupId not defined in policy. If using DHCP Option 234, logs may show NULL." "WARN"
            $config = [PSCustomObject]@{
                Setting     = "DOGroupId"
                Value       = "Not defined in policy"
                Status      = "Warning"
                Description = "No Group ID found in policy"
                Impact      = "May need DHCP Option 234 for cross-subnet peers"
            }
            
            $Buffers.Summary.Add([PSCustomObject]@{
                Test = "DO Group ID"
                Result = "Not defined in policy"
                Status = "WARN"
                Impact = "May limit cross-subnet peer capabilities"
            }) | Out-Null
            
            Add-Recommendation -Area "Configuration" -Recommendation "Configure DOGroupId via policy or DHCP Option 234 for optimal cross-subnet peer sharing" -Severity "Important" -Reference "https://learn.microsoft.com/en-us/windows/deployment/optimization/delivery-optimization-reference#delivery-optimization-options"
        } else {
            Write-Log "DOGroupId found: $groupId" "SUCCESS"
            $config = [PSCustomObject]@{
                Setting     = "DOGroupId"
                Value       = $groupId
                Status      = "Configured"
                Description = "Group ID defined via policy"
                Impact      = "Enables cross-subnet peer sharing"
            }
            
            $Buffers.Summary.Add([PSCustomObject]@{
                Test = "DO Group ID"
                Result = "Configured: $groupId"
                Status = "PASS"
                Impact = "Optimal cross-subnet peer sharing"
            }) | Out-Null
        }
        
        $Buffers.DOConfig.Add($config) | Out-Null
    } catch {
        Write-Log "NamedPolicy query for DOGroupId failed (likely unsupported OS)." "WARN"
        $config = [PSCustomObject]@{
            Setting     = "DOGroupId"
            Value       = "Unknown"
            Status      = "Error"
            Description = "Unable to detect Group ID on this OS version"
            Impact      = "Cannot determine impact on peer sharing"
        }
        
        $Buffers.Summary.Add([PSCustomObject]@{
            Test = "DO Group ID"
            Result = "Detection failed"
            Status = "WARN"
            Impact = "Cannot determine impact"
        }) | Out-Null
    }
}

# ───── DO SERVICE HEALTH AND ANALYSIS ─────

function Start-HealthCheck {
    Write-Log "Running DO Health Check..." "INFO" "HealthCheck"
    Show-Progress -Activity "Running DO health check" -PercentComplete 50
    
    try {
        Enable-DeliveryOptimizationVerboseLogs | Out-Null
        $status = Get-DeliveryOptimizationStatus -ErrorAction Stop
        
        # Add context information about download modes
        $modeDesc = switch ($status.DODownloadMode) {
            0 {"HTTP Only (No Peering) - Uses HTTP without peer-to-peer"}
            1 {"Local Network Peering Only - Peers with devices in the same subnet"}
            2 {"Local Network + Internet Peering - Peers with local subnet and internet devices"}
            3 {"Local Network + Internet + Group - Full peering across subnets with same Group ID"}
            99 {"Fallback Mode - Service Issues - DO service is experiencing problems"}
            default {"Unknown ($($status.DODownloadMode)) - Unrecognized download mode"}
        }
        
        # Create enhanced status object
        $enrichedStatus = [PSCustomObject]@{
            DODownloadMode = $status.DODownloadMode
            DODownloadModeDescription = $modeDesc
            NumberOfPeers = $status.NumberOfPeers
            MinFileSizeToCache = $status.MinFileSizeToCache
            MaxCacheSize = $status.MaxCacheSize
            CurrentCacheSize = $status.CurrentCacheSize
            CacheSizePercentage = if ($status.MaxCacheSize -gt 0) {
                [math]::Round(($status.CurrentCacheSize / $status.MaxCacheSize) * 100, 2)
            } else { 0 }
            PeerCachingAllowed = $status.PeerCachingAllowed
            Status = if ($status.DODownloadMode -eq 99) { "CRITICAL" } elseif ($status.DODownloadMode -eq 0) { "WARNING" } else { "GOOD" }
        }
        
        $Buffers.HealthCheck.Add($enrichedStatus) | Out-Null
        $Buffers.Stats.Add([PSCustomObject]@{
            Category       = "Health"
            DODownloadMode = $status.DODownloadMode
            NumberOfPeers  = $status.NumberOfPeers
        }) | Out-Null
        
        Write-Log "DODownloadMode: $($status.DODownloadMode) - $modeDesc" "INFO" "HealthCheck"
        
        # Handle different modes
        if ($status.DODownloadMode -eq 99) {
            Write-Log "DO is in fallback mode (99) - Service is experiencing issues!" "ERROR" "HealthCheck"
            Add-Recommendation -Area "Service Health" -Recommendation "DO is in fallback mode. Restart the Delivery Optimization service and check system health." -Severity "Critical"
        } elseif ($status.DODownloadMode -eq 0) {
            Write-Log "DO is in HTTP Only mode (0) - No peer-to-peer functionality is enabled" "WARN" "HealthCheck"
            Add-Recommendation -Area "Configuration" -Recommendation "Consider enabling peer caching by changing DO download mode from HTTP Only (0) to a higher level." -Severity "Important"
        } else {
            Write-Log "DO download mode is set to $($status.DODownloadMode) with $($status.NumberOfPeers) peers detected" "SUCCESS" "HealthCheck"
        }
        
        $Buffers.Summary.Add([PSCustomObject]@{
            Test = "DO Service Health"
            Result = "Mode: $($status.DODownloadMode) - $($modeDesc.Split(' - ')[0])"
            Status = if ($status.DODownloadMode -eq 99) { "CRITICAL" } elseif ($status.DODownloadMode -eq 0) { "WARN" } else { "PASS" }
            Impact = if ($status.DODownloadMode -eq 99) { "Service not functioning properly" } elseif ($status.DODownloadMode -eq 0) { "No bandwidth savings from peer caching" } else { "Optimal configuration" }
        }) | Out-Null
    } catch {
        Write-Log "HealthCheck failed: $_" "ERROR" "HealthCheck"
        $Buffers.Summary.Add([PSCustomObject]@{
            Test = "DO Service Health"
            Result = "Check failed"
            Status = "ERROR"
            Impact = "Unable to determine service health"
        }) | Out-Null
        
        Add-Recommendation -Area "Service Health" -Recommendation "Verify that the Delivery Optimization service is running and that you have sufficient permissions." -Severity "Critical"
    }
}

function Test-PeerConnectivity {
    Write-Log "Testing P2P ports (7680, 3544)..." "INFO" "P2P"
    Show-Progress -Activity "Testing peer connectivity" -PercentComplete 60
    
    try {
        $ipConfig = Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -and $_.NetAdapter.Status -eq "Up" }
        
        if ($ipConfig) {
            $gateway = $ipConfig[0].IPv4DefaultGateway.NextHop
            $subnet = $gateway -replace '\d+$', "1"
        } else {
            $subnet = "192.168.1.1"
        }
    } catch {
        $subnet = "192.168.1.1"
    }
    
    $peers = @($subnet, "192.168.0.1", $env:COMPUTERNAME)
    $ports = @(
        [PSCustomObject]@{Port = 7680; Description = "Primary DO peer-to-peer port"},
        [PSCustomObject]@{Port = 3544; Description = "Secondary DO peer-to-peer port"}
    )
    
    $jobs = @()
    foreach ($peer in $peers) {
        foreach ($portInfo in $ports) {
            $jobs += Start-Job -ScriptBlock {
                param($peerAddr, $portNum, $portDesc)
                
                $tcpClient = New-Object System.Net.Sockets.TcpClient
                try {
                    $result = $tcpClient.BeginConnect($peerAddr, $portNum, $null, $null)
                    $success = $result.AsyncWaitHandle.WaitOne(1000, $true)
                    
                    return [PSCustomObject]@{
                        Peer = $peerAddr
                        Port = $portNum
                        Description = $portDesc
                        TcpTestSucceeded = $success
                    }
                } catch {
                    return [PSCustomObject]@{
                        Peer = $peerAddr
                        Port = $portNum
                        Description = $portDesc
                        TcpTestSucceeded = $false
                    }
                } finally {
                    $tcpClient.Close()
                }
            } -ArgumentList $peer, $portInfo.Port, $portInfo.Description
        }
    }
    
    # Process results
    foreach ($job in $jobs) {
        $result = Receive-Job -Job $job -Wait
        $status = if ($result.TcpTestSucceeded) { "SUCCESS" } else { "FAILED" }
        $level = if ($result.TcpTestSucceeded) { "SUCCESS" } else { "WARN" }
        
        Write-Log "Peer $($result.Peer) | Port $($result.Port) | TCP: $status" $level "P2P"
        $Buffers.Peers.Add($result) | Out-Null
        
        Remove-Job -Job $job
    }
    
    # Add summary
    $portResults = @{
        7680 = [PSCustomObject]@{Success = 0; Failure = 0}
        3544 = [PSCustomObject]@{Success = 0; Failure = 0}
    }
    
    foreach ($result in $Buffers.Peers) {
        if ($result.TcpTestSucceeded) {
            $portResults[$result.Port].Success++
        } else {
            $portResults[$result.Port].Failure++
        }
    }
    
    foreach ($port in $ports.Port) {
        $total = $portResults[$port].Success + $portResults[$port].Failure
        $successRate = if ($total -gt 0) { [math]::Round(($portResults[$port].Success / $total) * 100, 0) } else { 0 }
        
        $Buffers.Summary.Add([PSCustomObject]@{
            Test = "Port $port Connectivity"
            Result = "$($portResults[$port].Success) of $total connections successful ($successRate%)"
            Status = if ($successRate -gt 50) { "PASS" } elseif ($successRate -gt 0) { "WARN" } else { "FAIL" }
            Impact = if ($successRate -eq 100) { "None" } elseif ($successRate -gt 50) { "Minor impact" } else { "Significant impact on peer-to-peer" }
        }) | Out-Null
        
        if ($successRate -lt 50) {
            Add-Recommendation -Area "Network" -Recommendation "Check firewall rules for port $port to ensure peer-to-peer connectivity" -Severity "Important" -Reference "https://learn.microsoft.com/en-us/windows/deployment/optimization/delivery-optimization-reference#ports"
        }
    }
}

function Test-MCCCheck {
    Write-Log "Checking Microsoft Connected Cache configuration..." "INFO" "MCC"
    Show-Progress -Activity "Checking Microsoft Connected Cache" -PercentComplete 70
    
    try {
        $diag = Get-DOConfig -Verbose
        
        # Create structured MCC configuration object
        $mccConfig = [PSCustomObject]@{
            Category = "MCC"
            Config   = ($diag | Out-String).Trim()
            ConnectedCacheServers = if ($diag.PSObject.Properties.Name -contains "ConnectedCacheServers") { $diag.ConnectedCacheServers } else { "None configured" }
            Status = "Configured"
            Description = "Microsoft Connected Cache allows on-premises caching of content"
            Impact = if (($diag.PSObject.Properties.Name -contains "ConnectedCacheServers") -and $diag.ConnectedCacheServers) {
                "Reduces internet bandwidth by serving content from local cache servers"
            } else {
                "Not utilizing on-premises caching benefits"
            }
        }
        
        $Buffers.MCC.Add($mccConfig) | Out-Null
        $Buffers.Stats.Add([PSCustomObject]@{
            Category = "MCC"
            Config   = ($diag | Out-String).Trim()
        }) | Out-Null
        
        if ($diag.PSObject.Properties.Name -contains "ConnectedCacheServers") {
            if ($diag.ConnectedCacheServers) {
                Write-Log "Connected Cache configured: $($diag.ConnectedCacheServers)" "SUCCESS" "MCC"
                $Buffers.Summary.Add([PSCustomObject]@{
                    Test = "Microsoft Connected Cache"
                    Result = "Configured"
                    Status = "PASS"
                    Impact = "Optimized for on-premises caching"
                }) | Out-Null
            } else {
                Write-Log "No Connected Cache servers defined" "WARN" "MCC"
                $Buffers.Summary.Add([PSCustomObject]@{
                    Test = "Microsoft Connected Cache"
                    Result = "Not configured"
                    Status = "WARN"
                    Impact = "Missing potential bandwidth optimization"
                }) | Out-Null
                
                Add-Recommendation -Area "Infrastructure" -Recommendation "Consider setting up Microsoft Connected Cache servers for bandwidth optimization" -Severity "Informational" -Reference "https://learn.microsoft.com/en-us/windows/deployment/optimization/delivery-optimization
-reference#connected-cache-server"
            }
        } else {
            Write-Log "Connected Cache configuration not found in DO config" "WARN" "MCC"
            $Buffers.Summary.Add([PSCustomObject]@{
                Test = "Microsoft Connected Cache"
                Result = "Not detected"
                Status = "WARN"
                Impact = "Missing potential bandwidth optimization"
            }) | Out-Null
            
            Add-Recommendation -Area "Infrastructure" -Recommendation "Consider setting up Microsoft Connected Cache servers for bandwidth optimization" -Severity "Informational" -Reference "https://learn.microsoft.com/en-us/windows/deployment/optimization/delivery-optimization-reference#connected-cache-server"
        }
    } catch {
        Write-Log "Failed to retrieve MCC config: $_" "ERROR" "MCC"
        $Buffers.Summary.Add([PSCustomObject]@{
            Test = "Microsoft Connected Cache"
            Result = "Check failed"
            Status = "ERROR"
            Impact = "Unable to determine MCC configuration"
        }) | Out-Null
    }
}

function Invoke-DOLogAnalysis {
    Write-Log "Running DO log analysis..." "INFO"
    Show-Progress -Activity "Analyzing DO logs" -PercentComplete 80

    try {
        # Process results
        $results = Get-DeliveryOptimizationLogAnalysis -ListConnections

        # Categorize and process results
        $categorizedResults = @()
        foreach ($result in $results) {
            # Extract basic stats
            $successCategory = "Unknown"

            if ($result.Result -eq "Success") {
                # Check if IP is in private range
                if ($result.Url -match "http.*://(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
                    $ip = $matches[1]
                    if ($ip -match "^(10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[0-1])\.)") {
                        $successCategory = "Local Peer"
                    } else {
                        $successCategory = "Internet Peer"
                    }
                } elseif ($result.Url -match "http.*://(.*?)\.") {
                    $domain = $matches[1]
                    if ($domain -match "windowsupdate|microsoft") {
                        $successCategory = "Microsoft Server"
                    } else {
                        $successCategory = "Internet Peer"
                    }
                }
            }

            $categorizedResults += [PSCustomObject]@{
                Timestamp = $result.TimeStamp
                Url = $result.Url
                Result = $result.Result
                Speed = $result.Speed
                Duration = $result.Duration
                Bytes = $result.Bytes
                SuccessCategory = if ($result.Result -eq "Success") { $successCategory } else { "N/A" }
                ErrorCategory = if ($result.Result -ne "Success") { "Connection Failed" } else { "N/A" }
            }
        }

        # Calculate statistics
        $total = $results.Count
        $success = ($results | Where-Object { $_.Result -eq "Success" }).Count
        $failed = $total - $success
        $successRate = if ($total -gt 0) { [math]::Round(($success / $total) * 100, 0) } else { 0 }

        Write-Log "Log analysis complete: $success successful ($successRate%), $failed failed connections" (if($successRate -gt 75) {"SUCCESS"} elseif($successRate -gt 25) {"INFO"} else {"WARN"})

        # Additional statistics
        $localPeers = ($categorizedResults | Where-Object { $_.SuccessCategory -eq "Local Peer" }).Count
        $internetPeers = ($categorizedResults | Where-Object { $_.SuccessCategory -eq "Internet Peer" }).Count
        $totalBytes = ($categorizedResults | Measure-Object -Property Bytes -Sum).Sum

        # Add summary information
        $Buffers.Summary.Add([PSCustomObject]@{
            Test = "DO Connection Success Rate"
            Result = "$success successful of $total total connections ($successRate%)"
            Status = if ($successRate -gt 75) { "PASS" } elseif ($successRate -gt 25) { "WARN" } else { "FAIL" }
            Impact = if ($successRate -gt 75) { "Effective peer sharing" } elseif ($successRate -gt 25) { "Partial peer sharing benefits" } else { "Limited peer sharing benefits" }
        }) | Out-Null

        $Buffers.Summary.Add([PSCustomObject]@{
            Test = "Peer Source Distribution"
            Result = "Local: $localPeers, Internet: $internetPeers"
            Status = "INFO"
            Impact = "Data on peer source distribution"
        }) | Out-Null

        if ($successRate -lt 25) {
            Add-Recommendation -Area "Network" -Recommendation "Low peer connection success rate ($successRate%). Check network connectivity and firewall rules." -Severity "Important"
        }

        $categorizedResults | ForEach-Object { $Buffers.SysInfo.Add($_) | Out-Null }

        return $results
    } catch {
        Write-Log "Log analysis failed: $_" "ERROR"
        $Buffers.Summary.Add([PSCustomObject]@{
            Test = "DO Log Analysis"
            Result = "Analysis failed"
            Status = "ERROR"
            Impact = "Unable to determine peer connection effectiveness"
        }) | Out-Null

        Add-Recommendation -Area "Diagnostics" -Recommendation "Run the analysis with administrative privileges to access DO logs." -Severity "Informational"
    }
}

function Test-TeamsImpact {
    Write-Log "Analyzing DO log for Teams content..." "INFO"
    Show-Progress -Activity "Analyzing Teams content in DO logs" -PercentComplete 75

    try {
        $teamLogs = Get-DeliveryOptimizationLog | Where-Object { $_.Message -match "Teams" }
        if ($teamLogs) {
            $count = $teamLogs.Count
            $enrichedLogs = $teamLogs | ForEach-Object {
                [PSCustomObject]@{
                    Timestamp = $_.Timestamp
                    Message = $_.Message
                    Impact = "Teams content cannot use peer caching (expected behavior)"
                    Recommendation = "No action needed - Microsoft Teams has its own delivery mechanism"
                }
            }
            $enrichedLogs | ForEach-Object { $Buffers.SysInfo.Add($_) | Out-Null }
            Write-Log "Found $count Teams-related DO log entries (non-peerable traffic)" "WARN"
            $Buffers.Summary.Add([PSCustomObject]@{
                Test = "Teams Content Analysis"
                Result = "$count Teams content entries found"
                Status = "INFO"
                Impact = "Teams content cannot use peer caching (expected behavior)"
            }) | Out-Null
            Add-Recommendation -Area "Content" -Recommendation "Teams content cannot use peer caching. This is expected behavior and not an issue." -Severity "Informational"
        } else {
            Write-Log "No Teams DO content found in logs" "INFO"
            $Buffers.Summary.Add([PSCustomObject]@{
                Test = "Teams Content Analysis"
                Result = "No Teams content found"
                Status = "PASS"
                Impact = "None"
            }) | Out-Null
        }
    } catch {
        Write-Log "Error analyzing Teams DO log: $_" "ERROR"
        $Buffers.Summary.Add([PSCustomObject]@{
            Test = "Teams Content Analysis"
            Result = "Analysis failed"
            Status = "ERROR"
            Impact = "Unable to determine Teams content impact"
        }) | Out-Null
    }
}

function Invoke-DOTroubleshooter {
    Write-Log "Running official DO Troubleshooter script..." "INFO"
    Show-Progress -Activity "Running DO Troubleshooter script" -PercentComplete 90
    
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    $offlineScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Scripts\DeliveryOptimizationTroubleshooter.ps1"
    $scriptBasePath = if ($isAdmin) { "C:\Program Files\WindowsPowerShell\Scripts" } else { Join-Path -Path $env:USERPROFILE -ChildPath "Documents\PowerShell\Scripts" }
    $scriptPath = Join-Path $scriptBasePath "DeliveryOptimizationTroubleshooter.ps1"
    $scriptInfoPath = Join-Path $scriptBasePath "InstalledScriptInfos"
    
    # Setup script environment
    if (-not (Test-Path $scriptBasePath)) { New-Item -Path $scriptBasePath -ItemType Directory -Force | Out-Null }
    if (-not (Test-Path $scriptInfoPath)) { New-Item -Path $scriptInfoPath -ItemType Directory -Force | Out-Null }
    
    # Ensure the troubleshooter script is available
    try {
        if (-not (Get-Command DeliveryOptimizationTroubleshooter.ps1 -ErrorAction SilentlyContinue)) {
            if (Test-Path $offlineScriptPath) {
                Write-Log "Found offline copy of DeliveryOptimizationTroubleshooter.ps1, copying to $scriptPath" "INFO"
                Copy-Item -Path $offlineScriptPath -Destination $scriptPath -Force
            } else {
                Write-Log "Attempting to download DeliveryOptimizationTroubleshooter.ps1..." "INFO"
                if (Get-Module -ListAvailable -Name PowerShellGet) {
                    try {
                        Install-Script -Name DeliveryOptimizationTroubleshooter -Force
                        Write-Log "Downloaded DeliveryOptimizationTroubleshooter.ps1 from PSGallery" "SUCCESS"
                    } catch {
                        Write-Log "Failed to download from PSGallery: $_" "ERROR"
                    }
                } else {
                    Write-Log "PowerShellGet not available, cannot download script" "WARN"
                }
            }
        }
        
        # Execute the troubleshooter
        $logPath = Join-Path $env:TEMP "DOT_FullScriptOutput.log"
        Write-Log "Executing DeliveryOptimizationTroubleshooter..." "INFO"
        & $scriptPath *> $logPath
        
        # Process results
        if (Test-Path $logPath) {
            $log = Get-Content -Path $logPath -Raw
            # Parse log for key information
            $logMatches = [regex]::Matches($log, "(?:PASS|FAIL|WARNING):\s+(.*?)(?=\r?\nP|\r?\nF|\r?\nW|\r?\n\r?\n|\z)")
            
            $results = @()
            foreach ($match in $logMatches) {
                $result = $match.Value
                $status = if ($result -match "^PASS") { "PASS" } elseif ($result -match "^FAIL") { "FAIL" } else { "WARN" }
                $message = $result -replace "^(?:PASS|FAIL|WARNING):\s+", ""
                
                $results += [PSCustomObject]@{
                    Status = $status
                    Message = $message
                    Impact = switch ($status) {
                        "PASS" { "No issues detected" }
                        "FAIL" { "Critical issue that needs attention" }
                        "WARN" { "Potential issue that may need investigation" }
                    }
                }
            }
            
            # Calculate statistics
            $passCount = ($results | Where-Object { $_.Status -eq "PASS" }).Count
            $warnCount = ($results | Where-Object { $_.Status -eq "WARN" }).Count
            $failCount = ($results | Where-Object { $_.Status -eq "FAIL" }).Count
            
            Write-Log "DeliveryOptimizationTroubleshooter results: $passCount passed, $warnCount warnings, $failCount failures" "INFO"
            
            # Add to summary
            $Buffers.Summary.Add([PSCustomObject]@{
                Test = "DO Troubleshooter Results"
                Result = "$passCount passed, $warnCount warnings, $failCount failures"
                Status = if ($failCount -gt 0) { "FAIL" } elseif ($warnCount -gt 0) { "WARN" } else { "PASS" }
                Impact = if ($failCount -gt 0) { "Critical issues detected" } elseif ($warnCount -gt 0) { "Potential issues detected" } else { "No issues detected" }
            }) | Out-Null
            
            $results | ForEach-Object { $Buffers.SysInfo.Add($_) | Out-Null }
        } else {
            Write-Log "DeliveryOptimizationTroubleshooter log not found" "WARN"
            $Buffers.Summary.Add([PSCustomObject]@{
                Test = "DO Troubleshooter"
                Result = "Execution failed or log not found"
                Status = "WARN"
                Impact = "Could not analyze official troubleshooter output"
            }) | Out-Null
        }
    } catch {
        Write-Log "Troubleshooter failed: $_" "ERROR"
        $Buffers.Summary.Add([PSCustomObject]@{
            Test = "DO Troubleshooter"
            Result = "Execution error"
            Status = "ERROR"
            Impact = "Could not run comprehensive troubleshooting"
        }) | Out-Null
    }
}
function New-ExecutiveSummary {
    Write-Log "Generating executive summary..." "INFO"
    
    $passed = ($Buffers.Summary | Where-Object { $_.Status -eq "PASS" }).Count
    $warnings = ($Buffers.Summary | Where-Object { $_.Status -eq "WARN" }).Count
    $criticalIssues = ($Buffers.Summary | Where-Object { $_.Status -eq "FAIL" -or $_.Status -eq "ERROR" -or $_.Status -eq "CRITICAL" }).Count
    
    $overallHealth = if ($criticalIssues -gt 0) {
        "CRITICAL - $criticalIssues critical issues detected"
    } elseif ($warnings -gt 0) {
        "WARNING - $warnings warnings detected"
    } else {
        "HEALTHY - All checks passed"
    }
    
    $summary = [PSCustomObject]@{
        SystemName = $env:COMPUTERNAME
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        OverallHealth = $overallHealth
        CriticalIssues = $criticalIssues
        Warnings = $warnings
        PassedChecks = $passed
        DODownloadMode = ($Buffers.HealthCheck | Where-Object { $null -ne $_.DODownloadMode } | Select-Object -First 1).DODownloadMode
        RecommendationsCount = $Buffers.Recommendations.Count
    }
    
    return $summary
}

function Export-DOExcelReport {
    if (-not (Test-Path -Path $OutputPath)) {
        try {
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
            Write-Log "Created output directory: $OutputPath" "INFO"
        } catch {
            Write-Log "Failed to create output directory: $_" "ERROR"
        }
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $excelPath = Join-Path -Path $OutputPath -ChildPath "DO_Report_$timestamp.xlsx"
    $csvFolder = Join-Path -Path $OutputPath -ChildPath "DO_Report_CSV_$timestamp"
    
    Write-Log "Exporting Excel report to $excelPath" "INFO"
    Show-Progress -Activity "Generating Excel report" -PercentComplete 95
    
    # Create Executive Summary
    $executiveSummary = New-ExecutiveSummary
    
    # Add all other sheets with enhanced formatting
    $executiveSummary | Export-Excel -Path $excelPath -WorksheetName "Summary" -AutoSize -TableStyle Medium9
    $Buffers.Summary | Export-Excel -Path $excelPath -WorksheetName "Test Results" -AutoSize -TableStyle Medium1 -Append
    $Buffers.Recommendations | Export-Excel -Path $excelPath -WorksheetName "Recommendations" -AutoSize -TableStyle Medium6 -Append
    $Buffers.HealthCheck | Export-Excel -Path $excelPath -WorksheetName "Health Check" -AutoSize -TableStyle Medium1 -Append
    $Buffers.MCC | Export-Excel -Path $excelPath -WorksheetName "Connected Cache" -AutoSize -TableStyle Medium3 -Append
    $Buffers.Peers | Export-Excel -Path $excelPath -WorksheetName "Peer Connectivity" -AutoSize -TableStyle Light9 -Append
    $Buffers.SysInfo | Export-Excel -Path $excelPath -WorksheetName "Log Analysis" -AutoSize -TableStyle Medium5 -Append
    $Buffers.DOConfig | Export-Excel -Path $excelPath -WorksheetName "DO Configuration" -AutoSize -TableStyle Medium5 -Append
    $Buffers.Errors | Export-Excel -Path $excelPath -WorksheetName "Errors" -AutoSize -TableStyle Medium3 -Append
    $Buffers.Log | Export-Excel -Path $excelPath -WorksheetName "Execution Log" -AutoSize -TableStyle Light8 -Append
    $Buffers.DiagnosticsData | Export-Excel -Path $excelPath -WorksheetName "Diagnostics Data" -AutoSize -TableStyle Medium4 -Append
    
    # Export DO error codes
    try {
        $errorTable = Get-DOErrorsTable
        $errorTable | Export-Excel -Path $excelPath -WorksheetName "Error Codes" -AutoSize -TableStyle Medium2 -Append
        Write-Log "Exported DO error code table with recommendations." "INFO"
    } catch {
        Write-Log "Failed to export DO error codes: $_" "ERROR"
    }
    
    # Export CSVs for additional analysis
    try {
        New-Item -Path $csvFolder -ItemType Directory -Force | Out-Null
        foreach ($key in $Buffers.Keys) {
            if ($Buffers[$key].Count -gt 0) {
                $csvPath = Join-Path -Path $csvFolder -ChildPath "$key.csv"
                $Buffers[$key] | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            }
        }
        Write-Log "CSV reports exported to $csvFolder for additional analysis" "INFO"
    } catch {
        Write-Log "CSV export failed: $_" "WARN"
    }
    
    Write-Host "`n📊 Report saved to: $excelPath" -ForegroundColor Cyan
    Write-Host "📋 CSV data exported to: $csvFolder" -ForegroundColor Cyan
    
    return $excelPath
}

# ───── DIAGNOSTICS ZIP PROCESSING ─────

function Expand-DiagnosticsZip {
    param (
        [string]$ZipPath,
        [string]$ExtractPath
    )
    
    Write-Log "Extracting diagnostics zip file..." "INFO"
    Show-Progress -Activity "Extracting diagnostics zip" -PercentComplete 5
    
    try {
        # Create extraction directory if it doesn't exist
        if (-not (Test-Path -Path $ExtractPath -PathType Container)) {
            New-Item -Path $ExtractPath -ItemType Directory -Force | Out-Null
        }
        
        # Extract the zip file
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $ExtractPath)
        Write-Log "Diagnostics zip extracted to $ExtractPath" "SUCCESS"
        return $true
    } catch {
        Write-Log "Error extracting diagnostics zip: $_" "ERROR"
        return $false
    }
}

function Invoke-DiagnosticsDataProcessing {
    param (
        [string]$DiagnosticsPath
    )
    
    Write-Log "Processing diagnostics data from extracted folder..." "INFO"
    Show-Progress -Activity "Processing diagnostics data" -PercentComplete 15
    
    try {
        # Look for DO-related files in the diagnostics folder
        $doFiles = Get-ChildItem -Path $DiagnosticsPath -Recurse -File | Where-Object {
            $_.Name -match "DeliveryOptimization|DOSvc|DO_" -or 
            $_.FullName -match "\\Windows\\Logs\\DISM\\|\\WindowsUpdate\\|\\Microsoft-Windows-DeliveryOptimization"
        }
        
        Write-Log "Found $($doFiles.Count) DO-related files in diagnostics package" "SUCCESS"
        
        # Process each file
        foreach ($file in $doFiles) {
            try {
                # Basic file type detection
                $fileType = if ($file.Name -match "\.log$") {
                    "Log File"
                } elseif ($file.Name -match "\.xml$") {
                    "Configuration"
                } elseif ($file.Name -match "\.etl$") {
                    "Event Trace Log"
                } else {
                    "Other"
                }
                
                # Extract basic stats
                $fileContent = Get-Content -Path $file.FullName -Raw -ErrorAction SilentlyContinue
                $lineCount = if ($null -ne $fileContent) { ($fileContent -split "`n").Count } else { 0 }
                $errorCount = if ($null -ne $fileContent) { [regex]::Matches($fileContent, "(?i)(error|fail|exception|0x8|critical)").Count } else { 0 }
                
                $fileSummary = [PSCustomObject]@{
                    FileName = $file.Name
                    FilePath = $file.FullName
                    FileType = $fileType
                    SizeKB = [math]::Round($file.Length / 1KB, 2)
                    LineCount = $lineCount
                    ErrorCount = $errorCount
                    ContainsErrors = $errorCount -gt 0
                    CreationTime = $file.CreationTime
                    LastWriteTime = $file.LastWriteTime
                }
                
                $Buffers.DiagnosticsData.Add($fileSummary) | Out-Null
                
                # Add recommendation for files with many errors
                if ($errorCount -gt 10) {
                    Write-Log "File $($file.Name) contains $errorCount potential errors/warnings" "WARN"
                    Add-Recommendation -Area "Diagnostics" -Recommendation "Review $($file.Name) which contains multiple errors ($errorCount found)" -Severity "Important"
                }
            } catch {
                Write-Log "Failed to process file $($file.Name): $_" "WARN"
            }
        }
        
        # Add summary of diagnostics data
        $Buffers.Summary.Add([PSCustomObject]@{
            Test = "Diagnostics Data Analysis"
            Result = "$($doFiles.Count) DO-related files found"
            Status = "INFO"
            Impact = "Additional diagnostic information available"
        }) | Out-Null
      } catch {
        Write-Log "Error processing diagnostics data: $_" "ERROR"
        $Buffers.Summary.Add([PSCustomObject]@{
            Test = "Diagnostics Data Analysis"
            Result = "Analysis failed"
            Status = "ERROR"
            Impact = "Limited diagnostic information available"
        }) | Out-Null
    }
    
    if ($doFiles.Count -eq 0) {
        $Buffers.DiagnosticsData.Add([PSCustomObject]@{
            Type = "Warning"
            Message = "No DO-related files found in diagnostics package"
        }) | Out-Null
        Write-Log "No Delivery Optimization related files found in diagnostics package" "WARN"
    }
}

# ───── MAIN EXECUTION ─────

Write-Host "`n🟦 Starting Delivery Optimization Troubleshooting..." -ForegroundColor Cyan
$startTime = Get-Date

Write-Host "▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬" -ForegroundColor Blue
Write-Host "    System: $env:COMPUTERNAME | User: $env:USERNAME | Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor DarkCyan
Write-Host "    Output Path: $OutputPath" -ForegroundColor DarkCyan
Write-Host "    " -NoNewline

# Create output directory if it doesn't exist
if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

# Process diagnostics zip if provided
$extractPath = $null
if ($DiagnosticsZip -and (Test-Path -Path $DiagnosticsZip)) {
    Write-Log "Diagnostics zip file provided: $DiagnosticsZip" "INFO"
    # Extract and process the zip file
    $extractPath = Join-Path -Path $env:TEMP -ChildPath "DODiagnostics_$([Guid]::NewGuid().ToString())"
    $extractSuccess = Expand-DiagnosticsZip -ZipPath $DiagnosticsZip -ExtractPath $extractPath
    if ($extractSuccess) {
        Invoke-DiagnosticsDataProcessing -DiagnosticsPath $extractPath
    }
} elseif ($DiagnosticsZip) {
    Write-Log "Specified DiagnosticsZip file not found: $DiagnosticsZip" "WARN"
}

# Run all diagnostics sequentially
Get-SystemInfo
Test-DOEndpoints
Test-DNSDPeerConfig
Test-DOGroupId
Start-HealthCheck
Test-PeerConnectivity
Test-MCCCheck
Invoke-DOLogAnalysis
Test-TeamsImpact
Invoke-DOTroubleshooter

# Generate and export the report
$reportPath = Export-DOExcelReport

# Calculate statistics
$endTime = Get-Date
$duration = $endTime - $startTime
$minutes = [math]::Floor($duration.TotalMinutes)
$seconds = $duration.Seconds

# Get overall status for final output
$criticalIssues = ($Buffers.Summary | Where-Object { $_.Status -eq "FAIL" -or $_.Status -eq "ERROR" -or $_.Status -eq "CRITICAL" }).Count
$warnings = ($Buffers.Summary | Where-Object { $_.Status -eq "WARN" }).Count

# Display completion message
Write-Host "`n✅ Delivery Optimization Troubleshooting completed in $minutes min $seconds sec" -ForegroundColor Green

# Display final status
if ($criticalIssues -gt 0) {
    Write-Host "⚠️ Found $criticalIssues critical issues that require attention" -ForegroundColor Red
    Write-Host "   Review the 'Recommendations' sheet for remediation steps" -ForegroundColor Red
} elseif ($warnings -gt 0) {
    Write-Host "⚠️ Found $warnings warnings that may need attention" -ForegroundColor Yellow
    Write-Host "   Review the 'Recommendations' sheet for potential improvements" -ForegroundColor Yellow
} else {
    Write-Host "👍 No issues detected. Delivery Optimization appears to be configured correctly" -ForegroundColor Green
}

Write-Host "`n🟦 Delivery Optimization Troubleshooter Complete 🟦" -ForegroundColor Cyan

# Open report if requested
if ($Show) {
    Write-Host "`n🔍 Opening Excel report..." -ForegroundColor Cyan
    Invoke-Item $reportPath
}

# Clean up temporary files if they exist
if ($extractPath -and (Test-Path -Path $extractPath)) {
    Remove-TemporaryFiles -Path $extractPath
}
