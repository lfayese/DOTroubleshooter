<#
.SYNOPSIS
  Generates a comprehensive Delivery Optimization troubleshooting report in an Excel workbook.

.DESCRIPTION
  This script conducts an in-depth analysis of Delivery Optimization (DO) on the local system and compiles the results into an Excel report using the ImportExcel module. It performs diagnostics across multiple areas, including:

    • Service health and status via Get-DeliveryOptimizationStatus.
    • Peer-to-peer connectivity tests on key DO ports (7680 and 3544) using parallel processing.
    • Validation of Microsoft Connected Cache (MCC) configuration.
    • DNS-SD configuration check for DO peer discovery.
    • Detection of DO Group ID settings, with warnings for potential DHCP-based group issues.
    • Endpoint availability tests for critical DO URLs.
    • Analysis of DO logs for non-peerable Teams traffic.
    • Execution and parsing of an offline Delivery Optimization Troubleshooter script.
    • Detailed DO log analysis and connection statistics.
    • Export of a lookup table for common DO error codes.

  The script leverages parallel processing to enhance performance during network tests and logs all diagnostic activities. Detailed logs are captured in multiple buffers and then exported both to the Excel report and as CSV files for further analysis. If the ImportExcel module is not available on the system, the script attempts to import a bundled version from a local "Modules\ImportExcel" directory.

.PARAMETER OutputPath
  Optional. Specifies the directory where the Excel report and accompanying CSV exports will be saved.
  Defaults to the current user's Desktop if not provided.

.PARAMETER Show
  Optional. If specified, the script automatically opens the generated Excel report upon completion.

.PARAMETER DiagnosticsZip
  Optional. Path to a diagnostics zip file containing Delivery Optimization logs for analysis.

.EXAMPLE
  .\Invoke-DoTroubleshooter.ps1 -OutputPath "C:\DOReports" -Show
  Runs the complete DO troubleshooting process, saves the report to "C:\DOReports", and opens the Excel report after processing.

.EXAMPLE
  .\Invoke-DoTroubleshooter.ps1 -DiagnosticsZip "C:\Path\To\DiagnosticsFile.zip" -Show
  Processes a diagnostics zip file and includes the findings in the report.

.NOTES
  - Requires PowerShell 5 or newer.
  - The script uses several DO-related cmdlets (e.g., Get-DeliveryOptimizationStatus, Get-DeliveryOptimizationLogAnalysis) that may require elevated privileges.
  - Parallel processing is employed for endpoint and peer connectivity tests to minimize runtime.
  - In the event of errors during diagnostics, the script logs detailed error messages for each category.
  - CSV exports provide a text-based backup of all buffers for external analysis.
  - For best results, run the script on a system where DO is actively configured and operational.
  - Troubleshooting output from the offline Delivery Optimization Troubleshooter is parsed into structured results for easier analysis.
  - Ensure that your environment meets all prerequisites, including network connectivity and administrative rights.

.TROUBLESHOOTING
  • If the report does not open automatically when using the -Show parameter, verify that the Excel file path is valid and that Excel is installed.
  • In case of missing DO logs or configuration data, ensure that the necessary DO services are running and that the local system is properly configured for Delivery Optimization.
  • If the ImportExcel module cannot be imported from the local directory, confirm that the module files are present in the expected "Modules\ImportExcel" folder relative to the script location.

.LINK
  https://learn.microsoft.com/en-us/windows/deployment/optimization/delivery-optimization

#>

param (
  [string]$OutputPath = [Environment]::GetFolderPath("Desktop"),
  [switch]$Show,
  [string]$DiagnosticsZip
)

# Use bundled ImportExcel if not available
$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "Modules\ImportExcel"
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
  Write-Host "[INFO] ImportExcel not found, importing from local module path..." -ForegroundColor Yellow
  Import-Module -Name $modulePath -Force
} else {
  Import-Module ImportExcel
}

# ───── GLOBAL BUFFERS ─────
$Buffers = @{
  Log         = [System.Collections.ArrayList]::new()
  HealthCheck = [System.Collections.ArrayList]::new()
  P2P         = [System.Collections.ArrayList]::new()
  MCC         = [System.Collections.ArrayList]::new()
  SysInfo     = [System.Collections.ArrayList]::new()
  Errors      = [System.Collections.ArrayList]::new()
  Stats       = [System.Collections.ArrayList]::new()
  DOConfig    = [System.Collections.ArrayList]::new()
  Peers       = [System.Collections.ArrayList]::new()
  Summary     = [System.Collections.ArrayList]::new() # New buffer for summary information
  Recommendations = [System.Collections.ArrayList]::new() # New buffer for recommendations
  DiagnosticsData = [System.Collections.ArrayList]::new() # New buffer for diagnostic data
}

# ───── LOGGING FUNCTION ─────
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

# ───── ADD RECOMMENDATION ─────
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

# ───── PROGRESS INDICATOR ─────
function Show-Progress {
  param (
    [string]$Activity,
    [int]$PercentComplete
  )

  Write-Progress -Activity "DO Troubleshooter" -Status $Activity -PercentComplete $PercentComplete
}

# ───── SYSTEM INFORMATION COLLECTION ─────
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
      CollectionTime  = Get-Date
      UserName        = "$env:USERDOMAIN\$env:USERNAME"
    }

    $Buffers.Summary.Add($sysInfo) | Out-Null
    Write-Log "System information collected successfully" "SUCCESS"
  }
  catch {
    Write-Log "Error collecting system information: $_" "ERROR"
  }
}

# ───── PARALLEL ENDPOINT TESTING ─────
function Test-DOEndpoints {
  Write-Log "Validating DO endpoint connectivity (parallel processing)..." "INFO"
  Show-Progress -Activity "Testing DO endpoints" -PercentComplete 20

  # Enhanced with descriptions for better context
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
        $result = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 5
        [PSCustomObject]@{
          Url         = $url
          Description = $description
          Required    = $required
          Status      = $result.StatusCode
          Success     = $true
          Error       = $null
        }
      } catch {
        [PSCustomObject]@{
          Url         = $url
          Description = $description
          Required    = $required
          Status      = 0
          Success     = $false
          Error       = $_.Exception.Message
        }
      }
    } -ArgumentList $endpoint.Url, $endpoint.Description, $endpoint.Required
  }

  # Collect results from jobs
  $results = @()
  foreach ($job in $jobs) {
    $result = Receive-Job -Job $job -Wait
    $results += $result

    if ($result.Success) {
      Write-Log "Success: $($result.Url) reachable (Status: $($result.Status))" "SUCCESS"
    } else {
      $severity = if ($result.Required) { "ERROR" } else { "WARN" }
      Write-Log "FAILED to reach $($result.Url): $($result.Error)" $severity

      if ($result.Required) {
        Add-Recommendation -Area "Network" -Recommendation "Critical DO endpoint $($result.Url) is unreachable. Please check network connectivity and firewall rules." -Severity "Critical" -Reference "https://learn.microsoft.com/en-us/windows/deployment/upgrade/upgrade-analytics-get-started#enable-data-sharing"
      }
    }

    Remove-Job -Job $job
  }

  # Add results to DOConfig buffer
  $results | ForEach-Object { $Buffers.DOConfig.Add($_) | Out-Null }

  # Add summary information
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

# ───── DNS-SD CONFIGURATION CHECK ─────
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
        Status      = "Enabled"
        Description = "DNS-SD peer discovery is properly configured"
        Impact      = "Optimal peer selection within network"
      }

      $Buffers.Summary.Add([PSCustomObject]@{
        Test = "DNS-SD Configuration"
        Result = "Properly configured (value=2)"
        Status = "PASS"
        Impact = "Optimal peer discovery"
      }) | Out-Null
    } else {
      Write-Log "DORestrictPeerSelectionBy not set to 2. DNS-SD likely disabled (value=$val)." "WARN"
      $config = [PSCustomObject]@{
        Setting     = "DORestrictPeerSelectionBy"
        Value       = $val
        Status      = "Misconfigured"
        Description = "DNS-SD peer discovery should be set to 2"
        Impact      = "Limited peer discovery capabilities"
      }

      Add-Recommendation -Area "Configuration" -Recommendation "Set DORestrictPeerSelectionBy registry value to 2 for optimal peer discovery" -Severity "Important" -Reference "https://learn.microsoft.com/en-us/windows/deployment/optimization/delivery-optimization-reference#delivery-optimization-options"

      $Buffers.Summary.Add([PSCustomObject]@{
        Test = "DNS-SD Configuration"
        Result = "Misconfigured (value=$val)"
        Status = "WARN"
        Impact = "Reduced peer discovery capabilities"
      }) | Out-Null
    }
  } catch {
    Write-Log "DNS-SD config missing or inaccessible." "WARN"
    $config = [PSCustomObject]@{
      Setting     = "DORestrictPeerSelectionBy"
      Value       = "N/A"
      Status      = "Missing"
      Description = "DNS-SD peer discovery setting not found"
      Impact      = "Potential peer discovery limitations"
    }

    Add-Recommendation -Area "Configuration" -Recommendation "Configure DNS-SD peer discovery by setting DORestrictPeerSelectionBy registry value to 2" -Severity "Important" -Reference "https://learn.microsoft.com/en-us/windows/deployment/optimization/delivery-optimization-reference#delivery-optimization-options"

    $Buffers.Summary.Add([PSCustomObject]@{
      Test = "DNS-SD Configuration"
      Result = "Setting not found"
      Status = "WARN"
      Impact = "Potential peer discovery limitations"
    }) | Out-Null
  }
  $Buffers.DOConfig.Add($config) | Out-Null
}

# ───── GROUP ID DETECTION ─────
function Test-DOGroupId {
  Write-Log "Checking DO Group ID via policy..." "INFO"
  Show-Progress -Activity "Checking DO Group ID" -PercentComplete 40

  try {
    $groupId = ([Windows.Management.Policies.NamedPolicy]::GetPolicyFromPath("DeliveryOptimization", "DOGroupId")).GetString()
    if ([string]::IsNullOrEmpty($groupId)) {
      Write-Log "DOGroupId not defined in policy. If using DHCP Option 234, logs may show NULL." "WARN"
      $config = [PSCustomObject]@{
        Setting     = "DOGroupId"
        Value       = "NULL"
        Status      = "Warning"
        Description = "No group ID defined. If using DHCP Option 234, verify DHCP server configuration."
        Impact      = "May limit cross-subnet peer sharing if DHCP Option 234 isn't configured"
      }

      Add-Recommendation -Area "Configuration" -Recommendation "Configure DOGroupId via policy or DHCP Option 234 for optimal cross-subnet peer sharing" -Severity "Important" -Reference "https://learn.microsoft.com/en-us/windows/deployment/optimization/delivery-optimization-reference#delivery-optimization-options"

      $Buffers.Summary.Add([PSCustomObject]@{
        Test = "DO Group ID"
        Result = "Not defined in policy"
        Status = "WARN"
        Impact = "May limit cross-subnet peer capabilities"
      }) | Out-Null
    } else {
      Write-Log "DOGroupId found: $groupId" "SUCCESS"
      $config = [PSCustomObject]@{
        Setting     = "DOGroupId"
        Value       = $groupId
        Status      = "Configured"
        Description = "Group ID is set through policy"
        Impact      = "Enables optimal cross-subnet peer sharing"
      }

      $Buffers.Summary.Add([PSCustomObject]@{
        Test = "DO Group ID"
        Result = "Configured: $groupId"
        Status = "PASS"
        Impact = "Optimal cross-subnet peer sharing"
      }) | Out-Null
    }
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
  $Buffers.DOConfig.Add($config) | Out-Null
}

# ───── HEALTH CHECK ─────
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

    # Enrich the status object with descriptions and recommendations
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
      Recommendation = if ($status.DODownloadMode -eq 99) {
        "DO service is in fallback mode. Restart the service or check for system issues."
      } elseif ($status.DODownloadMode -eq 0) {
        "No peering enabled. Consider enabling peer caching for bandwidth savings."
      } else {
        "Configuration appears optimal."
      }
    }

    $Buffers.HealthCheck.Add($enrichedStatus) | Out-Null
    $Buffers.Stats.Add([PSCustomObject]@{
      Category       = "Health"
      DODownloadMode = $status.DODownloadMode
      NumberOfPeers  = $status.NumberOfPeers
    }) | Out-Null

    Write-Log "DODownloadMode: $($status.DODownloadMode) - $modeDesc" "INFO" "HealthCheck"

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

# ───── PARALLEL PEER CONNECTIVITY TESTING ─────
function Test-PeerConnectivity {
  Write-Log "Testing P2P ports (7680, 3544) using parallel processing..." "INFO" "P2P"
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
        param($peer, $port, $description)
        $test = Test-NetConnection -ComputerName $peer -Port $port -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        [PSCustomObject]@{
          Peer            = $peer
          Port            = $port
          Description     = $description
          TcpTestSucceeded= $test.TcpTestSucceeded
          PingSucceeded   = $test.PingSucceeded
          RemoteAddress   = $test.RemoteAddress
          Impact          = if ($test.TcpTestSucceeded) { "None" } else { "May limit peer-to-peer capabilities" }
        }
      } -ArgumentList $peer, $portInfo.Port, $portInfo.Description
    }
  }

  $portResults = @{
    7680 = [PSCustomObject]@{Success = 0; Failure = 0}
    3544 = [PSCustomObject]@{Success = 0; Failure = 0}
  }

  foreach ($job in $jobs) {
    $result = Receive-Job -Job $job -Wait
    $Buffers.Peers.Add($result) | Out-Null
    $status = if ($result.TcpTestSucceeded) { "SUCCESS" } else { "FAILED" }
    $level = if ($result.TcpTestSucceeded) { "SUCCESS" } else { "WARN" }

    if ($result.TcpTestSucceeded) {
      $portResults[$result.Port].Success++
    } else {
      $portResults[$result.Port].Failure++
    }

    Write-Log "Peer $($result.Peer) | Port $($result.Port) | TCP: $status | $($result.Description)" $level "P2P"
    Remove-Job -Job $job
  }

  foreach ($port in $ports.Port) {
    $total = $portResults[$port].Success + $portResults[$port].Failure
    $successRate = if ($total -gt 0) { [math]::Round(($portResults[$port].Success / $total) * 100, 0) } else { 0 }

    $Buffers.Summary.Add([PSCustomObject]@{
      Test = "Port $port Connectivity"
      Result = "$($portResults[$port].Success) of $total connections successful ($successRate%)"
      Status = if ($successRate -gt 50) { "PASS" } elseif ($successRate -gt 0) { "WARN" } else { "FAIL" }
      Impact = if ($successRate -eq 100) { "None" } elseif ($successRate -gt 50) { "Minor impact on peer-to-peer" } else { "Significant impact on peer-to-peer capabilities" }
    }) | Out-Null

    if ($successRate -lt 50) {
      Add-Recommendation -Area "Network" -Recommendation "Check firewall rules for port $port to ensure peer-to-peer connectivity" -Severity "Important" -Reference "https://learn.microsoft.com/en-us/windows/deployment/optimization/delivery-optimization-reference#ports"
    }
  }
}

# ───── MCC CHECK ─────
function Test-MCCCheck {
  Write-Log "Checking MCC (Microsoft Connected Cache)..." "INFO" "MCC"
  Show-Progress -Activity "Checking Microsoft Connected Cache" -PercentComplete 70
  try {
    $diag = Get-DOConfig -Verbose

    # Create a more structured and readable version of the MCC configuration
    $mccConfig = [PSCustomObject]@{
      ConnectedCacheServers = if ($diag.PSObject.Properties.Name -contains "ConnectedCacheServers") { $diag.ConnectedCacheServers } else { "None configured" }
      Status = if (($diag.PSObject.Properties.Name -contains "ConnectedCacheServers") -and $diag.ConnectedCacheServers) { "Configured" } else { "Not configured" }
      Description = "Microsoft Connected Cache allows on-premises caching of content"
      Impact = if (($diag.PSObject.Properties.Name -contains "ConnectedCacheServers") -and $diag.ConnectedCacheServers) {
        "Reduces internet bandwidth by serving content from local cache servers"
      } else {
        "All content must be downloaded from the internet, potentially increasing bandwidth usage"
      }
      Recommendation = if (-not (($diag.PSObject.Properties.Name -contains "ConnectedCacheServers") -and $diag.ConnectedCacheServers)) {
        "Consider setting up Microsoft Connected Cache servers for bandwidth optimization if you have multiple devices"
      } else {
        "Current configuration is optimal"
      }
    }

    $Buffers.MCC.Add($mccConfig) | Out-Null
    $Buffers.MCC.Add($diag) | Out-Null

    $Buffers.Stats.Add([PSCustomObject]@{
      Category = "MCC"
      Config   = ($diag | Out-String).Trim()
    }) | Out-Null

    if ($diag.PSObject.Properties.Name -contains "ConnectedCacheServers") {
      if ($diag.ConnectedCacheServers) {
        Write-Log "Connected Cache servers found: $($diag.ConnectedCacheServers)" "SUCCESS" "MCC"
        $Buffers.Summary.Add([PSCustomObject]@{
          Test = "Microsoft Connected Cache"
          Result = "Configured"
          Status = "PASS"
          Impact = "Bandwidth savings from on-premises content caching"
        }) | Out-Null
      } else {
        Write-Log "No Connected Cache servers configured" "WARN" "MCC"
        Add-Recommendation -Area "Optimization" -Recommendation "Consider configuring Microsoft Connected Cache for bandwidth optimization" -Severity "Informational" -Reference "https://learn.microsoft.com/en-us/mem/configmgr/core/servers/deploy/configure/microsoft-connected-cache"

        $Buffers.Summary.Add([PSCustomObject]@{
          Test = "Microsoft Connected Cache"
          Result = "Not configured"
          Status = "INFO"
          Impact = "Potential for bandwidth optimization"
        }) | Out-Null
      }
    } else {
      Write-Log "Connected Cache configuration not found in DO config" "WARN" "MCC"
      $Buffers.Summary.Add([PSCustomObject]@{
        Test = "Microsoft Connected Cache"
        Result = "Configuration not found"
        Status = "INFO"
        Impact = "Potential for bandwidth optimization"
      }) | Out-Null
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

# ───── TEAMS IMPACT ANALYSIS ─────
function Test-TeamsImpact {
  Write-Log "Analyzing DO log for Teams content..." "INFO"
  Show-Progress -Activity "Analyzing Teams content in DO logs" -PercentComplete 75

  try {
    $teamLogs = Get-DeliveryOptimizationLog | Where-Object { $_.Message -match "Teams" }
    if ($teamLogs) {
      $count = $teamLogs.Count

      # Enrich the logs with impact information
      $enrichedLogs = $teamLogs | ForEach-Object {
        [PSCustomObject]@{
          Timestamp = $_.Timestamp
          Level = $_.Level
          Message = $_.Message
          Impact = "Teams content is not peerable, reducing bandwidth savings potential"
          Recommendation = "This is expected behavior as Teams content is not eligible for peer caching"
        }
      }

      $enrichedLogs | ForEach-Object { $Buffers.SysInfo.Add($_) | Out-Null }

      Write-Log "Found $count Teams-related DO log entries (non-peerable traffic)" "WARN"

      $timestamps = $teamLogs | ForEach-Object { $_.Timestamp } | Select-Object -First 5
      $recentTeams = $timestamps -join ", "
      Write-Log "Recent Teams activity: $recentTeams" "INFO"

      $Buffers.Summary.Add([PSCustomObject]@{
        Test = "Teams Content Analysis"
        Result = "$count entries found"
        Status = "INFO"
        Impact = "Teams content cannot use peer caching (expected behavior)"
      }) | Out-Null

      Add-Recommendation -Area "Content" -Recommendation "Teams content cannot use peer caching. This is expected behavior and not an issue." -Severity "Informational"
    } else {
      Write-Log "No Teams DO content found in logs" "INFO"
      $Buffers.Summary.Add([PSCustomObject]@{
        Test = "Teams Content Analysis"
        Result = "No entries found"
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

# ───── DO ERROR CODES TABLE ─────
function Get-DOErrorsTable {
  $errorsObj = @'
[
  { "ErrorCode": "0x80D01001", "Description": "Delivery Optimization was unable to provide the service.", "Recommendation": "Check if the DO service is running and properly configured." },
  { "ErrorCode": "0x80D02002", "Description": "Download of a file saw no progress within the defined period.", "Recommendation": "Check network connectivity and available bandwidth." },
  { "ErrorCode": "0x80D02003", "Description": "Job was not found.", "Recommendation": "The download job may have been cancelled or timed out." },
  { "ErrorCode": "0x80D02004", "Description": "There were no files in the job.", "Recommendation": "Verify that the content is available for download." },
  { "ErrorCode": "0x80D02005", "Description": "No downloads currently exist.", "Recommendation": "No action needed - informational only." },
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
      ErrorCode     = $_.ErrorCode
      Description   = $_.Description
      Recommendation = $_.Recommendation
      Impact        = "Potential download failures or service interruptions"
    }
  }
}

# ───── DO LOG ANALYSIS ─────
function Invoke-DOLogAnalysis {
  Write-Log "Running DO log analysis..." "INFO"
  Show-Progress -Activity "Analyzing DO logs" -PercentComplete 80

  try {
    $results = Get-DeliveryOptimizationLogAnalysis -ListConnections

    # Safely categorize and enrich the log data
    $categorizedResults = @()
    foreach ($result in $results) {
      $category = "Failed"
      if ($result.Result -eq "Success") {
        if ($result.Source -match "^(10\.|172\.(1[6-9]|2[0-9]|3[0-1])\.|192\.168\.)") {
          $category = "Local Peer"
        } else {
          $category = "Internet Peer"
        }
      }

      $impact = if ($result.Result -eq "Success") {
        "Bandwidth saved through peer sharing"
      } else {
        "Connection failed - no bandwidth savings"
      }

      $categorizedResults += [PSCustomObject]@{
        Time = $result.Time
        Source = $result.Source
        Destination = $result.Destination
        Result = $result.Result
        Bytes = $result.Bytes
        SuccessCategory = $category
        Impact = $impact
      }
    }

    $categorizedResults | ForEach-Object { $Buffers.SysInfo.Add($_) | Out-Null }

    $total   = $results.Count
    $success = ($results | Where-Object { $_.Result -eq "Success" }).Count
    $failed  = $total - $success

    $successRate = if ($total -gt 0) { [math]::Round(($success / $total) * 100, 0) } else { 0 }

    Write-Log "Log analysis complete: $success successful ($successRate%), $failed failed connections" (if($successRate -gt 75) {"SUCCESS"} elseif($successRate -gt 25) {"INFO"} else {"WARN"})

    # Add detailed stats
    $localPeers = ($categorizedResults | Where-Object { $_.SuccessCategory -eq "Local Peer" }).Count
    $internetPeers = ($categorizedResults | Where-Object { $_.SuccessCategory -eq "Internet Peer" }).Count
    $totalBytes = ($categorizedResults | Measure-Object -Property Bytes -Sum).Sum

    $Buffers.Summary.Add([PSCustomObject]@{
      Test = "Peer Connection Analysis"
      Result = "$success of $total connections successful ($successRate%)"
      Status = if ($successRate -gt 75) { "PASS" } elseif ($successRate -gt 25) { "WARN" } else { "FAIL" }
      Impact = if ($successRate -gt 75) { "Effective peer sharing" } elseif ($successRate -gt 25) { "Partial peer sharing benefits" } else { "Limited peer sharing benefits" }
    }) | Out-Null

    $Buffers.Summary.Add([PSCustomObject]@{
      Test = "Peer Distribution"
      Result = "Local: $localPeers, Internet: $internetPeers"
      Status = "INFO"
      Impact = "Data on peer source distribution"
    }) | Out-Null

    $Buffers.Summary.Add([PSCustomObject]@{
      Test = "Data Transferred"
      Result = "$([math]::Round($totalBytes/1MB, 2)) MB"
      Status = "INFO"
      Impact = "Total data transferred via peers"
    }) | Out-Null

    if ($successRate -lt 25) {
      Add-Recommendation -Area "Peer Connectivity" -Recommendation "Low peer connection success rate. Check network configuration and firewall rules." -Severity "Important"
    }
  } catch {
    Write-Log "Log analysis failed: $_" "ERROR"
    $Buffers.Summary.Add([PSCustomObject]@{
      Test = "Peer Connection Analysis"
      Result = "Analysis failed"
      Status = "ERROR"
      Impact = "Unable to determine peer connection effectiveness"
    }) | Out-Null

    Add-Recommendation -Area "Diagnostics" -Recommendation "Run the analysis with administrative privileges to access DO logs." -Severity "Informational"
  }
}

# ───── TROUBLESHOOTER EXECUTION ─────
function Invoke-DOTroubleshooter {
  Write-Log "Checking for DeliveryOptimizationTroubleshooter script..." "INFO"
  Show-Progress -Activity "Running DO Troubleshooter script" -PercentComplete 90

  $offlineScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Scripts\DeliveryOptimizationTroubleshooter.ps1"
  $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
  if ($isAdmin) {
    $scriptBasePath = "C:\Program Files\WindowsPowerShell\Scripts"
  } else {
    $scriptBasePath = Join-Path -Path $env:USERPROFILE -ChildPath "Documents\PowerShell\Scripts"
  }
  $scriptPath     = Join-Path $scriptBasePath "DeliveryOptimizationTroubleshooter.ps1"
  $scriptInfoPath = Join-Path $scriptBasePath "InstalledScriptInfos"
  try {
    if (-not (Test-Path $scriptBasePath)) { New-Item -Path $scriptBasePath -ItemType Directory -Force | Out-Null }
    if (-not (Test-Path $scriptInfoPath)) { New-Item -Path $scriptInfoPath -ItemType Directory -Force | Out-Null }
    if (-not (Get-Command DeliveryOptimizationTroubleshooter.ps1 -ErrorAction SilentlyContinue)) {
      if (Test-Path $offlineScriptPath) {
        Write-Log "Using offline DeliveryOptimizationTroubleshooter script from package." "INFO"
        Copy-Item -Path $offlineScriptPath -Destination $scriptPath -Force
        Copy-Item -Path "$PSScriptRoot\Scripts\DeliveryOptimizationTroubleshooter_InstalledScriptInfo.xml" `
            -Destination (Join-Path $scriptInfoPath "DeliveryOptimizationTroubleshooter_InstalledScriptInfo.xml") -Force
        Write-Log "Script and metadata copied to: $scriptBasePath" "INFO"
      } else {
        Write-Log "Offline Troubleshooter script not found in package." "ERROR"
        $Buffers.Summary.Add([PSCustomObject]@{
          Test = "DO Troubleshooter"
          Result = "Script not found"
          Status = "ERROR"
          Impact = "Could not run comprehensive troubleshooting"
        }) | Out-Null
        return
      }
    }
    $logPath = Join-Path $env:TEMP "DOT_FullScriptOutput.log"
    Write-Log "Executing DeliveryOptimizationTroubleshooter (all verifications)..." "INFO"
    & $scriptPath *> $logPath
    if (Test-Path $logPath) {
      $logContent = Get-Content $logPath
      $currentSection = ""
      $issuesFound = 0
      $passedChecks = 0

      # Enhanced parsing with result categorization
      foreach ($line in $logContent) {
        if ($line -match '^\s*-{5,}\s*$') { continue }
        if ($line -match "^\[.+\]") {
          $currentSection = $line.Trim()
          continue
        }

        $resultType = "INFO"
        if ($line -match "PASS|SUCCESS") {
          $resultType = "PASS"
          $passedChecks++
        } elseif ($line -match "FAIL|ERROR|CRITICAL") {
          $resultType = "FAIL"
          $issuesFound++
        } elseif ($line -match "WARN|WARNING") {
          $resultType = "WARN"
          $issuesFound++
        }

        if ($line.Trim()) {
          $Buffers.HealthCheck.Add([PSCustomObject]@{
            Timestamp = Get-Date
            Section   = $currentSection
            Message   = $line.Trim()
            ResultType = $resultType
            Recommendation = if ($resultType -eq "FAIL") {
              "Review and address this issue - it may impact DO functionality"
            } elseif ($resultType -eq "WARN") {
              "Consider addressing this warning if experiencing DO issues"
            } else {
              ""
            }
          }) | Out-Null
        }
      }

      Write-Log "DO Troubleshooter executed: $passedChecks checks passed, $issuesFound issues found" "INFO"

      $Buffers.Summary.Add([PSCustomObject]@{
        Test = "DO Troubleshooter"
        Result = "$passedChecks passed, $issuesFound issues found"
        Status = if ($issuesFound -eq 0) { "PASS" } elseif ($issuesFound -lt 3) { "WARN" } else { "FAIL" }
        Impact = if ($issuesFound -eq 0) { "No issues detected" } elseif ($issuesFound -lt 3) { "Minor issues may affect performance" } else { "Multiple issues detected that may impact functionality" }
      }) | Out-Null

      if ($issuesFound -gt 0) {
        Add-Recommendation -Area "General" -Recommendation "Review the DO Troubleshooter results in the HealthCheck tab for detailed findings" -Severity "Important"
      }
    } else {
      Write-Log "Troubleshooter ran, but no log file found." "WARN"
      $Buffers.Summary.Add([PSCustomObject]@{
        Test = "DO Troubleshooter"
        Result = "Log not found"
        Status = "WARN"
        Impact = "Could not analyze troubleshooter results"
      }) | Out-Null
    }
  } catch {
    Write-Log "Troubleshooter failed: $_" "ERROR"
    $Buffers.Summary.Add([PSCustomObject]@{
      Test = "DO Troubleshooter"
      Result = "Execution failed"
      Status = "ERROR"
      Impact = "Could not run comprehensive troubleshooting"
    }) | Out-Null
  }
}

# ───── CREATE EXECUTIVE SUMMARY ─────
function New-ExecutiveSummary {
  Write-Log "Generating executive summary..." "INFO"

  $criticalIssues = ($Buffers.Summary | Where-Object { $_.Status -eq "FAIL" -or $_.Status -eq "ERROR" -or $_.Status -eq "CRITICAL" }).Count
  $warnings = ($Buffers.Summary | Where-Object { $_.Status -eq "WARN" }).Count
  $passed = ($Buffers.Summary | Where-Object { $_.Status -eq "PASS" }).Count

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

# ───── EXPORT EXCEL REPORT ─────
function Export-DOExcelReport {
  if (-not (Test-Path -Path $OutputPath)) {
    try {
      New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
      Write-Log "Created output directory: $OutputPath" "INFO"
    } catch {
      Write-Log "Failed to create output directory: $_" "ERROR"
      $OutputPath = [Environment]::GetFolderPath("Desktop")
      Write-Log "Using desktop as fallback output location" "WARN"
    }
  }
  $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
  $excelPath = Join-Path -Path $OutputPath -ChildPath "DO_Report_$timestamp.xlsx"
  $csvFolder = Join-Path -Path $OutputPath -ChildPath "DO_Report_CSV_$timestamp"

  Write-Log "Exporting Excel report to $excelPath" "INFO"
  Show-Progress -Activity "Generating Excel report" -PercentComplete 95

  # Create Executive Summary
  $executiveSummary = New-ExecutiveSummary
  $executiveSummary | Export-Excel -Path $excelPath -WorksheetName "Summary" -AutoSize -TableStyle Medium9

  # Add all other sheets with enhanced formatting
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

  try {
    # Export DO error codes with enhanced information
    $errorTable = Get-DOErrorsTable
    $errorTable | Export-Excel -Path $excelPath -WorksheetName "Error Codes" -AutoSize -TableStyle Medium2 -Append
    Write-Log "Exported DO error code table with recommendations." "INFO"
  } catch {
    Write-Log "Failed to export DO error codes: $_" "ERROR"
  }

  try {
    New-Item -Path $csvFolder -ItemType Directory -Force | Out-Null
    foreach ($key in $Buffers.Keys) {
      if ($Buffers[$key].Count -gt 0) {
        $csvPath = Join-Path -Path $csvFolder -ChildPath "$key.csv"
        $Buffers[$key] | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
      }
    }
    Get-DOErrorsTable | Export-Csv -Path (Join-Path -Path $csvFolder -ChildPath "ErrorCodes.csv") -NoTypeInformation -Encoding UTF8
    Write-Log "CSV reports exported to $csvFolder for additional analysis" "INFO"
  } catch {
    Write-Log "CSV export failed: $_" "WARN"
  }

  if ($Show) { Invoke-Item $excelPath }
  Write-Host "`n📊 Report saved to: $excelPath" -ForegroundColor Cyan
  Write-Host "📋 CSV data exported to: $csvFolder" -ForegroundColor Cyan
  return $excelPath
}

# ───── EXTRACT DIAGNOSTICS ZIP ─────
function Extract-DiagnosticsZip {
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

# ───── PROCESS DIAGNOSTICS DATA ─────
function Process-DiagnosticsData {
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

    if ($doFiles.Count -eq 0) {
      Write-Log "No Delivery Optimization related files found in diagnostics package" "WARN"
      return
    }

    Write-Log "Found $($doFiles.Count) DO-related files in diagnostics package" "SUCCESS"

    # Process each file based on type
    foreach ($file in $doFiles) {
      $fileInfo = [PSCustomObject]@{
        FileName = $file.Name
        FilePath = $file.FullName
        FileSize = "{0:N2} KB" -f ($file.Length / 1KB)
        LastWriteTime = $file.LastWriteTime
        Category = "Unknown"
      }

      # Categorize the file
      if ($file.Name -match "\.etl$") {
        $fileInfo.Category = "ETL Log"
        # Process ETL files if needed
      }
      elseif ($file.Name -match "\.log$|\.txt$") {
        $fileInfo.Category = "Text Log"
        
        # Sample the first few lines to add context
        try {
          $sampleContent = Get-Content -Path $file.FullName -TotalCount 10 -ErrorAction Stop
          $relevantEntries = $sampleContent | Where-Object { $_ -match "DeliveryOptimization|BITS|WindowsUpdate|error|warning|fail" }
          
          if ($relevantEntries) {
            $fileInfo | Add-Member -MemberType NoteProperty -Name "SampleContent" -Value ($relevantEntries -join "`n")
          }
        }
        catch {
          Write-Log "Could not read content from $($file.Name): $_" "WARN"
        }
      }
      elseif ($file.Name -match "\.xml$|\.json$") {
        $fileInfo.Category = "Configuration"
        # Process configuration files if needed
      }
      elseif ($file.Name -match "\.cab$|\.zip$") {
        $fileInfo.Category = "Archive"
        # Process nested archives if needed
      }
      
      # Add to diagnostics buffer
      $Buffers.DiagnosticsData.Add($fileInfo) | Out-Null
    }

    # Look for Windows Update ETL logs that can be converted
    $wuEtlFiles = Get-ChildItem -Path $DiagnosticsPath -Recurse -File | Where-Object {
      $_.Name -match "\.etl$" -and $_.FullName -match "\\WindowsUpdate\\"
    }

    if ($wuEtlFiles.Count -gt 0) {
      Write-Log "Found $($wuEtlFiles.Count) Windows Update ETL logs" "INFO"
      
      # Create a temporary folder for converted logs
      $tempWULogFolder = Join-Path -Path $env:TEMP -ChildPath "WULogs_$([Guid]::NewGuid().ToString())"
      New-Item -Path $tempWULogFolder -ItemType Directory -Force | Out-Null
      
      # Try to convert the ETL files to text logs if Get-WindowsUpdateLog cmdlet is available
      if (Get-Command -Name Get-WindowsUpdateLog -ErrorAction SilentlyContinue) {
        $etlFolder = Split-Path -Path $wuEtlFiles[0].FullName -Parent
        $outputLog = Join-Path -Path $tempWULogFolder -ChildPath "WindowsUpdate.log"
        
        try {
          Get-WindowsUpdateLog -EtlPath $etlFolder -LogPath $outputLog -ErrorAction Stop
          Write-Log "Successfully converted Windows Update ETL logs to $outputLog" "SUCCESS"
          
          # Process converted log
          $wuLogEntries = Get-Content -Path $outputLog | Where-Object { 
            $_ -match "DeliveryOptimization|DO_|BITS" 
          }
          
          if ($wuLogEntries) {
            Write-Log "Found $($wuLogEntries.Count) DO-related entries in Windows Update logs" "INFO"
            
            foreach ($entry in $wuLogEntries) {
              $Buffers.DiagnosticsData.Add([PSCustomObject]@{
                FileName = "WindowsUpdate.log"
                Category = "Windows Update Log"
                Content = $entry
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
              }) | Out-Null
            }
          }
        }
        catch {
          Write-Log "Failed to convert Windows Update ETL logs: $_" "WARN"
        }
      }
      else {
        Write-Log "Get-WindowsUpdateLog cmdlet not available - cannot convert ETL files" "WARN"
      }
    }

    # Add summary of diagnostics data to Summary buffer
    $Buffers.Summary.Add([PSCustomObject]@{
      Test = "Diagnostics Data"
      Result = "$($doFiles.Count) DO-related files found"
      Status = if ($doFiles.Count -gt 0) { "PASS" } else { "WARN" }
      Impact = if ($doFiles.Count -gt 0) { "Additional diagnostic information available" } else { "Limited diagnostic information" }
    }) | Out-Null
  }
  catch {
    Write-Log "Error processing diagnostics data: $_" "ERROR"
  }
}

# ───── MAIN EXECUTION ─────
$startTime = Get-Date
Write-Host "`n🟦 Starting Delivery Optimization Troubleshooting..." -ForegroundColor Cyan
Write-Host "    System: $env:COMPUTERNAME | User: $env:USERNAME | Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor DarkCyan
Write-Host "    Output Path: $OutputPath" -ForegroundColor DarkCyan
Write-Host "    " -NoNewline
Write-Host "▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬" -ForegroundColor Blue

if (-not (Test-Path $OutputPath)) {
  New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

# Extract and process diagnostics data if provided
if ($DiagnosticsZip -and (Test-Path -Path $DiagnosticsZip)) {
  Write-Log "Diagnostics zip file provided: $DiagnosticsZip" "INFO"
  
  # Create a unique temporary folder for extraction
  $extractPath = Join-Path -Path $env:TEMP -ChildPath "DODiagnostics_$([Guid]::NewGuid().ToString())"
  
  # Extract the zip file
  $extractSuccess = Extract-DiagnosticsZip -ZipPath $DiagnosticsZip -ExtractPath $extractPath
  
  if ($extractSuccess) {
    # Process the extracted data
    Process-DiagnosticsData -DiagnosticsPath $extractPath
  }
}
else {
  if ($DiagnosticsZip) {
    Write-Log "Specified DiagnosticsZip file not found: $DiagnosticsZip" "WARN"
  }
}

# Run all diagnostics sequentially
Get-SystemInfo
Test-DOEndpoints
Test-DNSDPeerConfig
Test-DOGroupId
Start-HealthCheck
Test-MCCCheck
Test-PeerConnectivity
Invoke-DOTroubleshooter
Invoke-DOLogAnalysis
Test-TeamsImpact

# Generate final report
$reportPath = Export-DOExcelReport

$endTime = Get-Date
$duration = $endTime - $startTime
$minutes = [math]::Floor($duration.TotalMinutes)
$seconds = $duration.Seconds

Write-Host "`n✅ Delivery Optimization Troubleshooting completed in $minutes min $seconds sec" -ForegroundColor Green

# Get overall status for final output
$criticalIssues = ($Buffers.Summary | Where-Object { $_.Status -eq "FAIL" -or $_.Status -eq "ERROR" -or $_.Status -eq "CRITICAL" }).Count
$warnings = ($Buffers.Summary | Where-Object { $_.Status -eq "WARN" }).Count

if ($criticalIssues -gt 0) {
  Write-Host "⚠️ Found $criticalIssues critical issues that require attention" -ForegroundColor Red
  Write-Host "   Review the 'Recommendations' sheet for remediation steps" -ForegroundColor Red
} elseif ($warnings -gt 0) {
  Write-Host "⚠️ Found $warnings warnings that may need attention" -ForegroundColor Yellow
  Write-Host "   Review the 'Recommendations' sheet for potential improvements" -ForegroundColor Yellow
} else {
  Write-Host "👍 No issues detected. Delivery Optimization appears to be configured correctly" -ForegroundColor Green
}

Write-Host "`n📊 Report saved to: $reportPath" -ForegroundColor Cyan

if ($Show) {
  Write-Host "`n🔍 Opening Excel report..." -ForegroundColor Cyan
}

Write-Host "`n🟦 Delivery Optimization Troubleshooter Script Complete 🟦" -ForegroundColor Cyan
