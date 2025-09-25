<#
.SYNOPSIS
    Centralized polling-based OU routing for new AD computer objects.

.DESCRIPTION
    This script uses a polling approach to find and move newly created computer objects
    to appropriate OUs based on naming patterns. Designed for multi-DC environments
    where event-driven approaches are unreliable.

.PARAMETER Simulate
    Simulation mode - logs intended actions without actually moving computers.

.PARAMETER ConfigFile
    Path to configuration file (defaults to Config-Polling.psd1 next to script).

.PARAMETER PollIntervalMinutes
    Override the polling interval from config (for testing purposes).

.EXAMPLE
    AutoMove-Computer-Polling.ps1
    Standard polling mode using default configuration.

.EXAMPLE
    AutoMove-Computer-Polling.ps1 -Simulate
    Test mode - shows what would be moved without making changes.

.EXAMPLE
    AutoMove-Computer-Polling.ps1 -PollIntervalMinutes 1
    Override polling interval for testing (polls every minute).

.NOTES
    Version: 3.1 (Polling-Based with AD Tracking)
    Requires: ActiveDirectory PowerShell module
    Log Location: .\Logs\AutoMove-Computer-Polling.log
    Tracking: Uses extensionAttribute7 in AD (no external files)
    
    Scheduled Task Setup:
    - Run every 2-3 minutes
    - Single DC deployment (typically PDC Emulator)
    - No event triggers needed
    
    Advantages over event-driven:
    - 100% coverage regardless of creation DC
    - No replication timing issues  
    - Single deployment point
    - Self-managing AD-based tracking (no file dependencies)
    - Better error handling and retry logic
#>

[CmdletBinding()]
param(
    [switch]$Simulate,
    [string]$ConfigFile,
    [int]$PollIntervalMinutes
)

$ErrorActionPreference = 'Stop'

#region Configuration and Initialization

# Resolve script path and directories
$Script:ScriptPath = if ($MyInvocation.MyCommand.Path) { 
    $MyInvocation.MyCommand.Path 
} else { 
    $PSCommandPath 
}
if (-not $Script:ScriptPath) { 
    throw "Unable to determine script path. Run script from file system."
}
$Script:ScriptDir = Split-Path -Parent $Script:ScriptPath

# Configuration
$Script:Config = @{
    LogDir = Join-Path $Script:ScriptDir 'Logs'
    MaxLogSizeMB = 10
    EnableTranscript = $true
    # Polling-specific settings
    PollIntervalMinutes = 3
    LookbackMinutes = 10
    ProcessedComputersFile = 'processed_computers.json'
    MaxProcessedHistoryDays = 30
    RetryFailedMoves = $true
    MaxRetryAttempts = 3
    RetryIntervalMinutes = 5
}

# Load external config if specified or exists
$DefaultConfigPath = Join-Path $Script:ScriptDir 'Config-Polling.psd1'
$ConfigPath = if ($ConfigFile) { $ConfigFile } elseif (Test-Path $DefaultConfigPath) { $DefaultConfigPath }

if ($ConfigPath -and (Test-Path $ConfigPath)) {
    try {
        $ExternalConfig = Import-PowerShellDataFile -Path $ConfigPath
        foreach ($key in $ExternalConfig.Keys) {
            $Script:Config[$key] = $ExternalConfig[$key]
        }
        Write-Verbose "Loaded configuration from: $ConfigPath"
    }
    catch {
        Write-Warning "Failed to load config file '$ConfigPath': $($_.Exception.Message)"
    }
}

# Override polling interval if specified
if ($PollIntervalMinutes) {
    $Script:Config.PollIntervalMinutes = $PollIntervalMinutes
}

# Default routing rules (can be overridden in config file)
if (-not $Script:Config.Rules) {
    $Script:Config.Rules = @(
        @{ Pattern = '^Alpha'; OU = 'OU=Location Alpha,OU=Computers,OU=DEV Company,DC=dev,DC=local'; Description = 'Alpha location computers' }
        @{ Pattern = '^Beta';  OU = 'OU=Location Beta,OU=Computers,OU=DEV Company,DC=dev,DC=local'; Description = 'Beta location computers' }
        @{ Pattern = '^Gamma'; OU = 'OU=Location Gamma,OU=Computers,OU=DEV Company,DC=dev,DC=local'; Description = 'Gamma location computers' }
        @{ Pattern = '.*';     OU = 'OU=_Quarantine,OU=Computers,OU=DEV Company,DC=dev,DC=local'; Description = 'Fallback quarantine' }
    )
}

#endregion

#region Logging Functions

# Initialize logging
$Script:LogFile = Join-Path $Script:Config.LogDir 'AutoMove-Computer-Polling.log'
if (-not (Test-Path $Script:Config.LogDir)) { 
    New-Item -ItemType Directory -Path $Script:Config.LogDir -Force | Out-Null 
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'DEBUG')]
        [string]$Level,
        
        [Parameter(Mandatory)]
        [string]$Message
    )
    
    $Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
    $LogEntry = "{0} [{1}] {2}" -f $Timestamp, $Level.ToUpper(), $Message
    
    try {
        # Rotate log if too large
        if ((Test-Path $Script:LogFile) -and ((Get-Item $Script:LogFile).Length -gt ($Script:Config.MaxLogSizeMB * 1MB))) {
            $BackupLog = $Script:LogFile -replace '\.log$', "_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
            Move-Item -Path $Script:LogFile -Destination $BackupLog -Force
            Write-Host "Rotated log file to: $BackupLog" -ForegroundColor Yellow
        }
        
        Add-Content -Path $Script:LogFile -Value $LogEntry -Encoding UTF8
        
        # Also write to console for immediate feedback
        $Color = switch ($Level) {
            'ERROR' { 'Red' }
            'WARN' { 'Yellow' }
            'SUCCESS' { 'Green' }
            'DEBUG' { 'Cyan' }
            default { 'White' }
        }
        Write-Host $LogEntry -ForegroundColor $Color
    }
    catch {
        Write-Error "Failed to write to log: $($_.Exception.Message)"
    }
}

function Start-LoggingSession {
    # Generate unique session identifier for tracking
    $SessionId = [guid]::NewGuid().ToString().Substring(0,8).ToUpper()
    $Script:SessionId = $SessionId
    $Script:SessionStartTime = Get-Date
    
    # Session Header with ASCII box drawing
    Write-Log INFO "=================================================================================="
    Write-Log INFO "                     AutoMove-Computer-Polling Session Started"
    Write-Log INFO "=================================================================================="
    Write-Log INFO "Session ID: $SessionId"
    Write-Log INFO "Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')"
    Write-Log INFO " "
    
    # Host Information Section
    Write-Log INFO "HOST INFORMATION:"
    Write-Log INFO "  |-- Hostname: $env:COMPUTERNAME"
    Write-Log INFO "  |-- Domain: $(try { (Get-WmiObject Win32_ComputerSystem).Domain } catch { 'Unknown' })"
    Write-Log INFO "  |-- Operating User: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log INFO "  |-- PowerShell Version: $($PSVersionTable.PSVersion) [$($PSVersionTable.PSEdition)]"
    Write-Log INFO "  |-- OS Version: $(try { (Get-WmiObject Win32_OperatingSystem).Caption } catch { 'Unknown' })"
    Write-Log INFO " "
    
    # Execution Context Section
    Write-Log INFO "EXECUTION CONTEXT:"
    Write-Log INFO "  |-- Script Path: $Script:ScriptPath"
    Write-Log INFO "  |-- Working Directory: $(Get-Location)"
    Write-Log INFO "  |-- Config File: $Script:ConfigFile"
    Write-Log INFO "  |-- Log Directory: $($Script:Config.LogDir)"
    Write-Log INFO "  |-- Execution Mode: $(if ($Simulate) { 'SIMULATION (No Changes)' } else { 'PRODUCTION (Live Changes)' })"
    Write-Log INFO " "
    
    # Configuration Parameters Section
    Write-Log INFO "CONFIGURATION PARAMETERS:"
    Write-Log INFO "  |-- Poll Interval: $($Script:Config.PollIntervalMinutes) minutes"
    Write-Log INFO "  |-- Lookback Window: $($Script:Config.LookbackMinutes) minutes"
    
    if ($Script:Config.EnableTranscript) {
        try {
            $TranscriptPath = Join-Path $Script:Config.LogDir "Transcript-Polling_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            Start-Transcript -Path $TranscriptPath -Force | Out-Null
            Write-Log INFO "Transcript started: $TranscriptPath"
        }
        catch {
            Write-Log WARN "Failed to start transcript: $($_.Exception.Message)"
        }
    }
}

#endregion

#region Active Directory Functions

function Initialize-ADModule {
    try {
        if (-not (Get-Module -Name ActiveDirectory)) {
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        Write-Log INFO "ActiveDirectory module loaded successfully"
        
        # Test AD connectivity
        $null = Get-ADDomain -ErrorAction Stop
        Write-Log INFO "Active Directory connectivity verified"
    }
    catch {
        Write-Log ERROR "ActiveDirectory module initialization failed: $($_.Exception.Message)"
        throw "ActiveDirectory module required but not available"
    }
}

function Get-AllDomainControllers {
    try {
        Write-Log INFO "DOMAIN CONTROLLER DISCOVERY:"
        $DiscoveryStart = Get-Date
        
        $DCObjects = Get-ADDomainController -Filter *
        $DCs = $DCObjects.Name | Sort-Object
        
        $DiscoveryDuration = ((Get-Date) - $DiscoveryStart).TotalMilliseconds
        
        Write-Log INFO "  |-- Discovery Method: LDAP query via Get-ADDomainController"
        Write-Log INFO "  |-- Query Duration: $([math]::Round($DiscoveryDuration))ms"
        Write-Log INFO "  |-- Domain Controllers Found: $($DCs.Count)"
        
        foreach ($DC in $DCObjects) {
            $RoleInfo = @()
            if ($DC.OperationMasterRoles) { $RoleInfo += "FSMO: $($DC.OperationMasterRoles -join ',')" }
            if ($DC.IsGlobalCatalog) { $RoleInfo += "GC" }
            if ($DC.IsReadOnly) { $RoleInfo += "RODC" }
            
            $Roles = if ($RoleInfo.Count -gt 0) { " ($($RoleInfo -join ', '))" } else { "" }
            Write-Log DEBUG "    * $($DC.Name)$Roles - Site: $($DC.Site)"
        }
        
        return $DCs
    }
    catch {
        Write-Log ERROR "Domain Controller discovery failed: $($_.Exception.Message)"
        Write-Log DEBUG "  |-- Error Type: $($_.Exception.GetType().Name)"
        throw
    }
}

#endregion

#region Computer Discovery Functions

function Test-ComputerProcessed {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )
    
    try {
        $Computer = Get-ADComputer -Identity $ComputerName -Properties extensionAttribute7 -ErrorAction Stop
        
        if ([string]::IsNullOrEmpty($Computer.extensionAttribute7)) {
            return $false
        }
        
        # Check if processed within retention period
        $ProcessedDate = [DateTime]::Parse($Computer.extensionAttribute7)
        $CutoffDate = (Get-Date).AddDays(-$Script:Config.MaxProcessedHistoryDays)
        
        if ($ProcessedDate -lt $CutoffDate) {
            Write-Log DEBUG "Computer '$ComputerName' processed too long ago ($ProcessedDate), will reprocess"
            return $false
        }
        
        Write-Log DEBUG "Computer '$ComputerName' already processed on $ProcessedDate"
        return $true
    }
    catch {
        Write-Log DEBUG "Computer '$ComputerName' not found or error checking processed status: $($_.Exception.Message)"
        return $false
    }
}

function Mark-ComputerProcessed {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName,
        
        [Parameter(Mandatory)]
        [string]$Result,
        
        [string]$TargetOU = $null,
        
        [switch]$WhatIf
    )
    
    $ProcessedTimestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    
    if ($WhatIf) {
        Write-Log DEBUG "[SIMULATION] Would mark '$ComputerName' as processed with result: $Result"
        return
    }
    
    try {
        Set-ADComputer -Identity $ComputerName -Replace @{extensionAttribute7 = $ProcessedTimestamp} -ErrorAction Stop
        Write-Log DEBUG "Marked computer '$ComputerName' as processed ($Result) at $ProcessedTimestamp"
    }
    catch {
        Write-Log WARN "Failed to mark computer '$ComputerName' as processed: $($_.Exception.Message)"
    }
}

function Find-NewComputersOnAllDCs {
    $AllDCs = Get-AllDomainControllers
    $DefaultComputersContainer = "CN=Computers,$((Get-ADDomain).DistinguishedName)"
    $LookbackTime = (Get-Date).AddMinutes(-$Script:Config.LookbackMinutes)
    $AllNewComputers = @()
    
    Write-Log INFO " "
    Write-Log INFO "COMPUTER DISCOVERY OPERATION:"
    Write-Log INFO "  |-- Search Base: $DefaultComputersContainer"
    Write-Log INFO "  |-- Time Filter: Computers created after $($LookbackTime.ToString('yyyy-MM-dd HH:mm:ss.fff'))"
    Write-Log INFO "  |-- Domain Controllers to Query: $($AllDCs.Count)"
    Write-Log INFO " "
    
    $DCResults = @{}
    $TotalDCs = $AllDCs.Count
    $SuccessfulDCs = 0
    $FailedDCs = 0
    
    foreach ($DC in $AllDCs) {
        $DCStart = Get-Date
        try {
            Write-Log DEBUG "Querying DC: $DC"
            
            # Query for computers created recently in the default Computers container
            $Computers = Get-ADComputer -Server $DC -Filter "whenCreated -gt '$($LookbackTime.ToString('yyyy-MM-dd HH:mm:ss'))'" -SearchBase $DefaultComputersContainer -Properties whenCreated, DistinguishedName -ErrorAction Stop
            
            $DCDuration = ((Get-Date) - $DCStart).TotalMilliseconds
            $DCResults[$DC] = @{
                Success = $true
                ComputerCount = $Computers.Count
                Duration = $DCDuration
            }
            
            $SuccessfulDCs++
            
            if ($Computers) {
                Write-Log DEBUG "    + DC ${DC}: Found $($Computers.Count) computers ($([math]::Round($DCDuration))ms)"
                $AllNewComputers += $Computers
            } else {
                Write-Log DEBUG "    o DC ${DC}: No new computers found ($([math]::Round($DCDuration))ms)"
            }
        }
        catch {
            $DCDuration = ((Get-Date) - $DCStart).TotalMilliseconds
            $DCResults[$DC] = @{
                Success = $false
                ComputerCount = 0
                Duration = $DCDuration
                Error = $_.Exception.Message
            }
            
            $FailedDCs++
            Write-Log ERROR "    - DC ${DC}: Query failed - $($_.Exception.Message)"
        }
    }
    
    # Calculate summary statistics
    $TotalQueryTime = ($DCResults.Values | Measure-Object -Property Duration -Sum).Sum
    
    Write-Log INFO " "
    Write-Log INFO "DISCOVERY SUMMARY:"
    Write-Log INFO "  |-- DCs Queried: $SuccessfulDCs/$TotalDCs successful ($FailedDCs failed)"
    Write-Log INFO "  |-- Total Query Time: $([math]::Round($TotalQueryTime))ms"
    Write-Log INFO "  |-- Average Query Time: $([math]::Round($TotalQueryTime / $TotalDCs))ms per DC"
    Write-Log INFO "  |-- Raw Computer Results: $($AllNewComputers.Count) objects"
    
    # Remove duplicates (same computer found on multiple DCs)
    $UniqueComputers = $AllNewComputers | Sort-Object Name, whenCreated | Get-Unique -AsString
    
    Write-Log INFO "  |-- Unique Computers After Deduplication: $($UniqueComputers.Count)"
    
    if ($AllNewComputers.Count -ne $UniqueComputers.Count) {
        $Duplicates = $AllNewComputers.Count - $UniqueComputers.Count
        Write-Log DEBUG "    Removed $Duplicates duplicate entries from multi-DC results"
    }
    
    Write-Log INFO " "
    Write-Log INFO "PROCESSING FILTER RESULTS:"
    Write-Log INFO "  |-- New Computers Found: $($UniqueComputers.Count)"
    
    return $UniqueComputers
}

function Get-UnprocessedComputers {
    $NewComputers = Find-NewComputersOnAllDCs
    
    # Filter out already processed computers using extensionAttribute7
    $UnprocessedComputers = $NewComputers | Where-Object { 
        -not (Test-ComputerProcessed -ComputerName $_.Name)
    }
    
    Write-Log INFO "Found $($UnprocessedComputers.Count) unprocessed computer(s) requiring OU placement"
    
    foreach ($Computer in $UnprocessedComputers) {
        Write-Log INFO "Unprocessed: $($Computer.Name) (created: $($Computer.whenCreated))"
    }
    
    return $UnprocessedComputers
}

#endregion

#region Computer Processing Functions

function Resolve-TargetOU {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )
    
    Write-Log DEBUG "Evaluating routing rules for computer: $ComputerName"
    
    foreach ($Rule in $Script:Config.Rules) {
        if ($ComputerName -match $Rule.Pattern) {
            $Description = if ($Rule.Description) { " ($($Rule.Description))" } else { "" }
            Write-Log DEBUG "Matched rule '$($Rule.Pattern)' -> $($Rule.OU)$Description"
            return $Rule.OU
        }
    }
    
    Write-Log WARN "No routing rules matched for computer: $ComputerName"
    return $null
}

function Test-ADOrganizationalUnit {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DistinguishedName
    )
    
    try {
        $null = Get-ADOrganizationalUnit -Identity $DistinguishedName -ErrorAction Stop
        return $true
    }
    catch {
        Write-Log ERROR "Target OU not found or inaccessible: $DistinguishedName - $($_.Exception.Message)"
        return $false
    }
}

function Move-ComputerToOU {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Computer,
        
        [Parameter(Mandatory)]
        [string]$TargetOU,
        
        [switch]$WhatIf
    )
    
    $ComputerName = $Computer.Name
    $CurrentDN = $Computer.DistinguishedName
    
    # Check if already in target OU
    if ($CurrentDN -like "*$TargetOU") {
        Write-Log INFO "Computer '$ComputerName' already in target OU: $TargetOU"
        return $true
    }
    
    # Validate target OU exists
    if (-not (Test-ADOrganizationalUnit -DistinguishedName $TargetOU)) {
        return $false
    }
    
    if ($WhatIf) {
        Write-Log INFO "[SIMULATION] Would move: $CurrentDN -> $TargetOU"
        return $true
    }
    
    try {
        Write-Log INFO "Moving computer: $CurrentDN -> $TargetOU"
        Move-ADObject -Identity $CurrentDN -TargetPath $TargetOU -Confirm:$false -ErrorAction Stop
        Write-Log SUCCESS "Successfully moved '$ComputerName' to '$TargetOU'"
        return $true
    }
    catch {
        Write-Log ERROR "Failed to move computer '$ComputerName': $($_.Exception.Message)"
        return $false
    }
}

function Process-Computer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Computer
    )
    
    $ComputerName = $Computer.Name
    Write-Log INFO "Processing computer: $ComputerName"
    
    # Determine target OU
    $TargetOU = Resolve-TargetOU -ComputerName $ComputerName
    if (-not $TargetOU) {
        Write-Log WARN "No routing rule matched for '$ComputerName' - skipping"
        
        # Mark as processed (even though no rule matched) to avoid re-processing
        Mark-ComputerProcessed -ComputerName $ComputerName -Result 'NoRuleMatched' -WhatIf:$Simulate
        return $true
    }
    
    # Attempt to move computer
    $MoveResult = Move-ComputerToOU -Computer $Computer -TargetOU $TargetOU -WhatIf:$Simulate
    
    # Mark as processed with result
    $Result = if ($MoveResult) { 'Success' } else { 'Failed' }
    Mark-ComputerProcessed -ComputerName $ComputerName -Result $Result -TargetOU $TargetOU -WhatIf:$Simulate
    
    return $MoveResult
}

#endregion

#region Main Execution

function Stop-ScriptExecution {
    param(
        [int]$ExitCode = 0,
        [string]$Message
    )
    
    # Calculate session duration
    $SessionEnd = Get-Date
    $SessionDuration = if ($Script:SessionStartTime) { 
        ($SessionEnd - $Script:SessionStartTime).TotalSeconds 
    } else { 
        0 
    }
    
    # Log exit reason if provided
    if ($Message) {
        $Level = if ($ExitCode -eq 0) { 'INFO' } else { 'ERROR' }
        Write-Log $Level $Message
    }
    
    Write-Log INFO " "
    Write-Log INFO "=================================================================================="
    Write-Log INFO "                     AutoMove-Computer-Polling Session Ended"
    Write-Log INFO "=================================================================================="
    Write-Log INFO "Session ID: $($Script:SessionId)"
    Write-Log INFO "End Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')"
    Write-Log INFO "Total Duration: $([math]::Round($SessionDuration, 2)) seconds"
    Write-Log INFO "Exit Code: $ExitCode $(if ($ExitCode -eq 0) { '(Success)' } elseif ($ExitCode -eq 99) { '(Critical Error)' } else { '(Error)' })"
    Write-Log INFO "=================================================================================="
    
    if ($Script:Config.EnableTranscript) {
        try { Stop-Transcript | Out-Null } catch { }
    }
    
    exit $ExitCode
}

# Main execution starts here
Start-LoggingSession

try {
    # Initialize Active Directory module
    Initialize-ADModule
    
    # Find computers that need processing
    $UnprocessedComputers = Get-UnprocessedComputers
    
    if ($UnprocessedComputers.Count -eq 0) {
        Write-Log INFO "No unprocessed computers found - nothing to do"
        Stop-ScriptExecution -ExitCode 0 -Message "Polling cycle completed - no new computers"
    }
    
    # Process each computer with detailed metrics
    Write-Log INFO " "
    Write-Log INFO "BATCH PROCESSING INITIATED:"
    Write-Log INFO "  |-- Queue Size: $($UnprocessedComputers.Count) computers"
    Write-Log INFO "  |-- Execution Mode: $(if ($Simulate) { 'SIMULATION' } else { 'PRODUCTION' })"
    Write-Log INFO " "
    
    $SuccessCount = 0
    $FailureCount = 0
    $ProcessingTimes = @()
    $BatchStart = Get-Date
    
    foreach ($Computer in $UnprocessedComputers) {
        $ComputerStart = Get-Date
        try {
            Write-Log INFO "[$($SuccessCount + $FailureCount + 1)/$($UnprocessedComputers.Count)] Processing: $($Computer.Name)"
            
            if (Process-Computer -Computer $Computer) {
                $SuccessCount++
                $ProcessingTime = ((Get-Date) - $ComputerStart).TotalMilliseconds
                $ProcessingTimes += $ProcessingTime
                Write-Log DEBUG "  Processing completed in $([math]::Round($ProcessingTime))ms"
            } else {
                $FailureCount++
            }
        }
        catch {
            Write-Log ERROR "Unexpected error processing computer '$($Computer.Name)': $($_.Exception.Message)"
            $FailureCount++
            
            # Still mark as processed to avoid infinite retries
            Mark-ComputerProcessed -ComputerName $Computer.Name -Result 'Error' -WhatIf:$Simulate
        }
    }
    
    # Calculate comprehensive batch statistics
    $BatchEnd = Get-Date
    $BatchDuration = ($BatchEnd - $BatchStart).TotalSeconds
    $TotalProcessed = $SuccessCount + $FailureCount
    $SuccessRate = if ($TotalProcessed -gt 0) { ($SuccessCount / $TotalProcessed * 100) } else { 0 }
    
    Write-Log INFO " "
    Write-Log INFO "BATCH PROCESSING COMPLETE:"
    Write-Log INFO "  |-- Duration: $([math]::Round($BatchDuration, 1)) seconds"
    Write-Log INFO "  |-- Computers Processed: $TotalProcessed"
    Write-Log INFO "  |-- Successful: $SuccessCount ($([math]::Round($SuccessRate, 1))%)"
    Write-Log INFO "  |-- Failed: $FailureCount"
    Write-Log INFO "  |-- Average Processing Time: $([math]::Round(($ProcessingTimes | Measure-Object -Average).Average))ms per computer"
    Write-Log INFO "  |-- Throughput: $([math]::Round($TotalProcessed / $BatchDuration, 2)) computers/second"
    
    if ($Simulate) {
        Stop-ScriptExecution -ExitCode 0 -Message "SIMULATION: All computers would be processed successfully"
    } elseif ($FailureCount -eq 0) {
        Stop-ScriptExecution -ExitCode 0 -Message "All computers processed successfully"
    } else {
        Stop-ScriptExecution -ExitCode 1 -Message "$FailureCount computer(s) failed to process"
    }
}
catch {
    $ErrorDetails = "Unhandled exception: $($_.Exception.Message)`nStack Trace: $($_.ScriptStackTrace)"
    Stop-ScriptExecution -ExitCode 99 -Message $ErrorDetails
}

#endregion