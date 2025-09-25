<#
.SYNOPSIS
    Event-driven OU routing for new AD computer objects (Security Event 5137).

.DESCRIPTION
    This script automatically moves newly created computer objects to appropriate OUs based on naming patterns.
    It responds to Security Event 5137 (object creation) and applies configurable routing rules.

.PARAMETER RecordId
    Event Record ID from Task Scheduler macro "$(EventRecordID)" or manual numeric value.

.PARAMETER ComputerName  
    Manual computer name override for testing purposes (bypasses event lookup).

.PARAMETER Simulate
    Simulation mode - logs intended actions without actually moving computers.

.PARAMETER ConfigFile
    Path to configuration file (defaults to Config.psd1 next to script).

.EXAMPLE
    AutoMove-Computer.ps1 -RecordId "$(EventRecordID)"
    Standard usage from Task Scheduler.

.EXAMPLE
    AutoMove-Computer.ps1 -ComputerName "Alpha-PC001" -Simulate
    Test mode with specific computer name.

.NOTES
    Version: 2.0
    Requires: ActiveDirectory PowerShell module
    Log Location: .\Logs\AutoMove-Computer.log
    
    Task Scheduler Action:
    Program:   C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
    Arguments: -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\AutoMove-Computer.ps1" -RecordId "$(EventRecordID)"
    Start in:  C:\Scripts
#>

[CmdletBinding()]
param(
    [Parameter(ParameterSetName='Event')]
    [string]$RecordId,

    [Parameter(ParameterSetName='Manual')]
    [string]$ComputerName,

    [switch]$Simulate,

    [string]$ConfigFile
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
    ReplicationTimeoutSeconds = 60
    ReplicationRetryIntervalSeconds = 5
    EventLookupMinutes = 5
    MaxEvents = 500
    EnableTranscript = $true
    # Multi-DC Configuration
    EnableMultiDCFallback = $true
    DCLookupTimeoutSeconds = 15
    QueryOriginatingDCFirst = $true
    FallbackToAllDCs = $true
}

# Load external config if specified or exists
$DefaultConfigPath = Join-Path $Script:ScriptDir 'Config.psd1'
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
$Script:LogFile = Join-Path $Script:Config.LogDir 'AutoMove-Computer.log'
if (-not (Test-Path $Script:Config.LogDir)) { 
    New-Item -ItemType Directory -Path $Script:Config.LogDir -Force | Out-Null 
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes timestamped log entries to the log file.
    #>
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
    <#
    .SYNOPSIS
        Initializes the logging session with system information.
    #>
    Write-Log INFO "----- AutoMove-Computer Session Started -----"
    Write-Log INFO "Host: $env:COMPUTERNAME"
    Write-Log INFO "User: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log INFO "PowerShell: $($PSVersionTable.PSVersion)"
    Write-Log INFO "Script: $Script:ScriptPath"
    Write-Log INFO "Parameters: RecordId='$RecordId' ComputerName='$ComputerName' Simulate=$([bool]$Simulate)"
    
    if ($Script:Config.EnableTranscript) {
        try {
            $TranscriptPath = Join-Path $Script:Config.LogDir "Transcript_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            Start-Transcript -Path $TranscriptPath -Force | Out-Null
            Write-Log INFO "Transcript started: $TranscriptPath"
        }
        catch {
            Write-Log WARN "Failed to start transcript: $($_.Exception.Message)"
        }
    }
}

#endregion

#region Active Directory Module

function Initialize-ADModule {
    <#
    .SYNOPSIS
        Imports and validates the Active Directory PowerShell module.
    #>
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

#endregion

#region Event Processing Functions

function Get-SecurityEvent5137ByRecordId {
    <#
    .SYNOPSIS
        Retrieves Security Event 5137 by Record ID.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [long]$RecordId
    )
    
    try {
        $Event = Get-WinEvent -LogName Security -MaxEvents $Script:Config.MaxEvents |
            Where-Object { $_.Id -eq 5137 -and $_.RecordId -eq $RecordId } |
            Select-Object -First 1
            
        if ($Event) {
            Write-Log INFO "Found Security Event 5137 with RecordId: $RecordId"
            return $Event
        }
        else {
            Write-Log WARN "No Security Event 5137 found with RecordId: $RecordId"
            return $null
        }
    }
    catch {
        Write-Log ERROR "Failed to retrieve event by RecordId $RecordId : $($_.Exception.Message)"
        return $null
    }
}

function Get-RecentComputerCreationEvent {
    <#
    .SYNOPSIS
        Retrieves the most recent computer creation event (5137).
    #>
    [CmdletBinding()]
    param(
        [datetime]$Since = (Get-Date).AddMinutes(-$Script:Config.EventLookupMinutes)
    )
    
    try {
        Write-Log INFO "Searching for recent computer creation events since: $Since"
        
        $Events = Get-WinEvent -FilterHashtable @{
            LogName = 'Security'
            Id = 5137
            StartTime = $Since
        } -ErrorAction SilentlyContinue
        
        foreach ($Event in $Events) {
            try {
                $EventXml = [xml]$Event.ToXml()
                $ObjectClass = ($EventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'ObjectClass' }).'#text'
                
                if ($ObjectClass -eq 'computer') {
                    Write-Log INFO "Found recent computer creation event (RecordId: $($Event.RecordId))"
                    return $Event
                }
            }
            catch {
                Write-Log WARN "Failed to parse event RecordId $($Event.RecordId): $($_.Exception.Message)"
            }
        }
        
        Write-Log WARN "No recent computer creation events found since $Since"
        return $null
    }
    catch {
        Write-Log ERROR "Failed to search for recent events: $($_.Exception.Message)"
        return $null
    }
}

function ConvertFrom-SecurityEvent5137 {
    <#
    .SYNOPSIS
        Extracts computer information from Security Event 5137.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Diagnostics.Eventing.Reader.EventLogRecord]$Event
    )
    
    try {
        $EventXml = [xml]$Event.ToXml()
        $EventData = @{}
        
        foreach ($Data in $EventXml.Event.EventData.Data) {
            $EventData[$Data.Name] = [string]$Data.'#text'
        }
        
        if ($EventData['ObjectClass'] -ne 'computer') {
            throw "Event is not for a computer object (ObjectClass: $($EventData['ObjectClass']))"
        }
        
        $DistinguishedName = $EventData['ObjectDN']
        $ComputerName = ($DistinguishedName -split ',')[0] -replace '^CN=', ''
        $Creator = if ($EventData['SubjectDomainName'] -and $EventData['SubjectUserName']) {
            "$($EventData['SubjectDomainName'])\$($EventData['SubjectUserName'])"
        } else {
            'Unknown'
        }
        
        return @{
            ComputerName = $ComputerName
            DistinguishedName = $DistinguishedName
            Creator = $Creator
            EventData = $EventData
        }
    }
    catch {
        Write-Log ERROR "Failed to parse Security Event 5137: $($_.Exception.Message)"
        throw
    }
}

#endregion

#region Computer Processing Functions

function Resolve-TargetOU {
    <#
    .SYNOPSIS
        Determines the target OU for a computer based on routing rules.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )
    
    Write-Log INFO "Evaluating routing rules for computer: $ComputerName"
    
    foreach ($Rule in $Script:Config.Rules) {
        if ($ComputerName -match $Rule.Pattern) {
            $Description = if ($Rule.Description) { " ($($Rule.Description))" } else { "" }
            Write-Log INFO "Matched rule '$($Rule.Pattern)' -> $($Rule.OU)$Description"
            return $Rule.OU
        }
    }
    
    Write-Log WARN "No routing rules matched for computer: $ComputerName"
    return $null
}

function Wait-ForADReplication {
    <#
    .SYNOPSIS
        Waits for computer object to become visible in AD with multi-DC support.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName,
        
        [int]$TimeoutSeconds = $Script:Config.ReplicationTimeoutSeconds,
        [int]$RetryIntervalSeconds = $Script:Config.ReplicationRetryIntervalSeconds
    )
    
    Write-Log INFO "Searching for computer with multi-DC support: $ComputerName (timeout: ${TimeoutSeconds}s)"
    
    $MaxRetries = [math]::Ceiling($TimeoutSeconds / $RetryIntervalSeconds)
    
    # First try current DC multiple times
    for ($Attempt = 1; $Attempt -le [math]::Min($MaxRetries, 3); $Attempt++) {
        try {
            $Computer = Get-ADComputer -Identity $ComputerName -Properties DistinguishedName, whenCreated -ErrorAction Stop
            Write-Log SUCCESS "Computer '$ComputerName' found on current DC (attempt $Attempt)"
            return $Computer
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            Write-Log DEBUG "Attempt $Attempt of 3: Computer not found on current DC, waiting ${RetryIntervalSeconds}s..."
            if ($Attempt -lt 3) { Start-Sleep -Seconds $RetryIntervalSeconds }
        }
    }
    
    # If multi-DC fallback is enabled, try other DCs
    if ($Script:Config.EnableMultiDCFallback) {
        Write-Log INFO "Current DC search failed, trying multi-DC fallback"
        
        try {
            # Get list of available domain controllers
            $DCs = Get-ADDomainController -Filter * | Select-Object -ExpandProperty Name
            Write-Log INFO "Found $($DCs.Count) domain controllers for fallback search"
            
            foreach ($DC in $DCs) {
                # Skip current machine to avoid duplicate search
                if ($DC -eq $env:COMPUTERNAME) { continue }
                
                try {
                    Write-Log DEBUG "Trying DC: $DC"
                    $Computer = Get-ADComputer -Identity $ComputerName -Properties DistinguishedName, whenCreated -Server $DC -ErrorAction Stop
                    Write-Log SUCCESS "Computer '$ComputerName' found on DC '$DC' via multi-DC fallback"
                    return $Computer
                }
                catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                    Write-Log DEBUG "Computer not found on DC '$DC'"
                }
                catch {
                    Write-Log WARN "Error querying DC '$DC': $($_.Exception.Message)"
                }
            }
        }
        catch {
            Write-Log WARN "Failed to enumerate domain controllers: $($_.Exception.Message)"
        }
    }
    
    Write-Log ERROR "Computer '$ComputerName' not found on any available domain controller"
    throw "Computer not found in multi-DC environment: $ComputerName"
}

function Test-ADOrganizationalUnit {
    <#
    .SYNOPSIS
        Validates that the target OU exists in Active Directory.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DistinguishedName
    )
    
    try {
        $null = Get-ADOrganizationalUnit -Identity $DistinguishedName -ErrorAction Stop
        Write-Log INFO "Target OU validated: $DistinguishedName"
        return $true
    }
    catch {
        Write-Log ERROR "Target OU not found or inaccessible: $DistinguishedName - $($_.Exception.Message)"
        return $false
    }
}

function Move-ComputerToOU {
    <#
    .SYNOPSIS
        Moves a computer object to the specified OU.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [Microsoft.ActiveDirectory.Management.ADComputer]$Computer,
        
        [Parameter(Mandatory)]
        [string]$TargetOU,
        
        [string]$Creator = 'Unknown',
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
        Write-Log SUCCESS "Successfully moved '$ComputerName' to '$TargetOU' (Creator: $Creator)"
        return $true
    }
    catch {
        Write-Log ERROR "Failed to move computer '$ComputerName': $($_.Exception.Message)"
        return $false
    }
}

#endregion

#region Main Execution

function Stop-ScriptExecution {
    <#
    .SYNOPSIS
        Cleanly stops script execution with proper cleanup.
    #>
    param(
        [int]$ExitCode = 0,
        [string]$Message
    )
    
    if ($Message) {
        $Level = if ($ExitCode -eq 0) { 'INFO' } else { 'ERROR' }
        Write-Log $Level $Message
    }
    
    Write-Log INFO "----- AutoMove-Computer Session Ended (Exit Code: $ExitCode) -----"
    
    if ($Script:Config.EnableTranscript) {
        try { Stop-Transcript | Out-Null } catch { }
    }
    
    exit $ExitCode
}

# Initialize logging and environment
Start-LoggingSession

try {
    # Initialize Active Directory module (skip in simulate mode for testing)
    if (-not $Simulate) {
        Initialize-ADModule
    } else {
        Write-Log INFO "SIMULATION MODE: Skipping ActiveDirectory module initialization"
    }
    
    # Determine target computer
    $Creator = 'Unknown'
    $EventInfo = $null
    
    if ($ComputerName) {
        Write-Log INFO "Manual override mode: Using ComputerName='$ComputerName'"
    }
    else {
        # Parse RecordId if provided
        $RecordIdParsed = $null
        if ($RecordId -and $RecordId -match '^\d+$') {
            $RecordIdParsed = [long]$RecordId
        }
        
        # Try to get event by RecordId first
        $Event = $null
        if ($RecordIdParsed) {
            $Event = Get-SecurityEvent5137ByRecordId -RecordId $RecordIdParsed
        }
        
        # Fallback to recent event search
        if (-not $Event) {
            Write-Log WARN "No event found by RecordId '$RecordId', searching for recent computer creation events"
            $Event = Get-RecentComputerCreationEvent
        }
        
        if (-not $Event) {
            Stop-ScriptExecution -ExitCode 20 -Message "No suitable Security Event 5137 found for computer creation"
        }
        
        # Extract computer information from event
        try {
            $EventInfo = ConvertFrom-SecurityEvent5137 -Event $Event
            $ComputerName = $EventInfo.ComputerName
            $Creator = $EventInfo.Creator
            
            Write-Log INFO "Computer from event: Name='$ComputerName', DN='$($EventInfo.DistinguishedName)', Creator='$Creator'"
        }
        catch {
            Stop-ScriptExecution -ExitCode 21 -Message "Failed to extract computer information from event: $($_.Exception.Message)"
        }
    }
    
    # In simulate mode, skip AD operations and just test routing logic
    if ($Simulate) {
        Write-Log INFO "SIMULATION MODE: Skipping Active Directory operations"
        
        # Determine target OU using routing rules only
        $TargetOU = Resolve-TargetOU -ComputerName $ComputerName
        if (-not $TargetOU) {
            Stop-ScriptExecution -ExitCode 0 -Message "SIMULATION: No routing rule matched for '$ComputerName' - no action would be taken"
        }
        
        Write-Log INFO "SIMULATION: Computer '$ComputerName' would be moved to '$TargetOU' (Creator: $Creator)"
        Stop-ScriptExecution -ExitCode 0 -Message "SIMULATION: Computer '$ComputerName' routing test completed successfully"
    }
    
    # Wait for AD replication (only in non-simulate mode)
    try {
        $ADComputer = Wait-ForADReplication -ComputerName $ComputerName
    }
    catch {
        Stop-ScriptExecution -ExitCode 30 -Message "Computer '$ComputerName' not found in Active Directory: $($_.Exception.Message)"
    }
    
    # Determine target OU
    $TargetOU = Resolve-TargetOU -ComputerName $ComputerName
    if (-not $TargetOU) {
        Stop-ScriptExecution -ExitCode 0 -Message "No routing rule matched for '$ComputerName' - no action taken"
    }
    
    # Move computer to target OU (this will be WhatIf since we're not in simulate mode here)
    $MoveResult = Move-ComputerToOU -Computer $ADComputer -TargetOU $TargetOU -Creator $Creator -WhatIf:$false
    
    if ($MoveResult) {
        Stop-ScriptExecution -ExitCode 0 -Message "Moved computer '$ComputerName' successfully"
    }
    else {
        Stop-ScriptExecution -ExitCode 50 -Message "Failed to move computer '$ComputerName'"
    }
}
catch {
    $ErrorDetails = "Unhandled exception: $($_.Exception.Message)`nStack Trace: $($_.ScriptStackTrace)"
    Stop-ScriptExecution -ExitCode 99 -Message $ErrorDetails
}

#endregion
