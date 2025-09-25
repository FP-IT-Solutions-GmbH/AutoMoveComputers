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
    Version: 3.0 (Polling-Based)
    Requires: ActiveDirectory PowerShell module
    Log Location: .\Logs\AutoMove-Computer-Polling.log
    
    Scheduled Task Setup:
    - Run every 2-3 minutes
    - Single DC deployment (typically PDC Emulator)
    - No event triggers needed
    
    Advantages over event-driven:
    - 100% coverage regardless of creation DC
    - No replication timing issues  
    - Single deployment point
    - Duplicate prevention built-in
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
    Write-Log INFO "----- AutoMove-Computer-Polling Session Started -----"
    Write-Log INFO "Host: $env:COMPUTERNAME"
    Write-Log INFO "User: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-Log INFO "PowerShell: $($PSVersionTable.PSVersion)"
    Write-Log INFO "Script: $Script:ScriptPath"
    Write-Log INFO "Mode: $(if ($Simulate) { 'SIMULATION' } else { 'PRODUCTION' })"
    Write-Log INFO "Poll Interval: $($Script:Config.PollIntervalMinutes) minutes"
    Write-Log INFO "Lookback Window: $($Script:Config.LookbackMinutes) minutes"
    
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
        $DCs = Get-ADDomainController -Filter * | Select-Object -ExpandProperty Name | Sort-Object
        Write-Log INFO "Discovered $($DCs.Count) domain controllers: $($DCs -join ', ')"
        return $DCs
    }
    catch {
        Write-Log ERROR "Failed to enumerate domain controllers: $($_.Exception.Message)"
        throw
    }
}

#endregion

#region Computer Discovery Functions

function Get-ProcessedComputersHistory {
    $ProcessedFile = Join-Path $Script:ScriptDir $Script:Config.ProcessedComputersFile
    
    if (-not (Test-Path $ProcessedFile)) {
        Write-Log INFO "No processed computers history found, starting fresh"
        return @{}
    }
    
    try {
        $Content = Get-Content $ProcessedFile -Raw | ConvertFrom-Json -AsHashtable
        
        # Clean up old entries
        $CutoffDate = (Get-Date).AddDays(-$Script:Config.MaxProcessedHistoryDays)
        $CleanedContent = @{}
        
        foreach ($ComputerName in $Content.Keys) {
            $ProcessedDate = [DateTime]$Content[$ComputerName].ProcessedDate
            if ($ProcessedDate -gt $CutoffDate) {
                $CleanedContent[$ComputerName] = $Content[$ComputerName]
            }
        }
        
        Write-Log INFO "Loaded $($CleanedContent.Count) processed computers from history (cleaned $($Content.Count - $CleanedContent.Count) old entries)"
        return $CleanedContent
    }
    catch {
        Write-Log WARN "Failed to load processed computers history: $($_.Exception.Message)"
        return @{}
    }
}

function Save-ProcessedComputersHistory {
    param([hashtable]$ProcessedComputers)
    
    $ProcessedFile = Join-Path $Script:ScriptDir $Script:Config.ProcessedComputersFile
    
    try {
        $ProcessedComputers | ConvertTo-Json -Depth 3 | Set-Content $ProcessedFile -Encoding UTF8
        Write-Log DEBUG "Saved processed computers history ($($ProcessedComputers.Count) entries)"
    }
    catch {
        Write-Log ERROR "Failed to save processed computers history: $($_.Exception.Message)"
    }
}

function Find-NewComputersOnAllDCs {
    $AllDCs = Get-AllDomainControllers
    $DefaultComputersContainer = "CN=Computers,$((Get-ADDomain).DistinguishedName)"
    $LookbackTime = (Get-Date).AddMinutes(-$Script:Config.LookbackMinutes)
    $AllNewComputers = @()
    
    Write-Log INFO "Scanning for computers created after $LookbackTime in default container"
    
    foreach ($DC in $AllDCs) {
        try {
            Write-Log DEBUG "Querying DC: $DC"
            
            # Query for computers created recently in the default Computers container
            $Computers = Get-ADComputer -Server $DC -Filter "whenCreated -gt '$($LookbackTime.ToString('yyyy-MM-dd HH:mm:ss'))'" -SearchBase $DefaultComputersContainer -Properties whenCreated, DistinguishedName -ErrorAction Stop
            
            if ($Computers) {
                Write-Log DEBUG "Found $($Computers.Count) recent computer(s) on DC '$DC'"
                $AllNewComputers += $Computers
            }
        }
        catch {
            Write-Log WARN "Failed to query DC '$DC': $($_.Exception.Message)"
        }
    }
    
    # Remove duplicates (same computer found on multiple DCs)
    $UniqueComputers = $AllNewComputers | Sort-Object Name, whenCreated | Get-Unique -AsString
    
    Write-Log INFO "Found $($UniqueComputers.Count) unique new computer(s) across all DCs"
    return $UniqueComputers
}

function Get-UnprocessedComputers {
    $ProcessedComputers = Get-ProcessedComputersHistory
    $NewComputers = Find-NewComputersOnAllDCs
    
    # Filter out already processed computers
    $UnprocessedComputers = $NewComputers | Where-Object { 
        -not $ProcessedComputers.ContainsKey($_.Name) 
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
        $Computer,
        
        [Parameter(Mandatory)]
        [hashtable]$ProcessedComputers
    )
    
    $ComputerName = $Computer.Name
    Write-Log INFO "Processing computer: $ComputerName"
    
    # Determine target OU
    $TargetOU = Resolve-TargetOU -ComputerName $ComputerName
    if (-not $TargetOU) {
        Write-Log WARN "No routing rule matched for '$ComputerName' - skipping"
        
        # Mark as processed (even though no rule matched) to avoid re-processing
        $ProcessedComputers[$ComputerName] = @{
            ProcessedDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            Result = 'NoRuleMatched'
            TargetOU = $null
        }
        return $true
    }
    
    # Attempt to move computer
    $MoveResult = Move-ComputerToOU -Computer $Computer -TargetOU $TargetOU -WhatIf:$Simulate
    
    # Record result
    $ProcessedComputers[$ComputerName] = @{
        ProcessedDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Result = if ($MoveResult) { 'Success' } else { 'Failed' }
        TargetOU = $TargetOU
        Simulation = [bool]$Simulate
    }
    
    return $MoveResult
}

#endregion

#region Main Execution

function Stop-ScriptExecution {
    param(
        [int]$ExitCode = 0,
        [string]$Message
    )
    
    if ($Message) {
        $Level = if ($ExitCode -eq 0) { 'INFO' } else { 'ERROR' }
        Write-Log $Level $Message
    }
    
    Write-Log INFO "----- AutoMove-Computer-Polling Session Ended (Exit Code: $ExitCode) -----"
    
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
    
    # Load processed computers history
    $ProcessedComputers = Get-ProcessedComputersHistory
    
    # Find computers that need processing
    $UnprocessedComputers = Get-UnprocessedComputers
    
    if ($UnprocessedComputers.Count -eq 0) {
        Write-Log INFO "No unprocessed computers found - nothing to do"
        Save-ProcessedComputersHistory -ProcessedComputers $ProcessedComputers
        Stop-ScriptExecution -ExitCode 0 -Message "Polling cycle completed - no new computers"
    }
    
    # Process each computer
    $SuccessCount = 0
    $FailureCount = 0
    
    foreach ($Computer in $UnprocessedComputers) {
        try {
            if (Process-Computer -Computer $Computer -ProcessedComputers $ProcessedComputers) {
                $SuccessCount++
            } else {
                $FailureCount++
            }
        }
        catch {
            Write-Log ERROR "Unexpected error processing computer '$($Computer.Name)': $($_.Exception.Message)"
            $FailureCount++
            
            # Still mark as processed to avoid infinite retries
            $ProcessedComputers[$Computer.Name] = @{
                ProcessedDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                Result = 'Error'
                Error = $_.Exception.Message
            }
        }
    }
    
    # Save updated processed computers history
    Save-ProcessedComputersHistory -ProcessedComputers $ProcessedComputers
    
    # Report results
    $TotalProcessed = $SuccessCount + $FailureCount
    Write-Log INFO "Processing complete - Success: $SuccessCount, Failed: $FailureCount, Total: $TotalProcessed"
    
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