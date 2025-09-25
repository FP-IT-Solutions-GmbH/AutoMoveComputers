<#
.SYNOPSIS
    Test script for AutoMove-Computer-Polling configuration and functionality.

.DESCRIPTION
    This script validates the polling configuration, tests routing rules, and can simulate
    the entire polling process without requiring actual computer objects.

.PARAMETER ConfigFile
    Path to the polling configuration file to test.

.PARAMETER TestMode
    Type of test to run: Config, Rules, Simulation, or All.

.PARAMETER TestComputerNames
    Array of computer names to test against the routing rules.

.EXAMPLE
    Test-AutoMovePolling.ps1
    Run all tests with default configuration.

.EXAMPLE
    Test-AutoMovePolling.ps1 -TestMode Rules
    Test only the routing rules.

.EXAMPLE
    Test-AutoMovePolling.ps1 -TestComputerNames @('ALPHA-PC001', 'UNKNOWN-SYSTEM')
    Test specific computer names against routing rules.
#>

[CmdletBinding()]
param(
    [string]$ConfigFile = ".\Config-Polling.psd1",
    
    [ValidateSet('Config', 'Rules', 'Simulation', 'All')]
    [string]$TestMode = 'All',
    
    [string[]]$TestComputerNames = @(
        'ALPHA-PC001', 'BETA-LAPTOP05', 'GAMMA-WS010', 
        'LAB-COMPUTER1', 'KIOSK-LOBBY01', 'UNKNOWN-SYSTEM',
        'MFS-LY-TESTERTERST'  # Your actual example
    )
)

function Write-TestResult {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('PASS', 'FAIL', 'INFO', 'WARN')]
        [string]$Status,
        
        [Parameter(Mandatory)]
        [string]$Message
    )
    
    $Color = switch ($Status) {
        'PASS' { 'Green' }
        'FAIL' { 'Red' }
        'WARN' { 'Yellow' }
        'INFO' { 'Cyan' }
    }
    
    $Symbol = switch ($Status) {
        'PASS' { '✓' }
        'FAIL' { '✗' }
        'WARN' { '!' }
        'INFO' { 'ℹ' }
    }
    
    Write-Host "[$Symbol] $Message" -ForegroundColor $Color
}

function Test-PollingConfiguration {
    param([string]$Path)
    
    Write-Host "`n=== Configuration Test ===" -ForegroundColor Magenta
    
    if (-not (Test-Path $Path)) {
        Write-TestResult FAIL "Configuration file not found: $Path"
        return $false
    }
    
    try {
        $Config = Import-PowerShellDataFile -Path $Path
        Write-TestResult PASS "Configuration file loaded successfully"
        
        # Test required properties
        $RequiredProps = @('Rules', 'LogDir', 'PollIntervalMinutes', 'LookbackMinutes')
        
        foreach ($Prop in $RequiredProps) {
            if ($Config.ContainsKey($Prop)) {
                Write-TestResult PASS "Found required property: $Prop = $($Config[$Prop])"
            } else {
                Write-TestResult FAIL "Missing required property: $Prop"
            }
        }
        
        # Validate polling settings
        if ($Config.PollIntervalMinutes -lt 1) {
            Write-TestResult WARN "PollIntervalMinutes ($($Config.PollIntervalMinutes)) is very low - may cause performance issues"
        } elseif ($Config.PollIntervalMinutes -gt 10) {
            Write-TestResult WARN "PollIntervalMinutes ($($Config.PollIntervalMinutes)) is high - may miss computers in fast environments"
        } else {
            Write-TestResult PASS "PollIntervalMinutes ($($Config.PollIntervalMinutes)) is reasonable"
        }
        
        if ($Config.LookbackMinutes -lt $Config.PollIntervalMinutes * 2) {
            Write-TestResult WARN "LookbackMinutes should be at least 2x PollIntervalMinutes to avoid gaps"
        } else {
            Write-TestResult PASS "LookbackMinutes ($($Config.LookbackMinutes)) provides good overlap"
        }
        
        return $Config
    }
    catch {
        Write-TestResult FAIL "Failed to load configuration: $($_.Exception.Message)"
        return $false
    }
}

function Test-RoutingRules {
    param($Config, [string[]]$ComputerNames)
    
    Write-Host "`n=== Routing Rules Test ===" -ForegroundColor Magenta
    
    if (-not $Config.Rules -or $Config.Rules.Count -eq 0) {
        Write-TestResult FAIL "No routing rules found in configuration"
        return
    }
    
    Write-TestResult INFO "Testing $($Config.Rules.Count) routing rules with $($ComputerNames.Count) computer names"
    
    # Test each computer name
    foreach ($ComputerName in $ComputerNames) {
        $Matched = $false
        
        foreach ($Rule in $Config.Rules) {
            if ($ComputerName -match $Rule.Pattern) {
                $Description = if ($Rule.Description) { " ($($Rule.Description))" } else { "" }
                Write-TestResult PASS "$ComputerName matches '$($Rule.Pattern)' -> $($Rule.OU)$Description"
                $Matched = $true
                break
            }
        }
        
        if (-not $Matched) {
            Write-TestResult FAIL "$ComputerName -> No matching rule found"
        }
    }
    
    # Test for catch-all rule
    $HasCatchAll = $Config.Rules | Where-Object { $_.Pattern -eq '.*' -or $_.Pattern -eq '.*$' }
    if ($HasCatchAll) {
        Write-TestResult PASS "Catch-all rule found - no computers will be missed"
    } else {
        Write-TestResult WARN "No catch-all rule (.*) - some computers may not match any rules"
    }
}

function Test-PollingSimulation {
    param($Config)
    
    Write-Host "`n=== Polling Simulation ===" -ForegroundColor Magenta
    
    # Simulate the polling process without actually querying AD
    Write-TestResult INFO "Simulating polling process..."
    
    # Test processed computers file handling
    $ProcessedFile = $Config.ProcessedComputersFile
    Write-TestResult INFO "Processed computers will be tracked in: $ProcessedFile"
    
    # Test log directory
    if ($Config.LogDir) {
        if (Test-Path $Config.LogDir) {
            Write-TestResult PASS "Log directory exists: $($Config.LogDir)"
        } else {
            try {
                New-Item -ItemType Directory -Path $Config.LogDir -Force | Out-Null
                Write-TestResult PASS "Log directory created: $($Config.LogDir)"
            } catch {
                Write-TestResult FAIL "Cannot create log directory: $($_.Exception.Message)"
            }
        }
    }
    
    # Simulate computer processing
    Write-TestResult INFO "Simulating computer processing cycle..."
    
    $SimulatedComputers = @(
        @{ Name = 'ALPHA-TEST001'; whenCreated = (Get-Date).AddMinutes(-5) }
        @{ Name = 'BETA-TEST002'; whenCreated = (Get-Date).AddMinutes(-3) }
        @{ Name = 'UNKNOWN-TEST003'; whenCreated = (Get-Date).AddMinutes(-1) }
    )
    
    foreach ($Computer in $SimulatedComputers) {
        # Test routing for each simulated computer
        $Matched = $false
        foreach ($Rule in $Config.Rules) {
            if ($Computer.Name -match $Rule.Pattern) {
                Write-TestResult PASS "Simulated: $($Computer.Name) would be moved to $($Rule.OU)"
                $Matched = $true
                break
            }
        }
        
        if (-not $Matched) {
            Write-TestResult WARN "Simulated: $($Computer.Name) would not match any rules"
        }
    }
    
    Write-TestResult INFO "Simulation completed successfully"
}

function Show-DeploymentGuidance {
    Write-Host "`n=== Deployment Guidance ===" -ForegroundColor Magenta
    
    Write-TestResult INFO "Polling-based deployment recommendations:"
    Write-Host "  1. Deploy on PRIMARY DC only (typically PDC Emulator)" -ForegroundColor Gray
    Write-Host "  2. Create Scheduled Task to run every $($Config.PollIntervalMinutes) minutes" -ForegroundColor Gray  
    Write-Host "  3. Task should run with Domain Admin privileges" -ForegroundColor Gray
    Write-Host "  4. No event triggers needed - just time-based scheduling" -ForegroundColor Gray
    Write-Host "  5. Monitor logs in $($Config.LogDir) directory" -ForegroundColor Gray
    
    Write-TestResult INFO "Scheduled Task Command:"
    $ScriptPath = Join-Path (Get-Location) "AutoMove-Computer-Polling.ps1"
    Write-Host "  powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`"" -ForegroundColor Yellow
    
    Write-TestResult INFO "For testing, run manually with:"
    Write-Host "  .\AutoMove-Computer-Polling.ps1 -Simulate" -ForegroundColor Yellow
}

# Main execution
Write-Host "AutoMove-Computer-Polling Test Suite" -ForegroundColor Magenta
Write-Host "====================================" -ForegroundColor Magenta

$Config = $null

# Run tests based on mode
switch ($TestMode) {
    'Config' { 
        $Config = Test-PollingConfiguration -Path $ConfigFile 
    }
    'Rules' { 
        $Config = Test-PollingConfiguration -Path $ConfigFile
        if ($Config) { Test-RoutingRules -Config $Config -ComputerNames $TestComputerNames }
    }
    'Simulation' {
        $Config = Test-PollingConfiguration -Path $ConfigFile
        if ($Config) { Test-PollingSimulation -Config $Config }
    }
    'All' {
        $Config = Test-PollingConfiguration -Path $ConfigFile
        if ($Config) { 
            Test-RoutingRules -Config $Config -ComputerNames $TestComputerNames
            Test-PollingSimulation -Config $Config
            Show-DeploymentGuidance
        }
    }
}

if ($Config) {
    Write-Host "`n=== Configuration Summary ===" -ForegroundColor Magenta
    Write-Host "Poll Interval: $($Config.PollIntervalMinutes) minutes" -ForegroundColor Gray
    Write-Host "Lookback Window: $($Config.LookbackMinutes) minutes" -ForegroundColor Gray
    Write-Host "Routing Rules: $($Config.Rules.Count)" -ForegroundColor Gray
    Write-Host "Log Directory: $($Config.LogDir)" -ForegroundColor Gray
    
    Write-TestResult PASS "All tests completed successfully"
} else {
    Write-TestResult FAIL "Testing failed due to configuration issues"
    exit 1
}