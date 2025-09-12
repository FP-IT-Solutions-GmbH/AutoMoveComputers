<#
.SYNOPSIS
    Test script for AutoMove-Computer configuration and rules.

.DESCRIPTION
    This script validates the configuration file and tests routing rules without 
    requiring an actual Active Directory environment.

.PARAMETER ConfigFile
    Path to the configuration file to test (defaults to Config.psd1).

.PARAMETER TestComputerNames
    Array of computer names to test against the routing rules.

.EXAMPLE
    Test-AutoMoveConfig.ps1
    Tests the default configuration with sample computer names.

.EXAMPLE
    Test-AutoMoveConfig.ps1 -TestComputerNames @('ALPHA-PC001', 'BETA-LAPTOP05', 'UNKNOWN-SYSTEM')
    Tests specific computer names against the routing rules.
#>

[CmdletBinding()]
param(
    [string]$ConfigFile = ".\Config.psd1",
    [string[]]$TestComputerNames = @('ALPHA-PC001', 'BETA-LAPTOP05', 'GAMMA-WS010', 'LAB-COMPUTER1', 'KIOSK-LOBBY01', 'UNKNOWN-SYSTEM')
)

function Test-Configuration {
    param([string]$Path)
    
    Write-Host "Testing configuration file: $Path" -ForegroundColor Cyan
    
    if (-not (Test-Path $Path)) {
        Write-Host "[ERROR] Configuration file not found: $Path" -ForegroundColor Red
        return $false
    }
    
    try {
        $Config = Import-PowerShellDataFile -Path $Path
        Write-Host "[OK] Configuration file loaded successfully" -ForegroundColor Green
        
        # Validate required properties
        $RequiredProps = @('Rules', 'LogDir', 'ReplicationTimeoutSeconds')
        foreach ($Prop in $RequiredProps) {
            if (-not $Config.ContainsKey($Prop)) {
                Write-Host "[WARNING] Missing required property: $Prop" -ForegroundColor Yellow
            } else {
                Write-Host "[OK] Found property: $Prop" -ForegroundColor Green
            }
        }
        
        # Validate rules
        if ($Config.Rules -and $Config.Rules.Count -gt 0) {
            Write-Host "[OK] Found $($Config.Rules.Count) routing rules" -ForegroundColor Green
            
            foreach ($Rule in $Config.Rules) {
                if (-not $Rule.Pattern) {
                    Write-Host "[WARNING] Rule missing Pattern property" -ForegroundColor Yellow
                } elseif (-not $Rule.OU) {
                    Write-Host "[WARNING] Rule missing OU property for pattern: $($Rule.Pattern)" -ForegroundColor Yellow
                } else {
                    $Description = if ($Rule.Description) { " ($($Rule.Description))" } else { "" }
                    Write-Host "   [RULE] $($Rule.Pattern) -> $($Rule.OU)$Description" -ForegroundColor Gray
                }
            }
        } else {
            Write-Host "[ERROR] No routing rules found" -ForegroundColor Red
        }
        
        return $Config
    }
    catch {
        Write-Host "[ERROR] Failed to load configuration: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Test-RoutingRules {
    param($Config, [string[]]$ComputerNames)
    
    Write-Host "`nTesting routing rules with sample computer names:" -ForegroundColor Cyan
    
    foreach ($ComputerName in $ComputerNames) {
        $Matched = $false
        
        foreach ($Rule in $Config.Rules) {
            if ($ComputerName -match $Rule.Pattern) {
                $Description = if ($Rule.Description) { " ($($Rule.Description))" } else { "" }
                Write-Host "[MATCH] $ComputerName -> $($Rule.OU)$Description" -ForegroundColor Green
                $Matched = $true
                break
            }
        }
        
        if (-not $Matched) {
            Write-Host "[NO MATCH] $ComputerName -> No matching rule found" -ForegroundColor Red
        }
    }
}

# Main execution
Write-Host "AutoMove-Computer Configuration Tester" -ForegroundColor Magenta
Write-Host "=====================================" -ForegroundColor Magenta

$Config = Test-Configuration -Path $ConfigFile

if ($Config) {
    Test-RoutingRules -Config $Config -ComputerNames $TestComputerNames
    
    Write-Host "`n[SUMMARY] Configuration Summary:" -ForegroundColor Cyan
    Write-Host "   Log Directory: $($Config.LogDir)" -ForegroundColor Gray
    Write-Host "   Max Log Size: $($Config.MaxLogSizeMB) MB" -ForegroundColor Gray
    Write-Host "   Replication Timeout: $($Config.ReplicationTimeoutSeconds) seconds" -ForegroundColor Gray
    Write-Host "   Event Lookup Window: $($Config.EventLookupMinutes) minutes" -ForegroundColor Gray
    Write-Host "   Transcript Enabled: $($Config.EnableTranscript)" -ForegroundColor Gray
} else {
    Write-Host "`n[ERROR] Configuration test failed" -ForegroundColor Red
    exit 1
}

Write-Host "`n[OK] Configuration test completed" -ForegroundColor Green
