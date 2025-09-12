<#  AutoMove-Computer.ps1  (PowerShell 5.1 compatible)
    Event-driven OU routing for NEW AD computer objects (Security Event 5137).

    - Logs to .\Logs\AutoMove-Computer.log (next to this script)
    - Accepts RecordId as raw string (from Task macro) or manual -ComputerName override
    - Rules: name prefix Alpha/Beta/Gamma -> target OU
    - Optional fallback rule to a quarantine OU

    Task Scheduler Action:
    Program:   C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
    Arguments: -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\AutoMove-Computer.ps1" -RecordId "$(EventRecordID)"
    Start in:  C:\Scripts
#>

param(
  [string]$RecordId,            # may be "$(EventRecordID)" or a number
  [string]$ComputerName,        # manual override for testing
  [switch]$Simulate             # don't move, only log intended action
)

$ErrorActionPreference = 'Stop'

# --- Resolve script path/dirs even if started oddly ---
$ScriptPath = if ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path } else { $PSCommandPath }
if (-not $ScriptPath) { $ScriptPath = 'C:\Scripts\AutoMove-Computer.ps1' }
$ScriptDir  = Split-Path -Parent $ScriptPath

# --- Logging ---
$LogDir  = Join-Path $ScriptDir 'Logs'
$LogFile = Join-Path $LogDir  'AutoMove-Computer.log'
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
function Write-Log {
  param([string]$Level,[string]$Msg)
  $ts=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
  Add-Content -Path $LogFile -Value ("{0} [{1}] {2}" -f $ts,$Level.ToUpper(),$Msg)
}

# optional transcript for deep debugging
try {
  $TranscriptPath = Join-Path $LogDir ("Transcript_{0:yyyyMMdd_HHmmss}.txt" -f (Get-Date))
  Start-Transcript -Path $TranscriptPath -Force | Out-Null
} catch {}

Write-Log INFO "----- Task trigger start -----"
Write-Log INFO ("Host={0} User={1} PS={2} RecordIdRaw='{3}' ComputerNameOverride='{4}' Simulate={5}" -f `
  $env:COMPUTERNAME, [System.Security.Principal.WindowsIdentity]::GetCurrent().Name, $PSVersionTable.PSVersion, $RecordId, $ComputerName, [bool]$Simulate)

# --- Import AD module (required) ---
try {
  Import-Module ActiveDirectory -ErrorAction Stop
  Write-Log INFO "ActiveDirectory module imported."
} catch {
  Write-Log ERROR ("ActiveDirectory module not available: {0}" -f $_.Exception.Message)
  Stop-Transcript | Out-Null
  exit 10
}

# --- Routing rules (edit to your environment) ---
$Rules = @(
  @{ Pattern = '^Alpha'; OU = 'OU=Location Alpha,OU=Computers,OU=DEV Company,DC=dev,DC=local' }
  @{ Pattern = '^Beta';  OU = 'OU=Location Beta,OU=Computers,OU=DEV Company,DC=dev,DC=local' }
  @{ Pattern = '^Gamma'; OU = 'OU=Location Gamma,OU=Computers,OU=DEV Company,DC=dev,DC=local' }
  @{ Pattern = '.*';     OU = 'OU=_Quarantine,OU=Computers,OU=DEV Company,DC=dev,DC=local' } # optional fallback
)

# --- Helpers ---
function Get-5137ByRecordId {
  param([long]$Rid)
  Get-WinEvent -LogName Security -MaxEvents 500 |
    Where-Object { $_.Id -eq 5137 -and $_.RecordId -eq $Rid } |
    Select-Object -First 1
}
function Get-Recent5137Computer {
  param([datetime]$Since)
  Get-WinEvent -FilterHashtable @{LogName='Security'; Id=5137; StartTime=$Since} |
    Where-Object {
      $x = [xml]$_.ToXml()
      ($x.Event.EventData.Data | Where-Object { $_.Name -eq 'ObjectClass' }).'#text' -eq 'computer'
    } | Select-Object -First 1
}
function Resolve-TargetOU {
  param([string]$Name)
  foreach ($r in $Rules) { if ($Name -match $r.Pattern) { return $r.OU } }
  return $null
}

# --- Determine ComputerName (override OR from event 5137) ---
$Creator = $null
if ($ComputerName) {
  Write-Log INFO "Override mode: using ComputerName='$ComputerName' (skipping event lookup)."
} else {
  $RecordIdParsed = $null
  if ($RecordId -match '^\d+$') { $RecordIdParsed = [long]$RecordId }

  $evt = $null
  if ($RecordIdParsed) {
    $evt = Get-5137ByRecordId -Rid $RecordIdParsed
    if ($evt) { Write-Log INFO "Matched Security 5137 by RecordId=$RecordIdParsed" }
  }
  if (-not $evt) {
    $since = (Get-Date).AddMinutes(-5)
    Write-Log WARN "No event by RecordId. Falling back to search since $since"
    $evt = Get-Recent5137Computer -Since $since
    if (-not $evt) {
      Write-Log ERROR "No recent 5137 computer event found. Exiting."
      Stop-Transcript | Out-Null; exit 20
    }
  }
  [xml]$x = $evt.ToXml()
  $edata = @{}; foreach ($d in $x.Event.EventData.Data) { $edata[$d.Name] = [string]$d.'#text' }
  if ($edata['ObjectClass'] -ne 'computer') {
    Write-Log INFO "Event ObjectClass is '$($edata['ObjectClass'])'. Exit."
    Stop-Transcript | Out-Null; exit 0
  }
  $dn = $edata['ObjectDN']
  $ComputerName = ($dn -split ',')[0] -replace '^CN=',''
  $Creator = ('{0}\{1}' -f $edata['SubjectDomainName'],$edata['SubjectUserName'])
  Write-Log INFO ("New computer from event: Name={0} DN={1} Creator={2}" -f $ComputerName,$dn,($Creator))
}

# --- Wait for replication visibility (up to 60s) ---
$ad = $null
foreach ($i in 1..12) {
  try {
    $ad = Get-ADComputer -Identity $ComputerName -Properties DistinguishedName,whenCreated -ErrorAction Stop
    break
  } catch { Start-Sleep -Seconds 5 }
}
if (-not $ad) {
  Write-Log ERROR "Computer '$ComputerName' not visible in AD after 60s. Exiting."
  Stop-Transcript | Out-Null; exit 30
}

# --- Determine destination OU ---
$TargetOU = Resolve-TargetOU -Name $ComputerName
if (-not $TargetOU) {
  Write-Log WARN "No routing rule matched for '$ComputerName'. Exiting."
  Stop-Transcript | Out-Null; exit 0
}

# Already in target?
if ($ad.DistinguishedName -like "*$TargetOU") {
  Write-Log INFO ("{0} already in {1}. Nothing to do." -f $ComputerName,$TargetOU)
  Stop-Transcript | Out-Null; exit 0
}

# Ensure OU exists
try { $null = Get-ADOrganizationalUnit -Identity $TargetOU -ErrorAction Stop }
catch {
  Write-Log ERROR ("Target OU not found: {0} — {1}" -f $TargetOU,$_.Exception.Message)
  Stop-Transcript | Out-Null; exit 40
}

# --- Move or simulate ---
if ($Simulate) {
  Write-Log INFO ("SIMULATE move: {0}  -->  {1}" -f $ad.DistinguishedName,$TargetOU)
  Stop-Transcript | Out-Null; exit 0
}

try {
  Write-Log INFO ("Moving: {0}  -->  {1}" -f $ad.DistinguishedName,$TargetOU)
  Move-ADObject -Identity $ad.DistinguishedName -TargetPath $TargetOU -Confirm:$false
  $creatorOut = if ($Creator) { $Creator } else { 'n/a' }
  Write-Log SUCCESS ("Moved {0} to {1} (Creator={2})" -f $ComputerName,$TargetOU,$creatorOut)
} catch {
  Write-Log ERROR ("Move failed: {0}" -f $_.Exception.Message)
  Stop-Transcript | Out-Null; exit 50
}

Stop-Transcript | Out-Null
