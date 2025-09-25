
# Automatisch neue Computer in OU verschieben

## Überblick
Dieses Script verschiebt **neu erstellte Computerobjekte** automatisch in die richtige **Active Directory-OU## Logs & Troubleshooting

---

## Deployment & Maintenance (v2.0)

### Empfohlenes Setup
1. **Skript zentral ablegen:** `C:\Scripts\AutoMove-Computer.ps1`
2. **Konfiguration anpassen:** `C:\Scripts\Config.psd1` erstellen/editieren
3. **Logs-Verzeichnis:** `C:\Scripts\Logs\` (wird automatisch erstellt)
4. **GPO:** Geplante Aufgabe auf Domain Controllers OU verteilen

### Monitoring & Wartung
```powershell
# Log-Größe prüfen
Get-ChildItem "C:\Scripts\Logs\*.log" | Select Name, @{n='SizeMB';e={[math]::Round($_.Length/1MB,2)}}

# Letzte Ausführungen anzeigen
Get-Content "C:\Scripts\Logs\AutoMove-Computer.log" | Select-String "Session Started" | Select -Last 10

# Erfolgreiche Moves anzeigen
Get-Content "C:\Scripts\Logs\AutoMove-Computer.log" | Select-String "SUCCESS.*Moved" | Select -Last 20

# Echte Simulation für neue Regeln (ohne AD-Zugriff)
.\AutoMove-Computer.ps1 -ComputerName "TestComputer" -Simulate

# Routing-Test für mehrere Computer
@('ALPHA-PC001', 'BETA-WS002', 'LAB-SYSTEM01', 'UNKNOWN-PC') | ForEach-Object {
    .\AutoMove-Computer.ps1 -ComputerName $_ -Simulate
}
```
## Funktionsweise 
- **GPO** aktiviert **Verzeichnisdienständerungen überwachen** (siehe unten).
- **SACL** auf dem Quellcontainer (`CN=Computers`) löst Event 5137 bei neuer Computererstellung aus.
- **Geplante Aufgabe** auf jedem DC triggert auf Event 5137 und ruft das PowerShell-Skript auf.
- Das Skript:
  1. Liest Name und DN des neu erstellten Computers aus dem Eventlog.
  2. Prüft anhand von Regex-Regeln, in welche OU der Computer gehört.
  3. Verschiebt ihn automatisch dorthin (oder in eine Quarantäne-OU, falls kein Match).

---

## Gruppenrichtlinie – Auditing aktivieren

Erstelle eine neue **GPO** und verknüpfe sie mit der **Domain Controllers OU**.

**Pfad:**  
`Computerkonfiguration → Richtlinien → Windows-Einstellungen → Sicherheitseinstellungen → Erweiterte Überwachungsrichtlinienkonfiguration → Überwachungsrichtlinien → DS-Zugriff`

Aktiviere:
- **Verzeichnisdienstzugriff überwachen → Erfolgreich**
- **Verzeichnisdienständerungen überwachen → Erfolgreich**

> Danach `gpupdate /force` auf allen DCs ausführen.

---

## SACL auf Quellcontainer setzen
Auf `CN=Computers` (oder wo neue Rechner landen):
1. **Eigenschaften → Sicherheit → Erweitert → Überwachung → Hinzufügen**
2. Principal: **Jeder**
3. Anwenden auf: **Dieses Objekt und alle untergeordneten Objekte**
4. Berechtigung: **Erstellen von Computerobjekten**

Damit wird bei jeder Computererstellung ein Event 5137 geschrieben.

---

## Scheduled Task auf DC einrichten

### Trigger (XML-Filter)
Benutzerdefiniert → **XML** → *Manuell bearbeiten* aktivieren und einfügen:

```xml
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[(EventID=5137)]]
      and
      *[EventData[Data[@Name='ObjectClass']='computer']]
    </Select>
  </Query>
</QueryList>
```

**Optional → nur Objekte aus `CN=Computers` triggern (Domäne anpassen):**
```xml
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[(EventID=5137)]]
      and
      *[EventData[Data[@Name='ObjectClass']='computer']]
      and
      *[EventData[Data[@Name='ObjectDN'] and contains(.,',CN=Computers,DC=dev,DC=local')]]
    </Select>
  </Query>
</QueryList>
```

### Aktion
```
Programm:  C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
Argumente: -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\AutoMove-Computer.ps1" -RecordId "$(EventRecordID)"
Starten in: C:\Scripts
```
- **Mit höchsten Privilegien ausführen** aktivieren.
- Konto mit **Berechtigung zum Verschieben von Computern** verwenden (Delegation auf Quellcontainer und Ziel-OU).

---

## Konfiguration

### Config.psd1 (Neu in v2.0)
Das Script unterstützt jetzt eine externe Konfigurationsdatei `Config.psd1` im gleichen Verzeichnis:

```powershell
@{
    # Logging-Einstellungen
    LogDir = '.\Logs'
    MaxLogSizeMB = 10                    # Automatische Log-Rotation
    EnableTranscript = $true             # PowerShell Transcript für Debug
    
    # AD-Replikation
    ReplicationTimeoutSeconds = 60       # Max. Wartezeit auf AD-Sichtbarkeit
    ReplicationRetryIntervalSeconds = 5  # Wiederholungsintervall
    
    # Event-Verarbeitung
    EventLookupMinutes = 5              # Fallback: Events der letzten N Minuten
    MaxEvents = 500                     # Max. Events zum Durchsuchen
    
    # Routing-Regeln
    Rules = @(
        @{ Pattern = '^Alpha'; OU = 'OU=Location Alpha,...'; Description = 'Alpha Standort' }
        @{ Pattern = '^Beta';  OU = 'OU=Location Beta,...';  Description = 'Beta Standort' }
        # ... weitere Regeln
    )
}
```

### Script-Parameter (Erweitert)
```powershell
# Standard Task Scheduler Aufruf
AutoMove-Computer.ps1 -RecordId "$(EventRecordID)"

# Echte Simulation - testet nur Routing-Regeln ohne AD-Zugriff
AutoMove-Computer.ps1 -ComputerName "Alpha-PC001" -Simulate

# Benutzerdefinierte Konfigurationsdatei
AutoMove-Computer.ps1 -RecordId "12345" -ConfigFile "C:\Custom\MyConfig.psd1"
```

**Wichtig:** Der `-Simulate` Modus ist eine echte Simulation, die:
- ✅ Keine Active Directory-Module benötigt
- ✅ Keine AD-Verbindung aufbaut
- ✅ Keine Computer im AD sucht
- ✅ Nur Routing-Regeln testet und loggt
- ✅ Sofort terminiert ohne Wartezeiten

---

## Anpassung / Erweiterung

### Routing-Regeln anpassen
**Option 1: Direkt in Config.psd1** (empfohlen)
```powershell
Rules = @(
    @{ Pattern = '^ALPHA-'; OU = 'OU=Alpha Computers,DC=domain,DC=com'; Description = 'Alpha Standort' }
    @{ Pattern = '^BETA-';  OU = 'OU=Beta Computers,DC=domain,DC=com';  Description = 'Beta Standort' }
    @{ Pattern = '^LAB-';   OU = 'OU=Lab Systems,DC=domain,DC=com';     Description = 'Labor-PCs' }
    @{ Pattern = '.*';      OU = 'OU=Quarantine,DC=domain,DC=com';      Description = 'Fallback' }
)
```

**Wichtig:** 
- `Pattern` ist ein **Regex** für den Computernamen
- Verwende `^` für "beginnt mit" und `$` für "endet mit"
- Reihenfolge ist wichtig – erste Übereinstimmung gewinnt
- `.*` am Ende als Fallback für unbekannte Computer

### Erweiterte Beispiele
```powershell
# Nach Standort-Prefix
@{ Pattern = '^(NYC|BOS|ATL)-'; OU = 'OU=US East Coast,DC=corp,DC=com' }

# Nach Computer-Typ
@{ Pattern = '-VM\d+$'; OU = 'OU=Virtual Machines,DC=corp,DC=com' }

# Mehrere Patterns kombiniert
@{ Pattern = '^(KIOSK|INFO|DISPLAY)-'; OU = 'OU=Public Systems,DC=corp,DC=com' }
```

---

## Logs & Troubleshooting
- Logs liegen in `C:\Scripts\Logs\AutoMove-Computer.log` (plus Transcripts).
- Task-Historie: **201 Action completed** mit Rückgabecode **0** = OK.
- Häufige Fehler:
  - Task ohne „Starten in“ → Skript/Log nicht gefunden.
  - Event entsteht auf anderem DC → Task auf **allen DCs** ausrollen.
  - Fehlende Rechte → Delegation auf Quellcontainer/Ziel-OU prüfen.

---

## Deployment-Tipp
- **Empfohlen:** GPO-Preferences → Geplante Aufgabe auf **Domain Controllers OU** verteilen, damit sie auf allen DCs existiert.
- Skript zentral (z. B. `C:\Scripts`) per GPO oder Softwareverteilung auf alle DCs kopieren.

---
### Log-Level
- **INFO:** Normale Ausführung, Konfiguration, Fortschritt
- **SUCCESS:** Erfolgreiche Computer-Verschiebung
- **WARN:** Nicht-kritische Probleme (z.B. Event nicht gefunden, fallback zu recent search)
- **ERROR:** Kritische Fehler, die zum Skript-Abbruch führen
- **DEBUG:** Detaillierte Informationen für Fehlerdiagnose

### Exit Codes
- **0:** Erfolgreiche Ausführung oder Computer bereits am richtigen Ort
- **10:** ActiveDirectory PowerShell-Modul nicht verfügbar
- **20:** Kein passendes Security Event 5137 gefunden
- **21:** Event-Parsing fehlgeschlagen
- **30:** Computer nach Wartezeit nicht in AD sichtbar
- **40:** Ziel-OU existiert nicht oder ist nicht zugänglich
- **50:** Computer-Verschiebung fehlgeschlagen
- **99:** Unbehandelte Exception

### Häufige Probleme und Lösungen

#### Event nicht gefunden (Exit Code 20)
```
WARN: No event found by RecordId '$(EventRecordID)', searching for recent computer creation events
ERROR: No suitable Security Event 5137 found for computer creation
```
**Ursachen:**
- Task-Makro `$(EventRecordID)` wird nicht korrekt aufgelöst
- Event wurde auf anderem DC erstellt
- Zeitfenster zu kurz (EventLookupMinutes erhöhen)

**Lösungen:**
- Task auf allen DCs einrichten
- EventLookupMinutes in Config.psd1 erhöhen (Standard: 5 Minuten)
- Mit `-ComputerName` manuell testen

#### AD-Replikation Timeout (Exit Code 30)
```
ERROR: Computer 'ALPHA-PC001' not visible in AD after 60 seconds
```
**Ursachen:**
- Langsame AD-Replikation zwischen DCs
- Computer wurde sofort nach Erstellung wieder gelöscht

**Lösungen:**
- ReplicationTimeoutSeconds in Config.psd1 erhöhen
- Task auf dem DC einrichten, wo Computer erstellt werden

#### Berechtigungsfehler (Exit Code 50)
```
ERROR: Failed to move computer 'ALPHA-PC001': Access is denied
```
**Lösungen:**
- Task-Konto benötigt "Move" Berechtigung auf Quell- und Ziel-Container
- Computer Groups-Membership des Task-Kontos prüfen
- Delegation auf OU-Ebene einrichten

---

# AutoMove-Computer-Polling (Alternative Lösung)

## Überblick - Polling-basierter Ansatz

**Neue Alternative für Multi-DC Umgebungen:** Anstatt ereignisbasiert zu arbeiten, überwacht diese Variante kontinuierlich alle Domain Controller nach neuen Computern im Standard-Container.

### Warum Polling statt Events?

**Problem mit Events:**
- Events entstehen nur auf dem DC, wo Computer erstellt wird
- Script muss auf **jedem** DC laufen
- Komplexe Synchronisation zwischen DCs
- Replication-Timing-Probleme

**Lösung mit Polling:**
- ✅ **Ein einziger Deployment-Punkt** (z.B. PDC Emulator)
- ✅ **100% Abdeckung** - findet Computer egal auf welchem DC erstellt
- ✅ **Keine Replikations-Probleme** - fragt DCs direkt ab
- ✅ **Einfachere Architektur** - keine Event-Synchronisation
- ✅ **Bessere Fehlerbehandlung** - kann fehlgeschlagene Moves wiederholen

## Dateien

- **`AutoMove-Computer-Polling.ps1`** - Hauptscript (Polling-Variante v3.1)
- **`Config-Polling.psd1`** - Konfigurationsdatei für Polling
- **`Test-AutoMovePolling.ps1`** - Test-Utility für Konfiguration
- **AD extensionAttribute7** - Tracking bereits verarbeiteter Computer (im AD gespeichert)

## Schnellstart

### 1. Konfiguration testen
```powershell
.\Test-AutoMovePolling.ps1
```

### 2. Simulation
```powershell
.\AutoMove-Computer-Polling.ps1 -Simulate
```

### 3. Produktiv einsetzen
```powershell
# Geplante Aufgabe: Alle 3 Minuten
.\AutoMove-Computer-Polling.ps1
```

## Deployment

### Geplante Aufgabe (Empfohlen)
```
Programm: powershell.exe
Argumente: -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\AutoMove-Computer-Polling.ps1"
Trigger: Alle 3 Minuten wiederholen
Konto: Domain Admin oder delegierte Berechtigung
Ausführen auf: PDC Emulator (oder ein zentraler DC)
```

### Vs. Event-basierte Lösung

| Aspekt | Event-basiert | Polling-basiert |
|--------|---------------|-----------------|
| **Deployment** | Jeder DC | Ein DC |
| **Abdeckung** | Nur lokale Events | Alle DCs |
| **Latenz** | Sofort | 1-3 Minuten |
| **Komplexität** | Hoch | Niedrig |
| **Replication-Issues** | Ja | Nein |
| **Multi-DC Umgebung** | Problematisch | Optimal |

## Vorteile für Ihre Umgebung

**Ihr konkreter Fall:**
```
Host: MFS-DC01
Computer: MFS-LY-TESTERTERST
Problem: Computer auf anderem DC erstellt, MFS-DC01 wartete 60s auf Replikation
```

**Mit Polling-Lösung:**
```
[INFO] Discovered 5 domain controllers: MFS-DC01, MFS-DC02, MFS-DC03, MFS-DC04, MFS-DC05
[INFO] Querying DC: MFS-DC03  
[SUCCESS] Found MFS-LY-TESTERTERST on DC MFS-DC03
[SUCCESS] Successfully moved to target OU
```

**Resultat:** Statt 60s Timeout → 2s Erfolg, 100% Zuverlässigkeit.

---
