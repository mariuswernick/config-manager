<#
.SYNOPSIS
    SCCM User Device Affinity (UDA) Migration - Export & Import
.DESCRIPTION
    Zweistufiges Migrationsskript fuer UDA-Daten zwischen Alt- und Neu-SCCM.
    
    Stufe 1 (Export):  Liest alle aktiven UDAs aus der alten SCCM-Datenbank via SQL
    Stufe 2 (Import):  Setzt UDAs im neuen SCCM, prueft ob Geraet vorhanden ist
    
    Features:
    - Re-Run-faehig: Ueberspringt bereits gesetzte UDAs
    - Vollstaendiges Logging (Transcript + CSV-Report)
    - Zusammenfassung am Ende jedes Laufs
    - Kann nach jedem Onboarding-Batch erneut ausgefuehrt werden

.PARAMETER Mode
    Export  = UDAs aus Alt-SCCM exportieren
    Import  = UDAs in Neu-SCCM importieren
    Report  = Aktuellen Stand anzeigen (was wurde gesetzt, was steht aus)

.PARAMETER ExportCsvPath
    Pfad zur CSV-Datei (Export schreibt, Import liest)

.PARAMETER AltSqlServer
    SQL Server der alten SCCM-Umgebung (nur fuer Export)

.PARAMETER AltDatabase
    Datenbankname der alten SCCM-Umgebung, z.B. CM_P01 (nur fuer Export)

.PARAMETER NeuSiteCode
    Site Code der neuen SCCM-Umgebung (nur fuer Import)

.PARAMETER NeuSiteServer
    Site Server der neuen SCCM-Umgebung (nur fuer Import)

.PARAMETER LogPath
    Verzeichnis fuer Log-Dateien (Standard: Verzeichnis der CSV-Datei)

.EXAMPLE
    # Export aus Alt-SCCM
    .\SCCM-UDA-Migration.ps1 -Mode Export -AltSqlServer "SQL-ALT" -AltDatabase "CM_P01" -ExportCsvPath "C:\Migration\UDA_Export.csv"

.EXAMPLE
    # Import in Neu-SCCM (nach Onboarding)
    .\SCCM-UDA-Migration.ps1 -Mode Import -NeuSiteCode "N01" -NeuSiteServer "SCCM-NEU" -ExportCsvPath "C:\Migration\UDA_Export.csv"

.EXAMPLE
    # Report - Aktuellen Stand pruefen
    .\SCCM-UDA-Migration.ps1 -Mode Report -NeuSiteCode "N01" -NeuSiteServer "SCCM-NEU" -ExportCsvPath "C:\Migration\UDA_Export.csv"

.NOTES
    Autor:   Marius
    Version: 1.0
    Datum:   2026-03-10
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("Export", "Import", "Report")]
    [string]$Mode,

    [Parameter(Mandatory = $true)]
    [string]$ExportCsvPath,

    [Parameter(Mandatory = $false)]
    [string]$AltSqlServer,

    [Parameter(Mandatory = $false)]
    [string]$AltDatabase,

    [Parameter(Mandatory = $false)]
    [string]$NeuSiteCode,

    [Parameter(Mandatory = $false)]
    [string]$NeuSiteServer,

    [Parameter(Mandatory = $false)]
    [string]$LogPath
)

#region ============================================================
#       KONFIGURATION & INITIALISIERUNG
#endregion =========================================================

$ErrorActionPreference = "Stop"
$ProgressPreference    = "SilentlyContinue"

# Log-Verzeichnis bestimmen
if (-not $LogPath) {
    $LogPath = Split-Path -Path $ExportCsvPath -Parent
}
if (-not (Test-Path $LogPath)) {
    New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
}

$Timestamp    = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile      = Join-Path $LogPath "UDA_Migration_${Mode}_${Timestamp}.log"
$ReportCsv    = Join-Path $LogPath "UDA_Migration_Report_${Timestamp}.csv"

# Transcript starten
Start-Transcript -Path $LogFile -Append

try {

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  SCCM UDA Migration - Modus: $Mode" -ForegroundColor Cyan
Write-Host "  Gestartet: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "  CSV-Pfad:  $ExportCsvPath" -ForegroundColor Cyan
Write-Host "  Log-Pfad:  $LogFile" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

#region ============================================================
#       HILFSFUNKTIONEN
#endregion =========================================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "OK", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "OK"    { "Green" }
        "WARN"  { "Yellow" }
        "ERROR" { "Red" }
        default { "White" }
    }
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $color
}

function Connect-NeuSCCM {
    param(
        [string]$SiteCode,
        [string]$SiteServer
    )
    
    Write-Log "Verbinde mit Neu-SCCM: $SiteServer (Site: $SiteCode)"
    
    # ConfigMgr-Modul laden
    $ModulePath = "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
    if (-not (Test-Path $ModulePath)) {
        # Fallback: Standard-Installationspfad
        $ModulePath = "${env:ProgramFiles(x86)}\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1"
    }
    if (-not (Test-Path $ModulePath)) {
        throw "ConfigurationManager PowerShell-Modul nicht gefunden. Bitte SCCM Admin Console installieren."
    }
    
    Import-Module $ModulePath -Force
    
    # PSDrive erstellen falls nicht vorhanden
    if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer | Out-Null
    }
    
    Set-Location "${SiteCode}:"
    Write-Log "Verbunden mit $SiteServer" -Level "OK"
}

#region ============================================================
#       MODUS: EXPORT
#endregion =========================================================

if ($Mode -eq "Export") {
    
    # Parameter-Validierung
    if (-not $AltSqlServer -or -not $AltDatabase) {
        throw "Fuer den Export-Modus muessen -AltSqlServer und -AltDatabase angegeben werden."
    }
    
    Write-Log "Starte Export aus: $AltSqlServer\$AltDatabase"
    
    $Query = @"
SELECT 
    sys.Name0                           AS DeviceName,
    uda.UniqueUserName                  AS PrimaryUser,
    uda.RelationshipResourceID          AS RelationshipID,
    uda.CreationTime                    AS UDA_CreatedOn,
    CASE
        WHEN uda.Sources & 16 > 0 THEN 'Fast Install'
        WHEN uda.Sources & 8  > 0 THEN 'Windows Logon'
        WHEN uda.Sources & 4  > 0 THEN 'Usage Agent'
        WHEN uda.Sources & 2  > 0 THEN 'Administrator'
        WHEN uda.Sources & 1  > 0 THEN 'Software Catalog'
        ELSE CAST(uda.Sources AS VARCHAR)
    END                                 AS UDA_Source,
    uda.Sources                         AS UDA_SourceRaw
FROM v_UserMachineRelationship uda
JOIN v_R_System sys ON uda.MachineResourceID = sys.ResourceID
ORDER BY sys.Name0, uda.UniqueUserName
"@

    try {
        # SqlServer-Modul laden (falls verfuegbar)
        if (-not (Get-Module -Name SqlServer -ListAvailable -ErrorAction SilentlyContinue)) {
            Write-Log "SqlServer-Modul nicht gefunden, versuche SQLPS..." -Level "WARN"
            Import-Module SQLPS -DisableNameChecking -ErrorAction Stop
        } else {
            Import-Module SqlServer -ErrorAction Stop
        }
        
        $Results = Invoke-Sqlcmd -ServerInstance $AltSqlServer -Database $AltDatabase -Query $Query -QueryTimeout 120
        
        if (@($Results).Count -eq 0) {
            Write-Log "Keine aktiven UDAs gefunden!" -Level "WARN"
        } else {
            Write-Log "Gefunden: $($Results.Count) aktive UDA-Zuordnungen" -Level "OK"
            
            # Status-Spalte hinzufuegen fuer Import-Tracking
            $ExportData = $Results | Select-Object `
                DeviceName, `
                PrimaryUser, `
                UDA_Source, `
                UDA_SourceRaw, `
                UDA_CreatedOn, `
                @{Name = "ImportStatus";    Expression = { "Ausstehend" }}, `
                @{Name = "ImportTimestamp";  Expression = { "" }}, `
                @{Name = "ImportMessage";    Expression = { "" }}
            
            $ExportData | Export-Csv -Path $ExportCsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            Write-Log "Export gespeichert: $ExportCsvPath" -Level "OK"
            
            # Zusammenfassung
            Write-Host ""
            Write-Host "============================================================" -ForegroundColor Cyan
            Write-Host "  EXPORT ZUSAMMENFASSUNG" -ForegroundColor Cyan
            Write-Host "============================================================" -ForegroundColor Cyan
            Write-Host ""
            
            $BySource = $ExportData | Group-Object UDA_Source
            foreach ($grp in $BySource) {
                Write-Host "  $($grp.Name): $($grp.Count) Zuordnungen" -ForegroundColor White
            }
            
            $UniqueDevices = ($ExportData | Select-Object -ExpandProperty DeviceName -Unique).Count
            $UniqueUsers   = ($ExportData | Select-Object -ExpandProperty PrimaryUser -Unique).Count
            
            Write-Host ""
            Write-Host "  Geraete (eindeutig):  $UniqueDevices" -ForegroundColor Green
            Write-Host "  Benutzer (eindeutig): $UniqueUsers" -ForegroundColor Green
            Write-Host "  UDA-Zuordnungen:      $($ExportData.Count)" -ForegroundColor Green
            Write-Host ""
        }
        
    } catch {
        Write-Log "SQL-Fehler: $_" -Level "ERROR"
        throw
    }
}

#region ============================================================
#       MODUS: IMPORT
#endregion =========================================================

if ($Mode -eq "Import") {
    
    # Parameter-Validierung
    if (-not $NeuSiteCode -or -not $NeuSiteServer) {
        throw "Fuer den Import-Modus muessen -NeuSiteCode und -NeuSiteServer angegeben werden."
    }
    if (-not (Test-Path $ExportCsvPath)) {
        throw "Export-CSV nicht gefunden: $ExportCsvPath"
    }
    
    # CSV laden
    $UDAData = Import-Csv -Path $ExportCsvPath -Encoding UTF8 -Delimiter ";"
    Write-Log "CSV geladen: $($UDAData.Count) Eintraege"
    
    # Bereits importierte ueberspringen
    $Ausstehend = $UDAData | Where-Object { $_.ImportStatus -ne "OK" }
    $BereitsOK   = $UDAData | Where-Object { $_.ImportStatus -eq "OK" }
    
    if ($BereitsOK.Count -gt 0) {
        Write-Log "Ueberspringe $($BereitsOK.Count) bereits importierte UDAs" -Level "INFO"
    }
    Write-Log "Zu verarbeiten: $($Ausstehend.Count) UDAs"
    
    if ($Ausstehend.Count -eq 0) {
        Write-Log "Alle UDAs wurden bereits importiert. Nichts zu tun." -Level "OK"
        Stop-Transcript
        return
    }
    
    # Verbindung zum neuen SCCM
    $OriginalLocation = Get-Location
    try {
        Connect-NeuSCCM -SiteCode $NeuSiteCode -SiteServer $NeuSiteServer

        # Zaehler
        $CountOK        = 0
        $CountNotFound   = 0
        $CountAlready    = 0
        $CountError      = 0
        $Total           = $Ausstehend.Count
        $Current         = 0

        foreach ($Entry in $Ausstehend) {
            $Current++
            $Percent = [math]::Round(($Current / $Total) * 100, 1)
            Write-Progress -Activity "UDA Import" -Status "$Current von $Total ($Percent%)" -PercentComplete $Percent

            try {
                # Geraet im neuen SCCM suchen
                $Device = Get-CMDevice -Name $Entry.DeviceName -Fast -ErrorAction SilentlyContinue

                if (-not $Device) {
                    # Geraet noch nicht im neuen SCCM
                    $Entry.ImportStatus   = "GeraetNichtGefunden"
                    $Entry.ImportTimestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    $Entry.ImportMessage  = "Geraet noch nicht im neuen SCCM registriert"
                    $CountNotFound++
                    Write-Log "Geraet nicht gefunden: $($Entry.DeviceName) - wird beim naechsten Lauf erneut versucht" -Level "WARN"
                    continue
                }

                # Pruefen ob UDA bereits gesetzt ist
                $ExistingUDA = Get-CMUserDeviceAffinity -DeviceId $Device.ResourceID -ErrorAction SilentlyContinue |
                               Where-Object { $_.UniqueUserName -eq $Entry.PrimaryUser }

                if ($ExistingUDA) {
                    $Entry.ImportStatus   = "OK"
                    $Entry.ImportTimestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    $Entry.ImportMessage  = "UDA war bereits vorhanden"
                    $CountAlready++
                    Write-Log "Bereits vorhanden: $($Entry.DeviceName) -> $($Entry.PrimaryUser)" -Level "INFO"
                    continue
                }

                # UDA setzen
                Add-CMUserAffinityToDevice -DeviceId $Device.ResourceID -UserName $Entry.PrimaryUser

                $Entry.ImportStatus   = "OK"
                $Entry.ImportTimestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                $Entry.ImportMessage  = "Erfolgreich gesetzt"
                $CountOK++
                Write-Log "OK: $($Entry.DeviceName) -> $($Entry.PrimaryUser)" -Level "OK"

            } catch {
                $Entry.ImportStatus   = "Fehler"
                $Entry.ImportTimestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                $Entry.ImportMessage  = $_.Exception.Message
                $CountError++
                Write-Log "FEHLER bei $($Entry.DeviceName): $_" -Level "ERROR"
            }
        }

        Write-Progress -Activity "UDA Import" -Completed

        # CSV aktualisieren (Status zurueckschreiben)
        $UDAData | Export-Csv -Path $ExportCsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";" -Force
        Write-Log "CSV aktualisiert: $ExportCsvPath" -Level "OK"

        # Bericht erstellen - nur nicht-OK Eintraege
        $Pending = $UDAData | Where-Object { $_.ImportStatus -ne "OK" }
        if ($Pending.Count -gt 0) {
            $Pending | Export-Csv -Path $ReportCsv -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            Write-Log "Ausstehende UDAs exportiert: $ReportCsv" -Level "INFO"
        }
    } finally {
        Set-Location $OriginalLocation
    }
    
    # Zusammenfassung
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  IMPORT ZUSAMMENFASSUNG" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Neu gesetzt:            $CountOK" -ForegroundColor Green
    Write-Host "  Bereits vorhanden:      $CountAlready" -ForegroundColor Gray
    Write-Host "  Geraet nicht gefunden:  $CountNotFound" -ForegroundColor Yellow
    Write-Host "  Fehler:                 $CountError" -ForegroundColor Red
    Write-Host ""
    
    $GesamtOK = ($UDAData | Where-Object { $_.ImportStatus -eq "OK" }).Count
    $GesamtTotal = $UDAData.Count
    Write-Host "  Gesamtfortschritt:      $GesamtOK / $GesamtTotal ($([math]::Round(($GesamtOK / $GesamtTotal) * 100, 1))%)" -ForegroundColor Cyan
    Write-Host ""
    
    if ($CountNotFound -gt 0) {
        Write-Host "  HINWEIS: $CountNotFound Geraete noch nicht im neuen SCCM." -ForegroundColor Yellow
        Write-Host "  Skript nach dem naechsten Onboarding-Batch erneut ausfuehren." -ForegroundColor Yellow
        Write-Host ""
    }
}

#region ============================================================
#       MODUS: REPORT
#endregion =========================================================

if ($Mode -eq "Report") {
    
    if (-not (Test-Path $ExportCsvPath)) {
        throw "Export-CSV nicht gefunden: $ExportCsvPath"
    }
    
    $UDAData = Import-Csv -Path $ExportCsvPath -Encoding UTF8 -Delimiter ";"
    
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  UDA MIGRATION - STATUSBERICHT" -ForegroundColor Cyan
    Write-Host "  CSV: $ExportCsvPath" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $StatusGroups = $UDAData | Group-Object ImportStatus | Sort-Object Name
    
    foreach ($grp in $StatusGroups) {
        $color = switch ($grp.Name) {
            "OK"                  { "Green" }
            "Ausstehend"          { "Yellow" }
            "GeraetNichtGefunden" { "Yellow" }
            "Fehler"              { "Red" }
            default               { "White" }
        }
        Write-Host "  $($grp.Name): $($grp.Count)" -ForegroundColor $color
    }
    
    $TotalCount = $UDAData.Count
    $OKCount    = ($UDAData | Where-Object { $_.ImportStatus -eq "OK" }).Count
    $Percent    = if ($TotalCount -gt 0) { [math]::Round(($OKCount / $TotalCount) * 100, 1) } else { 0 }
    
    Write-Host ""
    Write-Host "  Gesamt:        $TotalCount UDAs" -ForegroundColor White
    Write-Host "  Fortschritt:   $OKCount / $TotalCount ($Percent%)" -ForegroundColor Cyan
    Write-Host ""
    
    # Optional: Geraete-Liste die noch ausstehen
    if ($NeuSiteCode -and $NeuSiteServer) {
        Write-Host "  Pruefe aktuelle Verfuegbarkeit im neuen SCCM..." -ForegroundColor Gray
        
        $OriginalLocation = Get-Location
        try {
            Connect-NeuSCCM -SiteCode $NeuSiteCode -SiteServer $NeuSiteServer

            $Pending = $UDAData | Where-Object { $_.ImportStatus -ne "OK" }
            $NowAvailable = 0

            foreach ($Entry in $Pending) {
                $Device = Get-CMDevice -Name $Entry.DeviceName -Fast -ErrorAction SilentlyContinue
                if ($Device) {
                    $NowAvailable++
                    Write-Host "  BEREIT: $($Entry.DeviceName) -> $($Entry.PrimaryUser)" -ForegroundColor Green
                }
            }
        } finally {
            Set-Location $OriginalLocation
        }
        
        if ($NowAvailable -gt 0) {
            Write-Host ""
            Write-Host "  $NowAvailable Geraete sind jetzt verfuegbar und koennen importiert werden!" -ForegroundColor Green
            Write-Host "  Fuehre das Skript mit -Mode Import erneut aus." -ForegroundColor Green
        } else {
            Write-Host ""
            Write-Host "  Keine neuen Geraete verfuegbar." -ForegroundColor Yellow
        }
        Write-Host ""
    }
}

#region ============================================================
#       AUFRAUMEN
#endregion =========================================================

} finally {
    Stop-Transcript
    Write-Host "Log gespeichert: $LogFile" -ForegroundColor Gray
}
