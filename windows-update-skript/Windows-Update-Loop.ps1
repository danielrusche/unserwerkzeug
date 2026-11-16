<#
.SYNOPSIS
    Installiert nur die normalen (empfohlenen) Windows-Updates, startet einmal neu, wartet nach
    dem Neustart 90 Sekunden und sucht erneut – bis keine gefilterten Updates mehr übrig sind.

.DESCRIPTION
    Ablauf:
      1. Suche nach Updates mit dem Windows-Update-Agenten (Microsoft.Update.Session), gefiltert
         nach den Regeln unten.
      2. Gefundene Updates herunterladen und installieren (eine Installationsrunde für alle).
      3. Erst nach Abschluss aller Installationen einmal neu starten.
      4. Der Neustart-Zustand liegt in einer Zustandsdatei; ein geplanter Task (SYSTEM, beim
         Systemstart) setzt den Ablauf nach dem Neustart fort und wartet dort 90 Sekunden.
      5. Erneute Suche. Keine gefilterten Updates mehr -> Fertig; Task, Zustandsdatei und
         Kopie des Skripts werden entfernt. Sonst beginnt der Ablauf wieder bei Schritt 1.

    FILTERKRITERIUM (bitte prüfen):
      Suchabfrage an Windows Update:
          IsInstalled=0 and Type='Software' and IsHidden=0
      Zusätzlich pro Update in der Auswahl:
          AutoSelectOnWebSites -eq $true
      Das Flag AutoSelectOnWebSites ist das Merkmal, das Windows für "empfohlene" Updates setzt
      (sie werden in der Windows-Update-Suche ohne den Bereich "Optionale Updates" angeboten).
      Optionale Updates (AutoSelectOnWebSites=false), Treiber (Type='Driver'), ausgeblendete
      Updates (IsHidden=1) und bereits installierte Updates werden NICHT installiert.
      Ob die Auswahl mit der Anzeige der Windows-Einstellungen exakt übereinstimmt, sollte auf
      einem Testrechner mit dem Parameter -DryRun verglichen werden.

    Parameter:
      -DryRun      Nur suchen und die gefilterte Liste protokollieren. Nichts installieren, kein
                   Neustart, kein Task.
      -MaxCycles   Sicherheitsgrenze für die Anzahl Durchläufe (Standard 10), damit ein
                   dauerhaft fehlschlagendes Update keine Endlosschleife erzeugt.

    Logfile: C:\ProgramData\WinUpdateLoop\update.log (bleibt nach dem Lauf erhalten).
    Aufräumen am Ende: Task "WinUpdateLoop-Resume", state.json und die Skriptkopie werden entfernt.

    Aufruf (PowerShell als Administrator):
      powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Windows-Update-Loop.ps1 -DryRun
      powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Windows-Update-Loop.ps1

.NOTES
    UNGETESTET: Dieses Skript wurde auf einem Linux-Rechner ohne Windows-Umgebung geschrieben
    und dort nicht ausgeführt. Vor dem produktiven Einsatz bitte zuerst mit -DryRun auf einem
    Windows-Testrechner (Windows PowerShell 5.1 oder PowerShell 7 unter Windows) prüfen.
#>

[CmdletBinding()]
param(
    [switch]$DryRun,
    [switch]$Resume,
    [int]$MaxCycles = 10
)

$ErrorActionPreference = 'Stop'

$WorkDir     = Join-Path $env:ProgramData 'WinUpdateLoop'
$LogFile     = Join-Path $WorkDir 'update.log'
$StateFile   = Join-Path $WorkDir 'state.json'
$ScriptCopy  = Join-Path $WorkDir 'Windows-Update-Loop.ps1'
$TaskName    = 'WinUpdateLoop-Resume'
$WaitSeconds = 90
$SearchCriteria = "IsInstalled=0 and Type='Software' and IsHidden=0"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    if (-not (Test-Path $WorkDir)) { New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null }
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal $id
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-FilteredUpdates {
    $session  = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $result   = $searcher.Search($SearchCriteria)

    $selected = New-Object -ComObject Microsoft.Update.UpdateColl
    $skipped  = New-Object System.Collections.ArrayList

    foreach ($u in $result.Updates) {
        if ($u.Type -ne 1) {
            [void]$skipped.Add("$($u.Title) (kein Software-Update, Typ $($u.Type))")
            continue
        }
        if ($u.IsHidden) {
            [void]$skipped.Add("$($u.Title) (ausgeblendet)")
            continue
        }
        if (-not $u.AutoSelectOnWebSites) {
            [void]$skipped.Add("$($u.Title) (optional, nicht empfohlen)")
            continue
        }
        [void]$selected.Add($u)
    }

    return @{
        Session  = $session
        Selected = $selected
        Skipped  = $skipped
        Total    = $result.Updates.Count
    }
}

function Install-Updates {
    param($Session, $Updates)

    foreach ($u in $Updates) {
        if (-not $u.EulaAccepted) { $u.AcceptEula() }
    }

    Write-Log "Lade $($Updates.Count) Update(s) herunter..."
    $downloader = $Session.CreateUpdateDownloader()
    $downloader.Updates = $Updates
    $dl = $downloader.Download()
    # OperationResultCode: 2 = erfolgreich
    if ($dl.ResultCode -ne 2) { throw "Download fehlgeschlagen (ResultCode $($dl.ResultCode))." }

    Write-Log "Installiere $($Updates.Count) Update(s)..."
    $installer = $Session.CreateUpdateInstaller()
    $installer.Updates = $Updates
    $inst = $installer.Install()

    $ok = 0
    for ($i = 0; $i -lt $Updates.Count; $i++) {
        $r = $inst.GetUpdateResult($i)
        $title = $Updates.Item($i).Title
        # 2 = erfolgreich, 3 = mit Fehlern, 4 = fehlgeschlagen, 5 = abgebrochen
        Write-Log "  Ergebnis [$($r.ResultCode)] $title (HResult $($r.HResult))"
        if ($r.ResultCode -eq 2) { $ok++ }
    }

    return @{
        ResultCode     = $inst.ResultCode
        RebootRequired = [bool]$inst.RebootRequired
        Succeeded      = $ok
        Count          = $Updates.Count
    }
}

function Register-ResumeTask {
    if ($PSCommandPath -ne $ScriptCopy) {
        Copy-Item -Path $PSCommandPath -Destination $ScriptCopy -Force
    }
    $action    = New-ScheduledTaskAction -Execute 'powershell.exe' `
                    -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowsStyle Hidden -File "{0}" -Resume' -f $ScriptCopy)
    $trigger   = New-ScheduledTaskTrigger -AtStartup
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                    -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 12)
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
        -Principal $principal -Settings $settings -Force | Out-Null
    Write-Log "Geplanter Task '$TaskName' für die Fortsetzung nach dem Neustart angelegt."
}

function Remove-Leftovers {
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    Remove-Item -Path $StateFile -Force -ErrorAction SilentlyContinue
    # Skriptkopie im Arbeitsordner entfernen; das Logfile bleibt bewusst erhalten.
    Remove-Item -Path $ScriptCopy -Force -ErrorAction SilentlyContinue
    Write-Log "Aufgeräumt: Task, Zustandsdatei und Skriptkopie entfernt."
}

function Invoke-Cycle {
    param([int]$Cycle)

    Write-Log "=== Durchlauf $Cycle: Suche (Filter: $SearchCriteria, nur AutoSelectOnWebSites=true) ==="
    $found = Get-FilteredUpdates

    foreach ($s in $found.Skipped) { Write-Log "  Ausgeschlossen: $s" }
    Write-Log "Gefunden insgesamt: $($found.Total); nach Filter zur Installation: $($found.Selected.Count)"
    foreach ($u in $found.Selected) { Write-Log "  Kandidat: $($u.Title)" }

    if ($found.Selected.Count -eq 0) { return 'Done' }

    if ($DryRun) {
        Write-Log 'Trockenlauf: Es wird nichts installiert.'
        return 'DryRun'
    }

    $res = Install-Updates -Session $found.Session -Updates $found.Selected
    Write-Log "Installation abgeschlossen: $($res.Succeeded) von $($res.Count) erfolgreich; Windows meldet Neustart nötig: $($res.RebootRequired)."

    # Ein einziger Neustart, erst nach Abschluss aller Installationen dieses Durchlaufs.
    @{ Cycle = $Cycle; Started = (Get-Date).ToString('o') } |
        ConvertTo-Json | Set-Content -Path $StateFile -Encoding UTF8

    Register-ResumeTask
    Write-Log 'Neustart in 30 Sekunden.'
    Start-Sleep -Seconds 30
    Restart-Computer -Force
    return 'Rebooting'
}

# ---------------- Hauptablauf ----------------
try {
    if (-not (Test-Admin)) { throw 'Das Skript muss als Administrator ausgeführt werden.' }
    if (-not (Test-Path $WorkDir)) { New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null }

    if ($Resume -and -not $DryRun) {
        if (-not (Test-Path $StateFile)) {
            Write-Log 'Resume ohne Zustandsdatei – nichts zu tun.'
            Remove-Leftovers
            exit 0
        }
        $state = Get-Content -Path $StateFile -Raw | ConvertFrom-Json
        Write-Log "Nach Neustart: warte $WaitSeconds Sekunden, bevor erneut gesucht wird."
        Start-Sleep -Seconds $WaitSeconds
        $cycle = [int]$state.Cycle + 1
    }
    else {
        $cycle = 1
    }

    if ($cycle -gt $MaxCycles) {
        throw "Sicherheitsgrenze von $MaxCycles Durchläufen erreicht. Abbruch ohne weiteren Neustart."
    }

    $outcome = Invoke-Cycle -Cycle $cycle

    if ($outcome -eq 'Done') {
        Write-Log 'Keine weiteren gefilterten Updates. Skript ist fertig.'
        Remove-Leftovers
    }
}
catch {
    Write-Log "FEHLER: $($_.Exception.Message)" 'ERROR'
    Write-Log "Details: $($_.ScriptStackTrace)" 'ERROR'
    if (-not $DryRun) { Remove-Leftovers }
    exit 1
}
