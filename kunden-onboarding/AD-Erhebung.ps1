<#
.SYNOPSIS
    Liest Active Directory NUR LESEND aus und erstellt CSV-Dateien sowie eine
    Markdown-Übersicht, deren Abschnitte den IT-Onboarding-Fragebogen spiegeln
    (Abschnitte 2 Standorte, 3 Nutzer, 4.2 Domäne/Admin-Gruppen, 5 Server, 8 Endgeräte).
    Zusätzlich entsteht AD-Uebersicht.html: eine übersichtliche, durchsuchbare Darstellung
    mit Zusammenfassung je Bereich. Sie braucht keine externen Dateien oder Module.
    Außerdem entstehen zwei Office-Dateien im Ausgabeordner:
    - Kunden-Onboarding-AD-<Domäne>.xlsx: Arbeitsmappe aus der Onboarding-Vorlage (Tabs 1 bis 11,
      leer) plus die AD-Tabs „AD Kennzahlen“, „AD Domäne“, „AD Domain Controller“, „AD Standorte“,
      „AD Subnetze“, „AD Nutzer“, „AD Computer“, „AD Gruppen“, „AD Privilegierte“.
    - AD-Zusammenfassung-<Domäne>.docx: Word-Zusammenfassung mit Ansprechpartnern, Standorten,
      Nutzer- und Systemangaben, offenen Punkten und AD-Kennzahlen.
    Die Vorlage ist im Skript eingebettet. Excel und Word werden ohne Office und ohne Module
    direkt als Dateipakete geschrieben. Eine vorhandene Datei wird nie überschrieben.

.DESCRIPTION
    Erhoben werden: Domäne, Domain Controller, AD-Standorte (Sites) und Subnetze,
    Organisationseinheiten, Nutzerkonten (Name, Anmeldename, Abteilung, Standort, Status),
    Computerkonten (Name, Betriebssystem, letzte Anmeldung) und Gruppen sowie die
    direkten Mitglieder der privilegierten Gruppen (nur Kontonamen).

    Nicht erhoben: Passwörter, Passworthashes, LAPS-/Recovery-Schlüssel, BitLocker-Schlüssel
    oder sonstige Geheimnisse. Es werden ausschließlich Get-Cmdlets verwendet; im Netzwerk
    oder im Verzeichnis wird nichts geändert.

    Voraussetzungen:
    - PowerShell 5.1 (Windows PowerShell) oder PowerShell 7 unter Windows.
    - ActiveDirectory-Modul (Teil der RSAT-Werkzeuge). Ohne Modul laufen nur die Teile
      ohne AD-Abfragen; alle AD-Abschnitte werden im Fehlerprotokoll als übersprungen vermerkt.
      Installation unter Windows 10/11 (Administrator):
        Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
      Auf einem Domain Controller/Server: Install-WindowsFeature RSAT-AD-PowerShell
    - Für die Excel- und Word-Datei ist keine weitere Software nötig (kein Office, kein openpyxl,
      kein Modul). Die Excel-Vorlage ist im Skript eingebettet.
    - Netzwerkzugriff auf einen Domain Controller (LDAP 389 / ADWS 9389, Kerberos 88).
    - Berechtigung: Die übergebenen Anmeldedaten brauchen Leserechte im AD. Für vollständige
      Gruppen- und Mitgliederdaten empfiehlt sich ein Domain-Admin-Konto, das der Kunde
      für die Erhebung bereitstellt. Die Konfigurationspartition (AD-Sites) ist für
      authentifizierte Benutzer lesbar.

.PARAMETER Domain
    DNS-Name der Domäne, z. B. firma.local.

.PARAMETER OutputDir
    Ausgabeordner. Standard: .\AD-Erhebung-<Domäne>-<Zeitstempel> im aktuellen Verzeichnis.

.PARAMETER Server
    Optional: bestimmter Domain Controller. Standard: die Domäne (DNS löst auf einen DC auf).

.PARAMETER InactiveDays
    Schwelle in Tagen, ab der Konten und Computer als inaktiv gelten. Standard: 90.

.PARAMETER TemplatePath
    Optional: eigene Excel-Vorlage (.xlsx) statt der eingebetteten Vorlage.

.EXAMPLE
    .\AD-Erhebung.ps1 -Domain firma.local
.EXAMPLE
    .\AD-Erhebung.ps1 -Domain firma.local -Server dc01.firma.local -OutputDir D:\Kunden\Firma\AD
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Domain,

    [string]$OutputDir,

    [string]$Server,

    [ValidateRange(1, 3650)]
    [int]$InactiveDays = 90,

    [string]$TemplatePath
)

$ErrorActionPreference = 'Stop'
$RunStart = Get-Date
$Target   = if ($Server) { $Server } else { $Domain }
$cutoff   = $RunStart.AddDays(-$InactiveDays)

$script:Problems = New-Object System.Collections.Generic.List[string]
$script:Ok       = @{}
$script:Md       = New-Object System.Text.StringBuilder
$script:Data     = @{
    Domain      = $null
    DCs         = @()
    Sites       = @()
    Subnets     = @()
    OUs         = @()
    Users       = @()
    Computers   = @()
    Groups      = @()
    PrivSummary = @()
    PrivDetail  = @()
}

# ---------------------------------------------------------------- Hilfsfunktionen

function YesNo { param($Flag) if ($Flag) { 'ja' } else { 'nein' } }

function Add-Problem {
    param([string]$Section, [string]$Message)
    $line = '[{0}] {1}' -f $Section, $Message
    $script:Problems.Add($line)
    Write-Warning $line
}

function Get-CnFromDn {
    param([string]$Dn)
    if ($Dn -match '^CN=([^,]+)') { return $Matches[1] }
    return $Dn
}

function Get-ParentCn {
    # CanonicalName "firma.local/Berlin/Vertrieb/Max" -> "firma.local/Berlin/Vertrieb"
    param([string]$CanonicalName)
    if (-not $CanonicalName) { return $null }
    $idx = $CanonicalName.LastIndexOf('/')
    if ($idx -lt 0) { return $CanonicalName }
    return $CanonicalName.Substring(0, $idx)
}

function Export-Result {
    param([string]$FileName, [object[]]$Rows)
    if (-not $Rows -or @($Rows).Count -eq 0) { return }
    $path = Join-Path $OutputDir $FileName
    @($Rows) | Export-Csv -LiteralPath $path -NoTypeInformation -Delimiter ';' -Encoding UTF8
    Write-Host "    CSV geschrieben: $path"
}

function Invoke-Section {
    # Ein Abschnitt darf fehlschlagen, ohne den Lauf zu beenden.
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$Action,
        [switch]$NeedsAD
    )
    if ($NeedsAD -and -not $script:AdAvailable) {
        Add-Problem $Name 'übersprungen: ActiveDirectory-Modul fehlt (RSAT installieren, siehe Hilfe).'
        return
    }
    Write-Host "==> $Name"
    try {
        & $Action
        $script:Ok[$Name] = $true
    }
    catch {
        Add-Problem $Name $_.Exception.Message
    }
}

function Add-Md {
    param([string]$Text = '')
    [void]$script:Md.AppendLine($Text)
}

function Format-Cell {
    param($Value)
    if ($null -eq $Value) { return 'unbekannt' }
    if ($Value -is [string] -and $Value.Trim() -eq '') { return 'unbekannt' }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd') }
    $text = [string]$Value
    $text = $text -replace '\|', '/'
    $text = $text -replace '[\r\n]+', ' '
    return $text.Trim()
}

function ConvertTo-MdTable {
    param([object[]]$Rows)
    $rows = @($Rows)
    if ($rows.Count -eq 0) { return "_Keine Einträge._`n" }
    $columns = @($rows[0].PSObject.Properties.Name)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('| ' + ($columns -join ' | ') + ' |')
    [void]$sb.AppendLine('|' + (($columns | ForEach-Object { '---' }) -join '|') + '|')
    foreach ($r in $rows) {
        $cells = foreach ($c in $columns) { Format-Cell $r.$c }
        [void]$sb.AppendLine('| ' + ($cells -join ' | ') + ' |')
    }
    return $sb.ToString()
}

function Add-Block {
    # Schreibt einen Markdown-Block nur, wenn der zugehörige Abschnitt erfolgreich war.
    param([string]$Key, [scriptblock]$Body)
    if (-not $script:Ok[$Key]) {
        Add-Md '_Nicht erhoben, Details im Fehlerprotokoll._'
        Add-Md ''
        return
    }
    & $Body
    Add-Md ''
}

# ---------------------------------------------------------------- Vorbereitung

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $script:AdAvailable = $true
}
catch {
    $script:AdAvailable = $false
    Write-Warning 'ActiveDirectory-Modul nicht verfügbar. AD-Abschnitte werden übersprungen.'
}

$Cred = Get-Credential -Message ('Lesender Domain-Admin-Zugang für Domäne {0}' -f $Domain)
if ($null -eq $Cred) {
    Write-Error 'Abbruch: Es wurden keine Anmeldedaten angegeben. Es wurde nichts abgefragt.'
    return
}

if (-not $OutputDir) {
    $OutputDir = Join-Path (Get-Location) ('AD-Erhebung-{0}-{1}' -f $Domain, $RunStart.ToString('yyyyMMdd-HHmmss'))
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
$OutputDir = (Resolve-Path -LiteralPath $OutputDir).ProviderPath

# Alle AD-Abfragen laufen lesend mit denselben Anmeldedaten und Zielen.
$adArgs = @{ Server = $Target; Credential = $Cred; ErrorAction = 'Stop' }

# ---------------------------------------------------------------- 1. Domäne und DCs (Fragebogen 4.2, 5, 6.1)

Invoke-Section -Name 'Domäne und Domain Controller' -NeedsAD -Action {
    $d = Get-ADDomain @adArgs
    $f = Get-ADForest @adArgs
    $script:Data.Domain = [pscustomobject][ordered]@{
        'DNS-Name'              = $d.DNSRoot
        'NetBIOS-Name'          = $d.NetBIOSName
        'Domänenfunktionsebene' = $d.DomainMode
        'Gesamtstruktur'        = $f.Name
        'Gesamtstrukturebene'   = $f.ForestMode
        'Schema-Master'         = $f.SchemaMaster
        'Domänennamenmaster'    = $f.DomainNamingMaster
        'PDC-Emulator'          = $d.PDCEmulator
        'RID-Master'            = $d.RIDMaster
        'Infrastruktur-Master'  = $d.InfrastructureMaster
        'Abgefragt über'        = $Target
    }

    $dcs = Get-ADDomainController -Filter * @adArgs
    $script:Data.DCs = @($dcs | ForEach-Object {
        [pscustomobject][ordered]@{
            'Name'                     = $_.Name
            'Hostname'                 = $_.HostName
            'IPv4'                     = $_.IPv4Address
            'AD-Standort'              = $_.Site
            'Betriebssystem'           = $_.OperatingSystem
            'Globaler Katalog'         = $(YesNo $_.IsGlobalCatalog)
            'Schreibgeschützt (RODC)'  = $(YesNo $_.IsReadOnly)
        }
    })
    Export-Result 'domaene.csv' @($script:Data.Domain)
    Export-Result 'domain-controller.csv' $script:Data.DCs
}

# ---------------------------------------------------------------- 2. AD-Sites und Subnetze (Fragebogen 2, 6.2)

Invoke-Section -Name 'AD-Standorte und Subnetze' -NeedsAD -Action {
    $sites = Get-ADReplicationSite -Filter * -Properties Description @adArgs
    $script:Data.Sites = @($sites | Sort-Object Name | ForEach-Object {
        [pscustomobject][ordered]@{
            'Standort (AD-Site)' = $_.Name
            'Beschreibung'       = $_.Description
        }
    })

    $subnets = Get-ADReplicationSubnet -Filter * -Properties Description, Location @adArgs
    $script:Data.Subnets = @($subnets | Sort-Object Name | ForEach-Object {
        [pscustomobject][ordered]@{
            'Subnetz'      = $_.Name
            'AD-Standort'  = Get-CnFromDn $_.Site
            'Ort'          = $_.Location
            'Beschreibung' = $_.Description
        }
    })
    Export-Result 'standorte-ad-sites.csv' $script:Data.Sites
    Export-Result 'subnetze.csv' $script:Data.Subnets
}

# ---------------------------------------------------------------- 3. Organisationseinheiten

Invoke-Section -Name 'Organisationseinheiten' -NeedsAD -Action {
    $ous = Get-ADOrganizationalUnit -Filter * -Properties Description, CanonicalName @adArgs
    $script:Data.OUs = @($ous | Sort-Object CanonicalName | ForEach-Object {
        [pscustomobject][ordered]@{
            'Pfad'         = $_.CanonicalName
            'Name'         = $_.Name
            'Beschreibung' = $_.Description
        }
    })
    Export-Result 'organisationseinheiten.csv' $script:Data.OUs
}

# ---------------------------------------------------------------- 4. Nutzerkonten (Fragebogen 3)
# Keine Passwort- oder Sicherheitsattribute werden angefordert.

Invoke-Section -Name 'Nutzerkonten' -NeedsAD -Action {
    $users = Get-ADUser -Filter * -Properties DisplayName, SamAccountName, UserPrincipalName, EmailAddress, Department, Title, Office, Enabled, LastLogonDate, whenCreated, CanonicalName @adArgs
    $script:Data.Users = @($users | Sort-Object SamAccountName | ForEach-Object {
        $inactive = (-not $_.LastLogonDate) -or ($_.LastLogonDate -lt $cutoff)
        $displayName = if ($_.DisplayName) { $_.DisplayName } else { $_.Name }
        [pscustomobject][ordered]@{
            'Name'                    = $displayName
            'Anmeldename'             = $_.SamAccountName
            'UPN'                     = $_.UserPrincipalName
            'E-Mail'                  = $_.EmailAddress
            'Abteilung'               = $_.Department
            'Funktion'                = $_.Title
            'Standort'                = $_.Office
            'Status'                  = $(if ($_.Enabled) { 'aktiv' } else { 'deaktiviert' })
            'Letzte Anmeldung'        = $_.LastLogonDate
            'Inaktiv'                 = $(YesNo $inactive)
            'Erstellt'                = $_.whenCreated
            'OU'                      = Get-ParentCn $_.CanonicalName
        }
    })
    Export-Result 'nutzer.csv' $script:Data.Users
}

# ---------------------------------------------------------------- 5. Computerkonten (Fragebogen 5, 8)

Invoke-Section -Name 'Computerkonten' -NeedsAD -Action {
    $computers = Get-ADComputer -Filter * -Properties DNSHostName, OperatingSystem, OperatingSystemVersion, LastLogonDate, IPv4Address, Enabled, Description, whenCreated, CanonicalName @adArgs
    $script:Data.Computers = @($computers | Sort-Object Name | ForEach-Object {
        $inactive = (-not $_.LastLogonDate) -or ($_.LastLogonDate -lt $cutoff)
        [pscustomobject][ordered]@{
            'Name'              = $_.Name
            'DNS-Hostname'      = $_.DNSHostName
            'Betriebssystem'    = $_.OperatingSystem
            'OS-Version'        = $_.OperatingSystemVersion
            'Letzte Anmeldung'  = $_.LastLogonDate
            'Status'            = $(if ($_.Enabled) { 'aktiv' } else { 'deaktiviert' })
            'Inaktiv'           = $(YesNo $inactive)
            'IPv4'              = $_.IPv4Address
            'Typ'               = $(if ($_.OperatingSystem -like '*Server*') { 'Server' } else { 'Client' })
            'OU'                = Get-ParentCn $_.CanonicalName
            'Beschreibung'      = $_.Description
            'Erstellt'          = $_.whenCreated
        }
    })
    Export-Result 'computerkonten.csv' $script:Data.Computers
}

# ---------------------------------------------------------------- 6. Gruppen (Fragebogen 4.2)

Invoke-Section -Name 'Gruppen' -NeedsAD -Action {
    $groups = Get-ADGroup -Filter * -Properties Description, Members, GroupScope, GroupCategory, CanonicalName @adArgs
    $script:Data.Groups = @($groups | Sort-Object Name | ForEach-Object {
        [pscustomobject][ordered]@{
            'Gruppe'                     = $_.Name
            'Bereich'                    = [string]$_.GroupScope
            'Typ'                        = [string]$_.GroupCategory
            'Anzahl Mitglieder (direkt)' = @($_.Members).Count
            'OU'                         = Get-ParentCn $_.CanonicalName
            'Beschreibung'               = $_.Description
        }
    })
    Export-Result 'gruppen.csv' $script:Data.Groups
}

# ---------------------------------------------------------------- 7. Privilegierte Gruppen (Fragebogen 4.2)
# Gruppen werden über well-known SIDs gesucht, damit die lokalisierten Namen egal sind.
# Es werden nur direkte Mitglieder und deren Kontonamen gelesen; keine Passwörter.

Invoke-Section -Name 'Privilegierte Gruppen' -NeedsAD -Action {
    $dom    = Get-ADDomain @adArgs
    $domSid = $dom.DomainSID.Value

    $targets = New-Object System.Collections.Generic.List[object]
    $targets.Add(@{ Id = "$domSid-512" })   # Domänen-Admins
    $targets.Add(@{ Id = "$domSid-518" })   # Schema-Admins
    $targets.Add(@{ Id = 'S-1-5-32-544' })  # Administratoren (eingebaut)
    $targets.Add(@{ Id = 'S-1-5-32-548' })  # Kontooperatoren
    $targets.Add(@{ Id = 'S-1-5-32-551' })  # Sicherungsoperatoren

    try {
        $rootDomain = (Get-ADForest @adArgs).RootDomain
        $rootSid = (Get-ADDomain -Identity $rootDomain @adArgs).DomainSID.Value
        $targets.Add(@{ Id = "$rootSid-519" })  # Organisations-Admins (Stammdomäne)
    }
    catch {
        Add-Problem 'Privilegierte Gruppen' ('Organisations-Admins nicht gelesen: ' + $_.Exception.Message)
    }

    $detail  = New-Object System.Collections.Generic.List[object]
    $summary = New-Object System.Collections.Generic.List[object]

    foreach ($t in $targets) {
        try {
            $g = Get-ADGroup -Identity $t.Id -Properties Members @adArgs
            $members = @(Get-ADGroupMember -Identity $g @adArgs)
            $names = @()
            foreach ($m in $members) {
                $detail.Add([pscustomobject][ordered]@{
                    'Gruppe'        = $g.Name
                    'Mitglied'      = $(if ($m.SamAccountName) { $m.SamAccountName } else { $m.Name })
                    'Objektart'     = $m.objectClass
                })
                $names += $(if ($m.SamAccountName) { $m.SamAccountName } else { $m.Name })
            }
            $summary.Add([pscustomobject][ordered]@{
                'Gruppe'                     = $g.Name
                'Anzahl Mitglieder (direkt)' = $members.Count
                'Mitglieder (Kontonamen)'    = ($names -join ', ')
            })
        }
        catch {
            Add-Problem 'Privilegierte Gruppen' ('Gruppe {0} nicht gelesen: {1}' -f $t.Id, $_.Exception.Message)
        }
    }

    $script:Data.PrivSummary = $summary.ToArray()
    $script:Data.PrivDetail  = $detail.ToArray()
    Export-Result 'privilegierte-gruppen-mitglieder.csv' $script:Data.PrivDetail
}

# ---------------------------------------------------------------- Markdown-Übersicht

function Build-Markdown {
    $u  = @($script:Data.Users)
    $c  = @($script:Data.Computers)
    $active   = @($u | Where-Object { $_.Status -eq 'aktiv' })
    $disabled = @($u | Where-Object { $_.Status -eq 'deaktiviert' })
    $inactiveActive = @($active | Where-Object { $_.Inaktiv -eq 'ja' })

    $cActive   = @($c | Where-Object { $_.Status -eq 'aktiv' })
    $cInactive = @($c | Where-Object { $_.Inaktiv -eq 'ja' })
    $servers   = @($c | Where-Object { $_.Typ -eq 'Server' })
    $clients   = @($c | Where-Object { $_.Typ -eq 'Client' })

    Add-Md "# AD-Erhebung: $Domain"
    Add-Md ''
    Add-Md ('Erhebungszeitpunkt: {0}' -f $RunStart.ToString('yyyy-MM-dd HH:mm zzz'))
    Add-Md ''
    Add-Md 'Lesende Erhebung aus Active Directory. Es wurden keine Passwörter, Hashes oder sonstigen Geheimnisse abgefragt. Die Werte „Letzte Anmeldung“ sind replizierte Näherungswerte (Abweichung bis etwa 14 Tage). Felder, die im AD nicht gepflegt sind, stehen als „unbekannt“.'
    Add-Md ''

    Add-Md '## Domäne (Fragebogen 4.2)'
    Add-Md ''
    Add-Block 'Domäne und Domain Controller' {
        if ($script:Data.Domain) {
            $rows = @($script:Data.Domain.PSObject.Properties | ForEach-Object {
                [pscustomobject][ordered]@{ 'Eigenschaft' = $_.Name; 'Wert' = $_.Value }
            })
            Add-Md (ConvertTo-MdTable $rows)
        }
    }

    Add-Md '## Domain Controller (Fragebogen 5 und 6.1)'
    Add-Md ''
    Add-Block 'Domäne und Domain Controller' { Add-Md (ConvertTo-MdTable $script:Data.DCs) }

    Add-Md '## Standorte (Fragebogen 2)'
    Add-Md ''
    Add-Md 'AD-Standorte (Sites):'
    Add-Md ''
    Add-Block 'AD-Standorte und Subnetze' { Add-Md (ConvertTo-MdTable $script:Data.Sites) }
    Add-Md 'Subnetze und Zuordnung zu den Standorten (Fragebogen 6.2 und 6.3):'
    Add-Md ''
    Add-Block 'AD-Standorte und Subnetze' { Add-Md (ConvertTo-MdTable $script:Data.Subnets) }
    Add-Md 'Adressen, Ansprechpartner vor Ort und Internetanschlüsse sind im AD nicht hinterlegt und bleiben im Fragebogen offen.'
    Add-Md ''

    Add-Md '## Nutzer (Fragebogen 3)'
    Add-Md ''
    Add-Block 'Nutzerkonten' {
        Add-Md "- Nutzerkonten gesamt: $($u.Count)"
        Add-Md "- Aktiv: $($active.Count)"
        Add-Md "- Deaktiviert: $($disabled.Count)"
        Add-Md "- Aktiv, aber seit mehr als $InactiveDays Tagen nicht angemeldet bzw. nie angemeldet: $($inactiveActive.Count)"
        Add-Md ''
        Add-Md 'Abteilungen (aus dem Feld „Abteilung“):'
        Add-Md ''
        $deptRows = @($u | Group-Object -Property Abteilung | Sort-Object Name | ForEach-Object {
            $grp = @($_.Group)
            [pscustomobject][ordered]@{
                'Abteilung'        = $(if ($_.Name) { $_.Name } else { '(nicht gepflegt)' })
                'Nutzer gesamt'    = $grp.Count
                'davon aktiv'      = @($grp | Where-Object { $_.Status -eq 'aktiv' }).Count
            }
        })
        Add-Md (ConvertTo-MdTable $deptRows)
        Add-Md 'Standorte der Nutzer (aus dem Feld „Standort“; Zuordnung zu Abschnitt 2 bitte prüfen):'
        Add-Md ''
        $locRows = @($u | Group-Object -Property Standort | Sort-Object Name | ForEach-Object {
            $grp = @($_.Group)
            [pscustomobject][ordered]@{
                'Standort (AD-Feld)' = $(if ($_.Name) { $_.Name } else { '(nicht gepflegt)' })
                'Nutzer gesamt'      = $grp.Count
                'davon aktiv'        = @($grp | Where-Object { $_.Status -eq 'aktiv' }).Count
            }
        })
        Add-Md (ConvertTo-MdTable $locRows)
        Add-Md 'Vollständige Nutzerliste mit Name, Anmeldename, Abteilung, Standort und Status: nutzer.csv'
    }
    Add-Md ''

    Add-Md '## Organisationseinheiten'
    Add-Md ''
    Add-Block 'Organisationseinheiten' { Add-Md (ConvertTo-MdTable $script:Data.OUs) }

    Add-Md '## Administrative Gruppen (Fragebogen 4.2)'
    Add-Md ''
    Add-Block 'Privilegierte Gruppen' { Add-Md (ConvertTo-MdTable $script:Data.PrivSummary) }
    Add-Block 'Gruppen' {
        Add-Md "Gruppen gesamt: $(@($script:Data.Groups).Count) (Details: gruppen.csv)"
    }
    Add-Md ''

    Add-Md '## Server (Fragebogen 5.2)'
    Add-Md ''
    Add-Block 'Computerkonten' {
        Add-Md (ConvertTo-MdTable ($servers | Select-Object Name, Betriebssystem, 'Letzte Anmeldung', Status, IPv4))
    }
    Add-Md ''

    Add-Md '## Endgeräte (Fragebogen 8)'
    Add-Md ''
    Add-Block 'Computerkonten' {
        Add-Md "- Computerkonten gesamt: $($c.Count) (Server: $($servers.Count), Clients: $($clients.Count))"
        Add-Md "- Aktiv: $($cActive.Count), inaktiv seit mehr als $InactiveDays Tagen oder nie angemeldet: $($cInactive.Count)"
        Add-Md ''
        Add-Md 'Clients nach Betriebssystem:'
        Add-Md ''
        $osRows = @($clients | Group-Object -Property Betriebssystem | Sort-Object Count -Descending | ForEach-Object {
            $grp = @($_.Group)
            [pscustomobject][ordered]@{
                'Betriebssystem'     = $(if ($_.Name) { $_.Name } else { '(nicht gepflegt)' })
                'Anzahl'             = $grp.Count
                'davon inaktiv'      = @($grp | Where-Object { $_.Inaktiv -eq 'ja' }).Count
            }
        })
        Add-Md (ConvertTo-MdTable $osRows)
        Add-Md 'Vollständige Liste mit Zuordnung zu Standort (OU) und letzter Anmeldung: computerkonten.csv'
    }
    Add-Md ''

    Add-Md '## Erhebungsfehler und übersprungene Abschnitte'
    Add-Md ''
    if ($script:Problems.Count -eq 0) {
        Add-Md 'Keine.'
    }
    else {
        foreach ($p in $script:Problems) { Add-Md "- $p" }
    }
    Add-Md ''
}

Build-Markdown

$mdPath = Join-Path $OutputDir 'AD-Uebersicht.md'
[System.IO.File]::WriteAllText($mdPath, $script:Md.ToString(), (New-Object System.Text.UTF8Encoding $false))

# ---------------------------------------------------------------- HTML-Übersicht
# Reine Darstellung der bereits gelesenen Daten. Kein Netzwerkzugriff, keine externen
# Stylesheets oder Skripte; die Suche ist ein kleines eingebettetes Skript.

$HtmlCss = @'
:root{--fg:#1d2433;--muted:#6b7280;--line:#e3e7ee;--accent:#1f4e79;--bg:#f6f7f9}
*{box-sizing:border-box}
body{margin:0;font:14px/1.5 -apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;color:var(--fg);background:var(--bg)}
.wrap{max-width:1200px;margin:0 auto;padding:0 32px}
header.top{background:#fff;border-bottom:1px solid var(--line)}
header.top .wrap{padding-top:22px;padding-bottom:22px}
.eyebrow{margin:0;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;font-size:11px}
h1{margin:4px 0 4px;font-size:24px;font-weight:600}
.meta{margin:0;color:var(--muted)}
main.wrap{padding-top:16px;padding-bottom:40px}
.toolbar{position:sticky;top:0;z-index:2;background:var(--bg);padding:12px 0;display:flex;gap:12px;align-items:center}
#q{flex:1;max-width:420px;padding:8px 12px;border:1px solid #c5ccd6;border-radius:6px;font:inherit;background:#fff}
nav{display:flex;flex-wrap:wrap;gap:6px 18px;margin:4px 0 16px}
nav a{color:var(--accent);text-decoration:none}
nav a:hover{text-decoration:underline}
section{background:#fff;border:1px solid var(--line);border-radius:8px;padding:20px 24px;margin:0 0 18px}
h2{margin:0 0 6px;font-size:18px;font-weight:600;color:var(--accent)}
h3{margin:20px 0 8px;font-size:14px;font-weight:600}
.intro{color:var(--muted);margin:0 0 8px}
.cards{display:flex;flex-wrap:wrap;gap:10px;margin:14px 0 4px}
.card{border:1px solid var(--line);border-radius:6px;padding:10px 14px;min-width:150px;background:#fafbfc}
.card .n{display:block;font-size:20px;font-weight:600}
.card .l{display:block;color:var(--muted);font-size:12px}
.tablewrap{overflow-x:auto;border:1px solid var(--line);border-radius:6px}
table.t{border-collapse:collapse;width:100%;font-size:13px}
th,td{text-align:left;padding:6px 10px;border-bottom:1px solid var(--line);vertical-align:top;white-space:nowrap}
th{background:#f2f4f7;font-weight:600}
tr:last-child td{border-bottom:0}
tr.inaktiv td{background:#fdf7e6}
.muted{color:var(--muted)}
footer{color:var(--muted);font-size:12px;margin-top:24px}
@media print{.toolbar,nav{display:none}section{break-inside:avoid}body{background:#fff}}
'@

$HtmlJs = @'
(function(){
  var q=document.getElementById('q'),qc=document.getElementById('qc');
  q.addEventListener('input',function(){
    var t=q.value.trim().toLowerCase(),hits=0;
    document.querySelectorAll('table.t tbody tr').forEach(function(r){
      var show=!t||r.textContent.toLowerCase().indexOf(t)>=0;
      r.style.display=show?'':'none';
      if(show&&t){hits++;}
    });
    qc.textContent=t?hits+' Treffer in Tabellenzeilen':'';
  });
})();
'@

function HtmlEnc {
    param($Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Fmt-Cell {
    param($Value)
    if ($null -eq $Value -or ($Value -is [string] -and $Value.Trim() -eq '')) {
        return '<span class="muted">unbekannt</span>'
    }
    if ($Value -is [datetime]) { return (HtmlEnc $Value.ToString('yyyy-MM-dd')) }
    return (HtmlEnc $Value)
}

function New-HtmlTable {
    param([object[]]$Rows)
    $rows = @($Rows | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { return '<p class="muted">Keine Einträge.</p>' }
    $columns = @($rows[0].PSObject.Properties.Name)
    $hasInactive = $columns -contains 'Inaktiv'
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="tablewrap"><table class="t"><thead><tr>')
    foreach ($col in $columns) { [void]$sb.Append('<th>' + (HtmlEnc $col) + '</th>') }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($r in $rows) {
        $cls = if ($hasInactive -and $r.Inaktiv -eq 'ja') { ' class="inaktiv"' } else { '' }
        [void]$sb.Append('<tr' + $cls + '>')
        foreach ($col in $columns) { [void]$sb.Append('<td>' + (Fmt-Cell $r.$col) + '</td>') }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table></div>')
    return $sb.ToString()
}

function Get-CardsHtml {
    param([object[]]$Items)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="cards">')
    for ($i = 0; $i -lt $Items.Count - 1; $i += 2) {
        [void]$sb.Append('<div class="card"><span class="n">' + (HtmlEnc $Items[$i + 1]) + '</span><span class="l">' + (HtmlEnc $Items[$i]) + '</span></div>')
    }
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

function Cnt {
    # Zählwert nur anzeigen, wenn der zugehörige Abschnitt erfolgreich gelesen wurde.
    param($Good, $N)
    if ([bool]$Good) { return [int]$N }
    return 'nicht erhoben'
}

function Add-HtmlSection {
    param([string]$Id, [string]$Title, [string]$Intro = '', [object[]]$Cards = @(), [string]$Body = '')
    $h = $script:Html
    [void]$h.AppendLine('<section id="' + $Id + '">')
    [void]$h.AppendLine('<h2>' + (HtmlEnc $Title) + '</h2>')
    if ($Intro) { [void]$h.AppendLine('<p class="intro">' + (HtmlEnc $Intro) + '</p>') }
    if ($Cards.Count -gt 0) { [void]$h.AppendLine((Get-CardsHtml $Cards)) }
    [void]$h.AppendLine($Body)
    [void]$h.AppendLine('</section>')
}

function Get-GroupTable {
    param([object[]]$Rows, [string]$Field, [string]$Label)
    $out = @()
    foreach ($g in @($Rows | Group-Object -Property $Field | Sort-Object Name)) {
        $grp = @($g.Group)
        $row = [ordered]@{}
        $row[$Label] = $(if ($g.Name) { $g.Name } else { '(nicht gepflegt)' })
        $row['Nutzer gesamt'] = $grp.Count
        $row['davon aktiv'] = @($grp | Where-Object { $_.Status -eq 'aktiv' }).Count
        $out += [pscustomobject]$row
    }
    return $out
}

function Build-Html {
    $script:Html = New-Object System.Text.StringBuilder
    $h = $script:Html
    $ok = $script:Ok
    $NotCollected = '<p class="muted">Nicht erhoben, Details im Fehlerprotokoll.</p>'

    $u        = @($script:Data.Users)
    $c        = @($script:Data.Computers)
    $active   = @($u | Where-Object { $_.Status -eq 'aktiv' })
    $disabled = @($u | Where-Object { $_.Status -eq 'deaktiviert' })
    $inactiveUsers = @($active | Where-Object { $_.Inaktiv -eq 'ja' })
    $cActive   = @($c | Where-Object { $_.Status -eq 'aktiv' })
    $cInactive = @($c | Where-Object { $_.Inaktiv -eq 'ja' })
    $servers  = @($c | Where-Object { $_.Typ -eq 'Server' })
    $clients  = @($c | Where-Object { $_.Typ -eq 'Client' })
    $dcs      = @($script:Data.DCs)
    $dcGC     = @($dcs | Where-Object { $_.'Globaler Katalog' -eq 'ja' })
    $dcRO     = @($dcs | Where-Object { $_.'Schreibgeschützt (RODC)' -eq 'ja' })
    $sites    = @($script:Data.Sites)
    $subnets  = @($script:Data.Subnets)
    $groups   = @($script:Data.Groups)
    $secGroups = @($groups | Where-Object { $_.Typ -eq 'Security' })
    $distGroups = @($groups | Where-Object { $_.Typ -eq 'Distribution' })
    $priv     = @($script:Data.PrivSummary)
    $privMembers = 0
    foreach ($p in $priv) { $privMembers += [int]$p.'Anzahl Mitglieder (direkt)' }
    $ous      = @($script:Data.OUs)
    $problems = @($script:Problems)

    $domRows = @()
    if ($script:Data.Domain) {
        $domRows = @($script:Data.Domain.PSObject.Properties | ForEach-Object {
            [pscustomobject][ordered]@{ 'Eigenschaft' = $_.Name; 'Wert' = $_.Value }
        })
    }
    $osRows = @($clients | Group-Object -Property Betriebssystem | Sort-Object Count -Descending | ForEach-Object {
        $grp = @($_.Group)
        [pscustomobject][ordered]@{
            'Betriebssystem' = $(if ($_.Name) { $_.Name } else { '(nicht gepflegt)' })
            'Anzahl'         = $grp.Count
            'davon inaktiv'  = @($grp | Where-Object { $_.Inaktiv -eq 'ja' }).Count
        }
    })

    # Kopf und Navigation
    [void]$h.AppendLine('<!DOCTYPE html>')
    [void]$h.AppendLine('<html lang="de"><head><meta charset="utf-8">')
    [void]$h.AppendLine('<meta name="viewport" content="width=device-width,initial-scale=1">')
    [void]$h.AppendLine('<title>AD-Erhebung ' + (HtmlEnc $Domain) + '</title>')
    [void]$h.AppendLine('<style>')
    [void]$h.AppendLine($HtmlCss)
    [void]$h.AppendLine('</style></head><body>')
    [void]$h.AppendLine('<header class="top"><div class="wrap">')
    [void]$h.AppendLine('<p class="eyebrow">Lesende AD-Erhebung · Kunden-Onboarding</p>')
    [void]$h.AppendLine('<h1>' + (HtmlEnc $Domain) + '</h1>')
    [void]$h.AppendLine('<p class="meta">Erhebungszeitpunkt: ' + (HtmlEnc ($RunStart.ToString('yyyy-MM-dd HH:mm zzz'))) + ' · Abgefragt über: ' + (HtmlEnc $Target) + '</p>')
    [void]$h.AppendLine('</div></header>')
    [void]$h.AppendLine('<main class="wrap">')
    [void]$h.AppendLine('<div class="toolbar"><input id="q" type="search" placeholder="In allen Tabellen suchen …" autocomplete="off"><span id="qc" class="muted"></span></div>')
    [void]$h.AppendLine('<nav><a href="#uebersicht">Zusammenfassung</a><a href="#domain">Domäne und DCs</a><a href="#standorte">Standorte</a><a href="#nutzer">Nutzer</a><a href="#computer">Computer</a><a href="#gruppen">Gruppen</a><a href="#privilegiert">Privilegierte Gruppen</a><a href="#ous">Organisationseinheiten</a><a href="#fehler">Fehler</a></nav>')

    # Zusammenfassung
    Add-HtmlSection -Id 'uebersicht' -Title 'Zusammenfassung' `
        -Intro 'Nicht gepflegte Felder stehen als „unbekannt“. Die vollständigen Daten liegen zusätzlich als CSV-Dateien im selben Ordner.' `
        -Cards @(
            'Nutzerkonten',        (Cnt $ok['Nutzerkonten'] $u.Count),
            'Computerkonten',      (Cnt $ok['Computerkonten'] $c.Count),
            'Gruppen',             (Cnt $ok['Gruppen'] $groups.Count),
            'Domain Controller',   (Cnt $ok['Domäne und Domain Controller'] $dcs.Count),
            'AD-Standorte',        (Cnt $ok['AD-Standorte und Subnetze'] $sites.Count),
            'Erhebungsfehler',     $problems.Count
        )

    # Domäne und DCs
    $body = if ($ok['Domäne und Domain Controller']) {
        '<h3>Domäne</h3>' + (New-HtmlTable $domRows) + '<h3>Domain Controller</h3>' + (New-HtmlTable $dcs)
    } else { $NotCollected }
    Add-HtmlSection -Id 'domain' -Title 'Domäne und Domain Controller' -Intro 'Fragebogen Abschnitt 4.2, 5 und 6.1.' `
        -Cards @('Domain Controller', (Cnt $ok['Domäne und Domain Controller'] $dcs.Count),
                 'Globaler Katalog', (Cnt $ok['Domäne und Domain Controller'] $dcGC.Count),
                 'Schreibgeschützt (RODC)', (Cnt $ok['Domäne und Domain Controller'] $dcRO.Count)) `
        -Body $body

    # Standorte
    $body = if ($ok['AD-Standorte und Subnetze']) {
        '<h3>AD-Standorte (Sites)</h3>' + (New-HtmlTable $sites) + '<h3>Subnetze</h3>' + (New-HtmlTable $subnets)
    } else { $NotCollected }
    Add-HtmlSection -Id 'standorte' -Title 'Standorte' `
        -Intro 'Fragebogen Abschnitt 2 und 6.2/6.3. Die Zuordnung der AD-Standorte zu den Fragebogen-Standorten ist manuell zu prüfen. Adressen und Ansprechpartner sind im AD nicht hinterlegt.' `
        -Cards @('AD-Standorte', (Cnt $ok['AD-Standorte und Subnetze'] $sites.Count),
                 'Subnetze', (Cnt $ok['AD-Standorte und Subnetze'] $subnets.Count)) `
        -Body $body

    # Nutzer
    $body = if ($ok['Nutzerkonten']) {
        '<h3>Nach Abteilung</h3>' + (New-HtmlTable (Get-GroupTable $u 'Abteilung' 'Abteilung')) +
        '<h3>Nach Standort (AD-Feld „Standort“)</h3>' + (New-HtmlTable (Get-GroupTable $u 'Standort' 'Standort')) +
        '<h3>Alle Nutzerkonten</h3>' + (New-HtmlTable $u)
    } else { $NotCollected }
    Add-HtmlSection -Id 'nutzer' -Title 'Nutzer' `
        -Intro ('Fragebogen Abschnitt 3. „Letzte Anmeldung“ ist ein replizierter Näherungswert (Abweichung bis etwa 14 Tage). Markierte Zeilen: aktiv, aber seit mehr als {0} Tagen nicht angemeldet oder nie angemeldet.' -f $InactiveDays) `
        -Cards @('Nutzerkonten', (Cnt $ok['Nutzerkonten'] $u.Count),
                 'Aktiv', (Cnt $ok['Nutzerkonten'] $active.Count),
                 'Deaktiviert', (Cnt $ok['Nutzerkonten'] $disabled.Count),
                 ('Aktiv, aber inaktiv (über {0} Tage)' -f $InactiveDays), (Cnt $ok['Nutzerkonten'] $inactiveUsers.Count)) `
        -Body $body

    # Computer
    $body = if ($ok['Computerkonten']) {
        '<h3>Server</h3>' + (New-HtmlTable ($servers | Select-Object 'Name', 'Betriebssystem', 'Letzte Anmeldung', 'Status', 'IPv4', 'Inaktiv')) +
        '<h3>Clients nach Betriebssystem</h3>' + (New-HtmlTable $osRows) +
        '<h3>Alle Computerkonten</h3>' + (New-HtmlTable $c)
    } else { $NotCollected }
    Add-HtmlSection -Id 'computer' -Title 'Computer' `
        -Intro 'Fragebogen Abschnitt 5.2 (Server) und 8 (Endgeräte). Der Typ „Server“ wird aus dem Betriebssystem abgeleitet.' `
        -Cards @('Computerkonten', (Cnt $ok['Computerkonten'] $c.Count),
                 'Server', (Cnt $ok['Computerkonten'] $servers.Count),
                 'Clients', (Cnt $ok['Computerkonten'] $clients.Count),
                 'Aktiv', (Cnt $ok['Computerkonten'] $cActive.Count),
                 ('Inaktiv (über {0} Tage)' -f $InactiveDays), (Cnt $ok['Computerkonten'] $cInactive.Count)) `
        -Body $body

    # Gruppen
    $body = if ($ok['Gruppen']) { New-HtmlTable $groups } else { $NotCollected }
    Add-HtmlSection -Id 'gruppen' -Title 'Gruppen' `
        -Intro 'Fragebogen Abschnitt 4.2. „Anzahl Mitglieder (direkt)“ zählt nur unmittelbare Mitglieder.' `
        -Cards @('Gruppen gesamt', (Cnt $ok['Gruppen'] $groups.Count),
                 'Sicherheitsgruppen', (Cnt $ok['Gruppen'] $secGroups.Count),
                 'Verteilergruppen', (Cnt $ok['Gruppen'] $distGroups.Count)) `
        -Body $body

    # Privilegierte Gruppen
    $body = if ($ok['Privilegierte Gruppen']) { New-HtmlTable $priv } else { $NotCollected }
    Add-HtmlSection -Id 'privilegiert' -Title 'Privilegierte Gruppen' `
        -Intro 'Fragebogen Abschnitt 4.2. Nur direkte Mitglieder; verschachtelte Gruppen sind nicht aufgelöst. Es werden nur Kontonamen angezeigt.' `
        -Cards @('Geprüfte Gruppen', (Cnt $ok['Privilegierte Gruppen'] $priv.Count),
                 'Direkte Mitglieder gesamt', (Cnt $ok['Privilegierte Gruppen'] $privMembers)) `
        -Body $body

    # Organisationseinheiten
    $body = if ($ok['Organisationseinheiten']) { New-HtmlTable $ous } else { $NotCollected }
    Add-HtmlSection -Id 'ous' -Title 'Organisationseinheiten' `
        -Cards @('OUs gesamt', (Cnt $ok['Organisationseinheiten'] $ous.Count)) `
        -Body $body

    # Fehler
    if ($problems.Count -eq 0) {
        $body = '<p class="muted">Keine.</p>'
    }
    else {
        $items = foreach ($p in $problems) { '<li>' + (HtmlEnc $p) + '</li>' }
        $body = '<ul>' + ($items -join '') + '</ul>'
    }
    Add-HtmlSection -Id 'fehler' -Title 'Erhebungsfehler und übersprungene Abschnitte' `
        -Cards @('Einträge', $problems.Count) -Body $body

    [void]$h.AppendLine('<footer>Es wurden ausschließlich lesende Abfragen ausgeführt. Die Übersicht enthält keine Passwörter, Hashes oder Schlüssel.</footer>')
    [void]$h.AppendLine('</main>')
    [void]$h.AppendLine('<script>')
    [void]$h.AppendLine($HtmlJs)
    [void]$h.AppendLine('</script></body></html>')
}

Build-Html
$htmlPath = Join-Path $OutputDir 'AD-Uebersicht.html'
[System.IO.File]::WriteAllText($htmlPath, $script:Html.ToString(), (New-Object System.Text.UTF8Encoding $false))

# ---------------------------------------------------------------- Excel- und Word-Datei
# Ohne Office und ohne Module: Die Dateien werden als OOXML-Pakete (ZIP mit XML) direkt geschrieben.
# Die Excel-Vorlage ist im Skript eingebettet; -TemplatePath nutzt stattdessen eine eigene Datei.
# Vorlage und Originale werden nie verändert: Es wird immer eine neue Datei im Ausgabeordner angelegt.

$TemplateBase64 = @'
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
R/AcZTACFAGGAEdAICAjIPxVzSdxgpIYAYoAQ4AjIBCQERAa8EmcoCRGgCLAEOAICARkBIQGnvy2jJIYAYoAQ4AjIBCQERAa8Emc
oCRGgCLAEOAICARkBIQGfBInKIkRoAgwBDgCAgEZAaEBn8QJSmIEKAIMAY6AQEBGQGjAJ3GCkhgBigBDgCMgEJAREBrwSZygJEaA
IsAQ4AgIBGQEhAZ8EicoiRGgCDAEOAICARkBoQGfxAlKYgQoAgwBjoBAQEZAYCD1SZyiJEaAIsAQ4AgIBGQEhAZ8EqcoiRGgCDAE
OAICARkBoQGfxClKYgQoAgwBjoBAQEZAaMAncYqSGAGKAEOAIyAQkBEQGvBJnKIkRoAiwBDgCAgEZASEBnyQpfvIOOlOPfw0W1Qq
04YoQ1bu5X1j3Etsb6Bv034gFe1zbeyurOwzn4SnPpvS0+ebiXlsr+OxCm0Jfck1+ZbrQvc7QyTLby3RNclNnelabezbqLORSEp9
JKVnI3Mn0bnPRubWxXKt/QaTW4d+8lfdztPOWSl0brUlda7jHkZSKfWplJ6PeEijHs6f9yD7Wy2//11b/YrcVOre5vfKWOemXLXb
YQ9WV0YTlmtTW+entv0229cmUyZ7S9r1K9yRJo65x7FypXtGsu//FjbP4t5HAm3hA20xG/G+iO6xzcafO7dG7feP0c3w9LVOCp31
u4dNbavcWk0y3e3BWee39a+Wa+dV1VFrCxR10ycbp+22/AdVZe4pd4/JnTtzdnTqftCt9pva+ze23HZ7qbelteWmO1xr5e5xe4Lj
d2VphzftZYa/OLj6D1BLAwQUAAAACACoqmxdsRw4C/EGAADhHAAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQ1LnhtbK1ZzW7bOBB+
FcKnBvCvJDtpkWSRSE5aJGmDGm2A3mR7bBOWSS9Fx0lOfZBie+kz9NSb36RPskPJNq2aFNvsHtrI/DTDb0bkN0PpeMnFNJ0ASPIw
S1h6UplIOX/VaKSDCczitM7nwBAZcTGLJf4U40Y6FxAPM6NZ0vCazU5jFlNWOT3Oxm7F6TFfyIQyuBUkXcxmsXg8h4QvTyqtymbg
PR1PZDbQOD2ex2PogfwwRwP82dj6GdIZsJRyRgSMTipnrVfdoJ2ZZLd8pLBMd65JOuHLS0GH1zh5mnlX8fU5nyr4zfCk0lQ0IYGB
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
'@

function XmlText {
    param($Value)
    if ($null -eq $Value) { return '' }
    $t = [string]$Value
    $t = $t -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    return $t.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

function CellText {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm') }
    if ($Value -is [bool]) { if ($Value) { return 'ja' } else { return 'nein' } }
    return [string]$Value
}

function ColName {
    param([int]$Index)
    $name = ''
    $n = $Index
    while ($n -gt 0) {
        $m = ($n - 1) % 26
        $name = ([string][char](65 + $m)) + $name
        $n = [int][math]::Floor(($n - 1) / 26)
    }
    return $name
}

function Get-FreePath {
    # Nie überschreiben: existiert der Name schon, bekommt er einen Zeitstempel-Zusatz.
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $Path }
    $dir  = Split-Path -Parent $Path
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $ext  = [System.IO.Path]::GetExtension($Path)
    return (Join-Path $dir ('{0}-{1}{2}' -f $base, (Get-Date -Format 'HHmmss'), $ext))
}

function Get-TemplateBytes {
    if ($TemplatePath) {
        return [System.IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $TemplatePath).ProviderPath)
    }
    return [System.Convert]::FromBase64String(($TemplateBase64 -replace '\s', ''))
}

function Read-ZipTexts {
    param([byte[]]$Bytes)
    $ms = [System.IO.MemoryStream]::new($Bytes)
    $result = [ordered]@{}
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($ms, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($e in $zip.Entries) {
                if ($e.FullName.EndsWith('/')) { continue }
                $sr = [System.IO.StreamReader]::new($e.Open(), [System.Text.UTF8Encoding]::new($false))
                try { $result[$e.FullName] = $sr.ReadToEnd() } finally { $sr.Dispose() }
            }
        }
        finally { $zip.Dispose() }
    }
    finally { $ms.Dispose() }
    return $result
}

function Add-ZipText {
    param($Zip, [string]$Name, [string]$Text)
    $entry = $Zip.CreateEntry($Name)
    $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes($Text)
    $s = $entry.Open()
    try { $s.Write($bytes, 0, $bytes.Length) } finally { $s.Dispose() }
}

function New-KennzahlRow {
    param($K, $V)
    return [pscustomobject][ordered]@{ 'Kennzahl' = $K; 'Wert' = $V }
}

function Get-Kennzahlen {
    # Eine Quelle für die Kennzahlen in Excel und Word.
    $ok  = $script:Ok
    $u   = @($script:Data.Users)
    $c   = @($script:Data.Computers)
    $dcs = @($script:Data.DCs)
    $okD = [bool]$ok['Domäne und Domain Controller']
    $okS = [bool]$ok['AD-Standorte und Subnetze']
    $okO = [bool]$ok['Organisationseinheiten']
    $okU = [bool]$ok['Nutzerkonten']
    $okC = [bool]$ok['Computerkonten']
    $okG = [bool]$ok['Gruppen']
    $okP = [bool]$ok['Privilegierte Gruppen']

    $active    = @($u | Where-Object { $_.Status -eq 'aktiv' })
    $disabled  = @($u | Where-Object { $_.Status -eq 'deaktiviert' })
    $inactiveU = @($active | Where-Object { $_.Inaktiv -eq 'ja' })
    $noDept    = @($u | Where-Object { -not $_.Abteilung })
    $noOffice  = @($u | Where-Object { -not $_.Standort })
    $servers   = @($c | Where-Object { $_.Typ -eq 'Server' })
    $clients   = @($c | Where-Object { $_.Typ -eq 'Client' })
    $cActive   = @($c | Where-Object { $_.Status -eq 'aktiv' })
    $cInactive = @($c | Where-Object { $_.Inaktiv -eq 'ja' })
    $groups    = @($script:Data.Groups)
    $secG      = @($groups | Where-Object { $_.Typ -eq 'Security' })
    $distG     = @($groups | Where-Object { $_.Typ -eq 'Distribution' })
    $dcGC      = @($dcs | Where-Object { $_.'Globaler Katalog' -eq 'ja' })
    $dcRO      = @($dcs | Where-Object { $_.'Schreibgeschützt (RODC)' -eq 'ja' })
    $privCount = [int](@($script:Data.PrivDetail).Count)

    $rows = New-Object System.Collections.ArrayList

    [void]$rows.Add((New-KennzahlRow 'Domäne (DNS-Name)' $Domain))
    [void]$rows.Add((New-KennzahlRow 'Erhebungszeitpunkt (UTC)' ($RunStart.ToUniversalTime().ToString('yyyy-MM-dd HH:mm'))))
    [void]$rows.Add((New-KennzahlRow 'Domain Controller' (Cnt $okD $dcs.Count)))
    [void]$rows.Add((New-KennzahlRow '  davon Globaler Katalog' (Cnt $okD $dcGC.Count)))
    [void]$rows.Add((New-KennzahlRow '  davon schreibgeschützt (RODC)' (Cnt $okD $dcRO.Count)))
    [void]$rows.Add((New-KennzahlRow 'AD-Standorte (Sites)' (Cnt $okS @($script:Data.Sites).Count)))
    [void]$rows.Add((New-KennzahlRow 'Subnetze' (Cnt $okS @($script:Data.Subnets).Count)))
    [void]$rows.Add((New-KennzahlRow 'Organisationseinheiten' (Cnt $okO @($script:Data.OUs).Count)))
    [void]$rows.Add((New-KennzahlRow 'Nutzerkonten gesamt' (Cnt $okU $u.Count)))
    [void]$rows.Add((New-KennzahlRow '  aktiv' (Cnt $okU $active.Count)))
    [void]$rows.Add((New-KennzahlRow '  deaktiviert' (Cnt $okU $disabled.Count)))
    [void]$rows.Add((New-KennzahlRow ('  aktiv, aber inaktiv (über {0} Tage oder nie angemeldet)' -f $InactiveDays) (Cnt $okU $inactiveU.Count)))
    [void]$rows.Add((New-KennzahlRow '  ohne Abteilungsangabe im AD' (Cnt $okU $noDept.Count)))
    [void]$rows.Add((New-KennzahlRow '  ohne Standortangabe im AD' (Cnt $okU $noOffice.Count)))
    [void]$rows.Add((New-KennzahlRow 'Computerkonten gesamt' (Cnt $okC $c.Count)))
    [void]$rows.Add((New-KennzahlRow '  davon Server' (Cnt $okC $servers.Count)))
    [void]$rows.Add((New-KennzahlRow '  davon Clients' (Cnt $okC $clients.Count)))
    [void]$rows.Add((New-KennzahlRow '  aktiv' (Cnt $okC $cActive.Count)))
    [void]$rows.Add((New-KennzahlRow ('  inaktiv (über {0} Tage oder nie angemeldet)' -f $InactiveDays) (Cnt $okC $cInactive.Count)))
    [void]$rows.Add((New-KennzahlRow 'Gruppen gesamt' (Cnt $okG $groups.Count)))
    [void]$rows.Add((New-KennzahlRow '  Sicherheitsgruppen' (Cnt $okG $secG.Count)))
    [void]$rows.Add((New-KennzahlRow '  Verteilergruppen' (Cnt $okG $distG.Count)))
    [void]$rows.Add((New-KennzahlRow 'Direkte Mitglieder in privilegierten Gruppen' (Cnt $okP $privCount)))
    [void]$rows.Add((New-KennzahlRow 'Erhebungsfehler / übersprungene Abschnitte' @($script:Problems).Count))

    return @($rows.ToArray())
}

function New-SheetXml {
    param([string]$Title, [string]$Note, [string[]]$Columns, [object[]]$Rows, [string]$TabColor)
    $Rows  = @($Rows | Where-Object { $null -ne $_ })
    $ncol  = $Columns.Count
    $lastCol = ColName $ncol
    $widths = @()
    for ($c = 0; $c -lt $ncol; $c++) {
        $w = $Columns[$c].Length + 2
        foreach ($r in $Rows) {
            $len = (CellText ($r.($Columns[$c]))).Length + 2
            if ($len -gt $w) { $w = $len }
        }
        if ($w -gt 60) { $w = 60 }
        if ($w -lt 10) { $w = 10 }
        $widths += $w
    }
    $lastRow = 4 + [math]::Max(1, $Rows.Count)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    [void]$sb.Append('<sheetPr><tabColor rgb="' + $TabColor + '"/></sheetPr>')
    [void]$sb.Append('<dimension ref="A1:' + $lastCol + [string]$lastRow + '"/>')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="15"/><cols>')
    for ($c = 0; $c -lt $ncol; $c++) {
        $i = $c + 1
        [void]$sb.Append('<col min="' + $i + '" max="' + $i + '" width="' + $widths[$c] + '" customWidth="1"/>')
    }
    [void]$sb.Append('</cols><sheetData>')
    [void]$sb.Append('<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>' + (XmlText $Title) + '</t></is></c></row>')
    [void]$sb.Append('<row r="2"><c r="A2" s="2" t="inlineStr"><is><t>' + (XmlText $Note) + '</t></is></c></row>')
    [void]$sb.Append('<row r="4">')
    for ($c = 0; $c -lt $ncol; $c++) {
        $ref = (ColName ($c + 1)) + '4'
        [void]$sb.Append('<c r="' + $ref + '" s="7" t="inlineStr"><is><t>' + (XmlText $Columns[$c]) + '</t></is></c>')
    }
    [void]$sb.Append('</row>')
    if ($Rows.Count -eq 0) {
        [void]$sb.Append('<row r="5"><c r="A5" s="8" t="inlineStr"><is><t>Keine Einträge.</t></is></c></row>')
    }
    else {
        $rn = 4
        foreach ($r in $Rows) {
            $rn++
            [void]$sb.Append('<row r="' + $rn + '">')
            for ($c = 0; $c -lt $ncol; $c++) {
                $ref  = (ColName ($c + 1)) + [string]$rn
                $text = CellText ($r.($Columns[$c]))
                if ($text -eq '') {
                    [void]$sb.Append('<c r="' + $ref + '" s="8"/>')
                }
                else {
                    [void]$sb.Append('<c r="' + $ref + '" s="8" t="inlineStr"><is><t xml:space="preserve">' + (XmlText $text) + '</t></is></c>')
                }
            }
            [void]$sb.Append('</row>')
        }
    }
    [void]$sb.Append('</sheetData><pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')
    return $sb.ToString()
}

function Get-XlsxSheets {
    param([object[]]$Kennzahlen)
    $ok = $script:Ok
    $stand = 'Stand: {0} (UTC). Leere Felder = im AD nicht gepflegt. Keine Passwörter, Hashes oder Schlüssel.' -f $RunStart.ToUniversalTime().ToString('dd.MM.yyyy HH:mm')
    $notErh = 'Nicht erhoben, Details im Fehlerprotokoll.'

    $domRows = @()
    if ($script:Data.Domain) {
        $domRows = @($script:Data.Domain.PSObject.Properties | ForEach-Object {
            [pscustomobject][ordered]@{ 'Eigenschaft' = $_.Name; 'Wert' = $_.Value }
        })
    }

    $sheets = @(
        @{ Name = 'AD Kennzahlen';         Title = 'AD-Kennzahlen';              Note = $stand;   Columns = @('Kennzahl', 'Wert');                                                  Rows = $Kennzahlen;                          Tab = 'FF1F4E79' },
        @{ Name = 'AD Domäne';             Title = 'AD Domäne';                  Note = $(if ($ok['Domäne und Domain Controller']) { $stand } else { $notErh }); Columns = @('Eigenschaft', 'Wert');                                 Rows = $domRows;                             Tab = 'FF2E75B6' },
        @{ Name = 'AD Domain Controller';  Title = 'AD Domain Controller';       Note = $(if ($ok['Domäne und Domain Controller']) { $stand } else { $notErh }); Columns = @('Name', 'Hostname', 'IPv4', 'AD-Standort', 'Betriebssystem', 'Globaler Katalog', 'Schreibgeschützt (RODC)'); Rows = @($script:Data.DCs);       Tab = 'FF2E75B6' },
        @{ Name = 'AD Standorte';          Title = 'AD Standorte (Sites)';       Note = $(if ($ok['AD-Standorte und Subnetze']) { $stand } else { $notErh }); Columns = @('Standort (AD-Site)', 'Beschreibung');                                    Rows = @($script:Data.Sites);                Tab = 'FF2E75B6' },
        @{ Name = 'AD Subnetze';           Title = 'AD Subnetze';                Note = $(if ($ok['AD-Standorte und Subnetze']) { $stand } else { $notErh }); Columns = @('Subnetz', 'AD-Standort', 'Ort', 'Beschreibung');                         Rows = @($script:Data.Subnets);              Tab = 'FF2E75B6' },
        @{ Name = 'AD Nutzer';             Title = 'AD Nutzerkonten';            Note = $(if ($ok['Nutzerkonten']) { $stand } else { $notErh }); Columns = @('Name', 'Anmeldename', 'UPN', 'E-Mail', 'Abteilung', 'Funktion', 'Standort', 'Status', 'Letzte Anmeldung', 'Inaktiv', 'Erstellt', 'OU'); Rows = @($script:Data.Users);      Tab = 'FF548235' },
        @{ Name = 'AD Computer';           Title = 'AD Computerkonten';          Note = $(if ($ok['Computerkonten']) { $stand } else { $notErh }); Columns = @('Name', 'DNS-Hostname', 'Betriebssystem', 'OS-Version', 'Letzte Anmeldung', 'Status', 'Inaktiv', 'IPv4', 'Typ', 'OU', 'Beschreibung', 'Erstellt'); Rows = @($script:Data.Computers);  Tab = 'FF548235' },
        @{ Name = 'AD Gruppen';            Title = 'AD Gruppen';                 Note = $(if ($ok['Gruppen']) { $stand } else { $notErh }); Columns = @('Gruppe', 'Bereich', 'Typ', 'Anzahl Mitglieder (direkt)', 'OU', 'Beschreibung'); Rows = @($script:Data.Groups);            Tab = 'FF548235' },
        @{ Name = 'AD Privilegierte';      Title = 'AD Privilegierte Gruppen (direkte Mitglieder)'; Note = $(if ($ok['Privilegierte Gruppen']) { $stand } else { $notErh }); Columns = @('Gruppe', 'Mitglied', 'Objektart'); Rows = @($script:Data.PrivDetail);   Tab = 'FFC00000' }
    )
    return $sheets
}

function Build-Xlsx {
    param([string]$Path, [object[]]$Kennzahlen)
    $parts = Read-ZipTexts (Get-TemplateBytes)
    $wb  = [string]$parts['xl/workbook.xml']
    $rel = [string]$parts['xl/_rels/workbook.xml.rels']
    $ct  = [string]$parts['[Content_Types].xml']

    $sheetTags = ''
    $relTags   = ''
    $ctTags    = ''
    $newParts  = [ordered]@{}
    $i = 0
    foreach ($s in (Get-XlsxSheets -Kennzahlen $Kennzahlen)) {
        $i++
        $num = 12 + $i
        $sheetTags += '<sheet name="' + (XmlText $s.Name) + '" sheetId="' + [string]$num + '" state="visible" r:id="rIdAD' + [string]$i + '" />'
        $relTags   += '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet' + [string]$num + '.xml" Id="rIdAD' + [string]$i + '" />'
        $ctTags    += '<Override PartName="/xl/worksheets/sheet' + [string]$num + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />'
        $newParts['xl/worksheets/sheet' + [string]$num + '.xml'] = New-SheetXml -Title $s.Title -Note $s.Note -Columns $s.Columns -Rows $s.Rows -TabColor $s.Tab
    }

    $wb  = $wb.Replace('</sheets>', $sheetTags + '</sheets>')
    $rel = $rel.Replace('</Relationships>', $relTags + '</Relationships>')
    $ct  = $ct.Replace('</Types>', $ctTags + '</Types>')
    $ov  = [string]$parts['xl/worksheets/sheet1.xml']
    $stamp = $RunStart.ToUniversalTime().ToString('dd.MM.yyyy HH:mm')
    $ov  = $ov.Replace('Noch nicht vorhanden: AD-Erhebung.ps1 wurde noch nicht ausgeführt', (XmlText ('Eingelesen am {0} (UTC) aus AD-Erhebung, Domäne {1}; Ergebnisse in den Tabs „AD …“' -f $stamp, $Domain)))

    $parts['xl/workbook.xml']             = $wb
    $parts['xl/_rels/workbook.xml.rels']  = $rel
    $parts['[Content_Types].xml']         = $ct
    $parts['xl/worksheets/sheet1.xml']    = $ov

    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($k in $parts.Keys) { Add-ZipText $zip $k $parts[$k] }
            foreach ($k in $newParts.Keys) { Add-ZipText $zip $k $newParts[$k] }
        }
        finally { $zip.Dispose() }
    }
    finally { $fs.Dispose() }
}

# ---- Word (WordprocessingML)

function DocRun {
    param([string]$Text, [switch]$Bold, [switch]$White, [int]$Size = 0)
    $rpr = ''
    if ($Bold)  { $rpr += '<w:b/>' }
    if ($White) { $rpr += '<w:color w:val="FFFFFF"/>' }
    if ($Size -gt 0) { $rpr += '<w:sz w:val="' + [string]$Size + '"/>' }
    if ($rpr) { $rpr = '<w:rPr>' + $rpr + '</w:rPr>' }
    return '<w:r>' + $rpr + '<w:t xml:space="preserve">' + (XmlText $Text) + '</w:t></w:r>'
}

function DocPara {
    param([string]$Text, [string]$Style = '', [switch]$Bold, [switch]$Bullet)
    $ppr = ''
    if ($Style)  { $ppr += '<w:pStyle w:val="' + $Style + '"/>' }
    if ($Bullet) {
        $ppr += '<w:ind w:left="284" w:hanging="284"/>'
        $Text = [string][char]0x2022 + ' ' + $Text
    }
    if ($ppr) { $ppr = '<w:pPr>' + $ppr + '</w:pPr>' }
    return '<w:p>' + $ppr + (DocRun $Text -Bold:$Bold) + '</w:p>'
}

function DocCell {
    param([string]$Text, [int]$Width, [switch]$Header)
    $shade = ''
    if ($Header) { $shade = '<w:shd w:val="clear" w:color="auto" w:fill="1F4E79"/>' }
    $run = DocRun $Text -Bold:$Header -White:$Header -Size 18
    return '<w:tc><w:tcPr><w:tcW w:w="' + [string]$Width + '" w:type="dxa"/>' + $shade + '</w:tcPr>' +
           '<w:p><w:pPr><w:spacing w:before="20" w:after="20" w:line="240" w:lineRule="auto"/></w:pPr>' + $run + '</w:p></w:tc>'
}

function DocTable {
    param([string[]]$Columns, [object[]]$Rows)
    $Rows = @($Rows | Where-Object { $null -ne $_ })
    $n  = $Columns.Count
    $gw = [int][math]::Floor(9638 / $n)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>')
    foreach ($edge in @('top', 'left', 'bottom', 'right', 'insideH', 'insideV')) {
        [void]$sb.Append('<w:' + $edge + ' w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>')
    }
    [void]$sb.Append('</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="70" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>')
    for ($i = 0; $i -lt $n; $i++) { [void]$sb.Append('<w:gridCol w:w="' + [string]$gw + '"/>') }
    [void]$sb.Append('</w:tblGrid>')

    [void]$sb.Append('<w:tr><w:trPr><w:tblHeader/></w:trPr>')
    foreach ($col in $Columns) { [void]$sb.Append((DocCell $col $gw -Header)) }
    [void]$sb.Append('</w:tr>')

    if ($Rows.Count -eq 0) {
        [void]$sb.Append('<w:tr>' + (DocCell 'Keine Einträge.' $gw))
        for ($i = 1; $i -lt $n; $i++) { [void]$sb.Append((DocCell '' $gw)) }
        [void]$sb.Append('</w:tr>')
    }
    else {
        foreach ($r in $Rows) {
            [void]$sb.Append('<w:tr>')
            foreach ($col in $Columns) { [void]$sb.Append((DocCell (CellText ($r.$col)) $gw)) }
            [void]$sb.Append('</w:tr>')
        }
    }
    [void]$sb.Append('</w:tbl>')
    return $sb.ToString()
}

function DocAdd {
    param([string]$Xml)
    [void]$script:DocBody.Append($Xml)
}

function Build-Docx {
    param([string]$Path, [object[]]$Kennzahlen, [string]$XlsxName)
    $ok = $script:Ok
    $script:DocBody = New-Object System.Text.StringBuilder
    $u       = @($script:Data.Users)
    $c       = @($script:Data.Computers)
    $dcs     = @($script:Data.DCs)
    $sites   = @($script:Data.Sites)
    $subnets = @($script:Data.Subnets)
    $priv    = @($script:Data.PrivDetail)
    $okD = [bool]$ok['Domäne und Domain Controller']
    $okS = [bool]$ok['AD-Standorte und Subnetze']
    $okU = [bool]$ok['Nutzerkonten']
    $okC = [bool]$ok['Computerkonten']
    $okG = [bool]$ok['Gruppen']
    $okP = [bool]$ok['Privilegierte Gruppen']
    $notErh = 'Nicht erhoben, Details im Fehlerprotokoll.'
    $active    = @($u | Where-Object { $_.Status -eq 'aktiv' })
    $disabled  = @($u | Where-Object { $_.Status -eq 'deaktiviert' })
    $inactiveU = @($active | Where-Object { $_.Inaktiv -eq 'ja' })
    $noDept    = @($u | Where-Object { -not $_.Abteilung })
    $noOffice  = @($u | Where-Object { -not $_.Standort })
    $servers   = @($c | Where-Object { $_.Typ -eq 'Server' })
    $clients   = @($c | Where-Object { $_.Typ -eq 'Client' })
    $cInactive = @($c | Where-Object { $_.Inaktiv -eq 'ja' })
    $groups    = @($script:Data.Groups)
    $osRows = @($clients | Group-Object -Property Betriebssystem | Sort-Object Count -Descending | ForEach-Object {
        $grp = @($_.Group)
        [pscustomobject][ordered]@{
            'Betriebssystem' = $(if ($_.Name) { $_.Name } else { '(nicht gepflegt)' })
            'Anzahl'         = $grp.Count
            'davon inaktiv'  = @($grp | Where-Object { $_.Inaktiv -eq 'ja' }).Count
        }
    })

    # Kopf
    DocAdd (DocPara ('AD-Zusammenfassung: ' + $Domain) 'Title')
    DocAdd (DocPara ('Erhebungszeitpunkt: ' + $RunStart.ToString('yyyy-MM-dd HH:mm zzz') + ' · Abgefragt über: ' + $Target + ' · Lesende Erhebung, keine Passwörter, Hashes oder Schlüssel.') 'Meta')
    DocAdd (DocPara 'Felder, die im AD nicht gepflegt sind, stehen als „unbekannt“ bzw. bleiben leer. „Letzte Anmeldung“ ist ein replizierter Näherungswert (Abweichung bis etwa 14 Tage). Die vollständigen Listen stehen in der Excel-Datei und in den CSV-Dateien im selben Ordner.' 'Meta')

    # 1 Ansprechpartner
    DocAdd (DocPara '1. Ansprechpartner' 'Heading1')
    DocAdd (DocPara 'Die Ansprechpartner des Kunden stehen im Fragebogen (Abschnitt 1). Das AD enthält keine Kontaktdaten dieser Rollen; die Felder bleiben offen.')
    $roles = @('Hauptansprechpartner', 'Stellvertretung', 'Geschäftsführung / Entscheider', 'Ansprechpartner Notfall (außerhalb der Geschäftszeiten)', 'Externer Dienstleister (Telefonie, Internet, Drucker)')
    $roleRows = @($roles | ForEach-Object {
        [pscustomobject][ordered]@{ 'Rolle' = $_; 'Name' = 'offen'; 'E-Mail' = 'offen'; 'Telefon (Büro)' = 'offen'; 'Telefon (mobil)' = 'offen' }
    })
    DocAdd (DocTable @('Rolle', 'Name', 'E-Mail', 'Telefon (Büro)', 'Telefon (mobil)') $roleRows)
    DocAdd (DocPara 'Administrative Konten aus dem AD (direkte Mitglieder privilegierter Gruppen, für IT-Rückfragen zu klären)' 'Heading2')
    if ($okP) { DocAdd (DocTable @('Gruppe', 'Mitglied', 'Objektart') $priv) }
    else { DocAdd (DocPara $notErh) }

    # 2 Standorte
    DocAdd (DocPara '2. Standorte' 'Heading1')
    if ($okS) {
        DocAdd (DocPara ('AD-Standorte (Sites): {0}. Die Zuordnung zu den Fragebogen-Standorten (Adresse, Ansprechpartner vor Ort, Internetanschluss) ist manuell zu prüfen; Adressen sind im AD nicht hinterlegt.' -f $sites.Count))
        $siteRows = @($sites | ForEach-Object {
            $sn = $_.'Standort (AD-Site)'
            [pscustomobject][ordered]@{
                'AD-Standort'       = $sn
                'Beschreibung'      = $_.Beschreibung
                'Subnetze'          = @($subnets | Where-Object { $_.'AD-Standort' -eq $sn }).Count
                'Domain Controller' = @($dcs | Where-Object { $_.'AD-Standort' -eq $sn }).Count
            }
        })
        DocAdd (DocTable @('AD-Standort', 'Beschreibung', 'Subnetze', 'Domain Controller') $siteRows)
        DocAdd (DocPara 'Subnetze' 'Heading2')
        DocAdd (DocTable @('Subnetz', 'AD-Standort', 'Ort', 'Beschreibung') $subnets)
    }
    else { DocAdd (DocPara $notErh) }
    if ($okU) {
        DocAdd (DocPara 'Nutzer je Standort (AD-Feld „Standort“; Zuordnung zu den Fragebogen-Standorten bitte prüfen)' 'Heading2')
        DocAdd (DocTable @('Standort (AD-Feld)', 'Nutzer gesamt', 'davon aktiv') (Get-GroupTable $u 'Standort' 'Standort (AD-Feld)'))
    }

    # 3 Nutzer
    DocAdd (DocPara '3. Nutzer' 'Heading1')
    if ($okU) {
        DocAdd (DocPara ('Nutzerkonten gesamt: {0}' -f $u.Count) -Bullet)
        DocAdd (DocPara ('Aktiv: {0}, deaktiviert: {1}' -f $active.Count, $disabled.Count) -Bullet)
        DocAdd (DocPara ('Aktiv, aber seit mehr als {0} Tagen nicht angemeldet oder nie angemeldet: {1}' -f $InactiveDays, $inactiveU.Count) -Bullet)
        DocAdd (DocPara 'Abteilungen (AD-Feld „Abteilung“)' 'Heading2')
        DocAdd (DocTable @('Abteilung', 'Nutzer gesamt', 'davon aktiv') (Get-GroupTable $u 'Abteilung' 'Abteilung'))
        DocAdd (DocPara ('Ohne Abteilungsangabe im AD: {0} Nutzer. Ohne Standortangabe im AD: {1} Nutzer.' -f $noDept.Count, $noOffice.Count))
    }
    else { DocAdd (DocPara $notErh) }

    # 4 Systeme und Domäne
    DocAdd (DocPara '4. Systeme und Domäne' 'Heading1')
    DocAdd (DocPara 'Domäne' 'Heading2')
    if ($okD -and $script:Data.Domain) {
        $domRows = @($script:Data.Domain.PSObject.Properties | ForEach-Object { [pscustomobject][ordered]@{ 'Eigenschaft' = $_.Name; 'Wert' = $_.Value } })
        DocAdd (DocTable @('Eigenschaft', 'Wert') $domRows)
    }
    else { DocAdd (DocPara $notErh) }
    DocAdd (DocPara 'Domain Controller' 'Heading2')
    if ($okD) { DocAdd (DocTable @('Name', 'IPv4', 'AD-Standort', 'Betriebssystem', 'Globaler Katalog', 'Schreibgeschützt (RODC)') $dcs) }
    else { DocAdd (DocPara $notErh) }
    DocAdd (DocPara 'Server' 'Heading2')
    if ($okC) {
        DocAdd (DocPara ('Computerkonten gesamt: {0} (Server: {1}, Clients: {2}). Aktiv: {3}, inaktiv seit mehr als {4} Tagen oder nie angemeldet: {5}.' -f $c.Count, $servers.Count, $clients.Count, @($c | Where-Object { $_.Status -eq 'aktiv' }).Count, $InactiveDays, $cInactive.Count))
        DocAdd (DocTable @('Name', 'Betriebssystem', 'Letzte Anmeldung', 'Status', 'IPv4') $servers)
        DocAdd (DocPara 'Clients nach Betriebssystem' 'Heading2')
        DocAdd (DocTable @('Betriebssystem', 'Anzahl', 'davon inaktiv') $osRows)
    }
    else { DocAdd (DocPara $notErh) }
    DocAdd (DocPara 'Gruppen' 'Heading2')
    if ($okG) { DocAdd (DocPara ('Gruppen gesamt: {0} (Sicherheitsgruppen: {1}, Verteilergruppen: {2}).' -f $groups.Count, @($groups | Where-Object { $_.Typ -eq 'Security' }).Count, @($groups | Where-Object { $_.Typ -eq 'Distribution' }).Count)) }
    else { DocAdd (DocPara $notErh) }

    # 5 Offene Punkte
    DocAdd (DocPara '5. Offene Punkte' 'Heading1')
    DocAdd (DocPara 'Kundenname, Hauptansprechpartner, Stellvertretung, Geschäftsführung und Notfallkontakt mit Kontaktdaten (Fragebogen Abschnitt 1)' -Bullet)
    DocAdd (DocPara 'Adressen, Ansprechpartner vor Ort und Internetanschluss je Standort (Abschnitte 2 und 7.1)' -Bullet)
    DocAdd (DocPara 'Zuordnung der AD-Standorte und Subnetze zu den Fragebogen-Standorten (manuell)' -Bullet)
    if ($okU) {
        DocAdd (DocPara ('AD-Felder Abteilung und Standort prüfen: {0} Nutzer ohne Abteilung, {1} ohne Standort.' -f $noDept.Count, $noOffice.Count) -Bullet)
        DocAdd (DocPara ('{0} aktive Konten ohne Anmeldung seit mehr als {1} Tagen und {2} deaktivierte Konten mit dem Kunden klären.' -f $inactiveU.Count, $InactiveDays, $disabled.Count) -Bullet)
    }
    DocAdd (DocPara 'Virtualisierung, Firewall, VPN, Internetanbieter und Namenskonvention sind nicht im AD hinterlegt (Abschnitte 5, 6 und 7).' -Bullet)
    DocAdd (DocPara 'Benutzernamen und Ablageorte der Admin-Passwörter erfassen (Abschnitt 4), ohne Passwörter.' -Bullet)
    DocAdd (DocPara 'Vorhandene Dokumentation und ihr Ablageort (Abschnitt 10).' -Bullet)
    DocAdd (DocPara 'Wer Bestellungen und Störungsmeldungen freigibt (Abschnitt 11.2).' -Bullet)
    $problems = @($script:Problems)
    if ($problems.Count -gt 0) {
        DocAdd (DocPara ('Erhebungsfehler: {0} (Details in fehlerprotokoll.txt, falls vorhanden):' -f $problems.Count) -Bullet)
        foreach ($p in $problems) { DocAdd (DocPara $p -Bullet) }
    }

    # 6 Kennzahlen
    DocAdd (DocPara '6. AD-Kennzahlen' 'Heading1')
    DocAdd (DocTable @('Kennzahl', 'Wert') $Kennzahlen)
    DocAdd (DocPara ('Excel-Datei mit den AD-Tabs: {0}' -f $XlsxName) 'Meta')

    $sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>'
    $docXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>' +
        $script:DocBody.ToString() + $sect + '</w:body></w:document>'

    $contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>
'@
    $rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>
'@
    $docRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
'@
    $styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="de-DE"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="60"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="40"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="2" w:color="1F4E79"/></w:pBdr><w:spacing w:before="320" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="23"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Meta"><w:name w:val="Meta"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="80"/></w:pPr><w:rPr><w:color w:val="595959"/><w:sz w:val="18"/></w:rPr></w:style></w:styles>
'@
    $core = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        '<dc:title>' + (XmlText ('AD-Zusammenfassung ' + $Domain)) + '</dc:title><dc:creator>AD-Erhebung.ps1</dc:creator>' +
        '<dcterms:created xsi:type="dcterms:W3CDTF">' + $RunStart.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') + '</dcterms:created></cp:coreProperties>'

    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            Add-ZipText $zip '[Content_Types].xml' $contentTypes
            Add-ZipText $zip '_rels/.rels' $rootRels
            Add-ZipText $zip 'docProps/core.xml' $core
            Add-ZipText $zip 'word/document.xml' $docXml
            Add-ZipText $zip 'word/styles.xml' $styles
            Add-ZipText $zip 'word/_rels/document.xml.rels' $docRels
        }
        finally { $zip.Dispose() }
    }
    finally { $fs.Dispose() }
}

# ---- Aufruf: beide Dateien, Fehler gehen ins Fehlerprotokoll und brechen den Lauf nicht ab

Write-Host '==> Excel- und Word-Datei'
Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
$xlsxPath = Get-FreePath (Join-Path $OutputDir ('Kunden-Onboarding-AD-{0}.xlsx' -f $Domain))
$docxPath = Get-FreePath (Join-Path $OutputDir ('AD-Zusammenfassung-{0}.docx' -f $Domain))
$kennzahlen = Get-Kennzahlen

try {
    Build-Xlsx -Path $xlsxPath -Kennzahlen $kennzahlen
    Write-Host "    Excel geschrieben: $xlsxPath"
}
catch { Add-Problem 'Excel-Datei' $_.Exception.Message }

try {
    Build-Docx -Path $docxPath -Kennzahlen $kennzahlen -XlsxName ([System.IO.Path]::GetFileName($xlsxPath))
    Write-Host "    Word geschrieben: $docxPath"
}
catch { Add-Problem 'Word-Datei' $_.Exception.Message }

if ($script:Problems.Count -gt 0) {
    $logPath = Join-Path $OutputDir 'fehlerprotokoll.txt'
    [System.IO.File]::WriteAllLines($logPath, $script:Problems.ToArray(), (New-Object System.Text.UTF8Encoding $false))
}

Write-Host ''
Write-Host "Fertig. Ausgabeordner: $OutputDir"
Write-Host "Übersicht: $mdPath"
Write-Host "Übersicht (HTML): $htmlPath"
Write-Host "Excel: $xlsxPath"
Write-Host "Word: $docxPath"
Write-Host ('Erhebungsfehler / übersprungene Abschnitte: {0}' -f $script:Problems.Count)
Write-Host 'Es wurden ausschließlich lesende Abfragen ausgeführt.'
