# start.ps1 - Phase 1: Erstkonfiguration
# Dieses Skript ist für die erste Konfigurationsphase verantwortlich,
# die unmittelbar nach der Windows-Installation und vor dem ersten Neustart stattfindet.
# Kompatibel mit Windows PowerShell 5.1 (kein ternärer Operator).

# --- FESTE VARIABLEN UND STANDARDS ---
$FixedDomainName    = "EBENESER.lokal"
$DefaultSubnetMask  = "255.255.255.0"
$DefaultGateway     = "192.168.2.1"
$DefaultDns         = "192.168.200.4"

# --- PFADE ---
$ScriptRoot   = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ProgressFile = Join-Path $ScriptRoot "fortschritt.json"
$LogDir       = Join-Path $ScriptRoot "protokolle"
$LogFile      = Join-Path $LogDir    "setup_log.txt"


# --- LOKALE AUSFÜHRUNG (kein Pendrive nach der Installation nötig) ---
$LocalRoot = "C:\SetupScripts\Win11Deploy"

# Wenn start.ps1 aus ...\SetupScripts\ gestartet wird, liegen Konfigurationen\ und Installationsprogramme\ typischerweise
# eine Ebene höher (Deployment-Root).
$DeploymentRoot = Split-Path -Parent $ScriptRoot
$HasSiblingFolders = (Test-Path (Join-Path $DeploymentRoot "Konfigurationen")) -or (Test-Path (Join-Path $DeploymentRoot "Installationsprogramme"))

if ($ScriptRoot -ne $LocalRoot) {
    try {
        if (-not (Test-Path $LocalRoot)) { New-Item -ItemType Directory -Path $LocalRoot -Force | Out-Null }

        # 1) Skripte/Logs aus SetupScripts\ nach Win11Deploy kopieren (damit LocalStart unverändert bleibt)
        Copy-Item -Path (Join-Path $ScriptRoot "*") -Destination $LocalRoot -Recurse -Force -ErrorAction Stop

        # 2) Falls vorhanden: Geschwister-Ordner (Konfigurationen / Installationsprogramme) aus Deployment-Root mitnehmen
        if ($HasSiblingFolders) {
            foreach ($folder in @("Konfigurationen","Installationsprogramme")) {
                $src = Join-Path $DeploymentRoot $folder
                if (Test-Path $src) {
                    $dst = Join-Path $LocalRoot $folder
                    Copy-Item -Path $src -Destination $dst -Recurse -Force -ErrorAction Stop
                }
            }
        }

        # Relaunch local
        $LocalStart = Join-Path $LocalRoot "start.ps1"
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$LocalStart`"" -Verb RunAs
        exit 0
    } catch {
        # Wenn Kopieren fehlschlägt, laufen wir notfalls von der aktuellen Quelle weiter.
    }
}

# Nach Relauch: Pfade auf lokal setzen (nur wenn vorhanden)
if (Test-Path (Join-Path $LocalRoot "start.ps1")) {
    $ScriptRoot   = $LocalRoot
    $ProgressFile = Join-Path $ScriptRoot "fortschritt.json"
    $LogDir       = Join-Path $ScriptRoot "protokolle"
    $LogFile      = Join-Path $LogDir    "setup_log.txt"
}


# --- FUNKTIONEN ---

function Ensure-Directory {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

# Funktion zum Protokollieren von Nachrichten in der Log-Datei und auf der Konsole
function Write-Log {
    param([Parameter(Mandatory=$true)][string]$Message)

    Ensure-Directory -Path $LogDir

    $Timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [start.ps1] - $Message"

    # UTF-8 (Windows PowerShell 5.1 schreibt mit BOM)
    Add-Content -Path $LogFile -Value $LogMessage -Encoding UTF8
    Write-Host $Message
}

# Funktion zum Aktualisieren des Fortschritts in der JSON-Datei (preserviert configuracoesIniciais)
function Update-Progress {
    param(
        [Parameter(Mandatory=$true)][int]$Etapa,
        [Parameter(Mandatory=$true)][string]$Descricao,
        [hashtable]$InitialConfig = $null
    )

    $existing = $null
    if (Test-Path $ProgressFile) {
        try {
            $existing = Get-Content -Path $ProgressFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        } catch {
            $existing = $null
        }
    }

    # Bestimme, welche configuracoesIniciais geschrieben werden sollen
    $configToWrite = $null

    if ($InitialConfig) {
        # Sensible Daten aus dem Log entfernen, bevor sie in die Datei geschrieben werden
        $ConfigForFile = @{}
        foreach ($k in $InitialConfig.Keys) { $ConfigForFile[$k] = $InitialConfig[$k] }

        if ($ConfigForFile.AdminPassword) {
            # SecureString exportierbar machen
            $ConfigForFile.AdminPassword = ConvertFrom-SecureString -SecureString $ConfigForFile.AdminPassword
        }
        if ($ConfigForFile.DomainCredential) {
            # Nur Benutzername speichern
            try { $ConfigForFile.DomainCredential = $ConfigForFile.DomainCredential.UserName } catch { }
        }

        $configToWrite = $ConfigForFile
    } elseif ($existing -and $existing.configuracoesIniciais) {
        $configToWrite = $existing.configuracoesIniciais
    }

    $Progress = [ordered]@{
        etapaAtual = $Etapa
        descricao  = $Descricao
        timestamp  = (Get-Date).ToString("o")  # ISO 8601
    }

    if ($configToWrite) {
        $Progress.configuracoesIniciais = $configToWrite
    }

    # UTF-8 (Windows PowerShell 5.1 schreibt mit BOM)
    $Progress | ConvertTo-Json -Depth 10 | Set-Content -Path $ProgressFile -Encoding UTF8

    Write-Log "Fortschritt aktualisiert: Schritt $Etapa - $Descricao"
}

# Hilfsfunktion zur Umwandlung einer Subnetzmaske in eine Präfixlänge (z.B. 255.255.255.0 -> 24)
function ConvertTo-PrefixLength {
    param([Parameter(Mandatory=$true)][string]$SubnetMask)
    try {
        $bits = 0
        foreach ($octet in $SubnetMask.Split('.')) {
            $b = [Convert]::ToString([int]$octet, 2).PadLeft(8,'0')
            $bits += ($b.ToCharArray() | Where-Object { $_ -eq '1' }).Count
        }
        return $bits
    } catch {
        Write-Log "FEHLER: Ungültige Subnetzmaske: $SubnetMask"
        return $null
    }
}

# Funktion zum Sammeln aller Anfangseinstellungen vom Bediener
function Get-InitialConfiguration {
    param(
        [Parameter(Mandatory=$true)][string]$DefaultSubnetMask,
        [Parameter(Mandatory=$true)][string]$DefaultGateway,
        [Parameter(Mandatory=$true)][string]$DefaultDns,
        [Parameter(Mandatory=$true)][string]$FixedDomainName
    )

    $Config    = @{}
    $Confirmed = $false

    while (-not $Confirmed) {
        Write-Host "--- Sammlung der Anfangsinformationen ---"

        # 1. Computername
        $Config.ComputerName = Read-Host "1/6: Geben Sie den Namen für diesen Computer ein"

        # 2. Administrator-Passwort
        Write-Host "2/6: Legen Sie ein Passwort für den lokalen 'admin'-Benutzer fest"
        $Config.AdminPassword = $null

        while (-not $Config.AdminPassword) {
            $P1 = Read-Host "  - Geben Sie das Administrator-Passwort ein" -AsSecureString
            $P2 = Read-Host "  - Bestätigen Sie das Passwort" -AsSecureString

            $BSTR1 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($P1)
            $BSTR2 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($P2)
            try {
                $P1PlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR1)
                $P2PlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR2)
                if ($P1PlainText -eq $P2PlainText) {
                    $Config.AdminPassword = $P1
                } else {
                    Write-Host "Die Passwörter stimmen nicht überein. Bitte erneut versuchen." -ForegroundColor Red
                }
            } finally {
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR1)
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR2)
            }
        }

        # 3. Maschinenprofil
        $Config.MachineProfile = ""
        while ($Config.MachineProfile -ne 'G' -and $Config.MachineProfile -ne 'D') {
            $Config.MachineProfile = (Read-Host "3/6: Wählen Sie das Maschinenprofil (G für Allgemeine Nutzung / D für Dedizierten Benutzer)").ToUpper()
            if ($Config.MachineProfile -ne 'G' -and $Config.MachineProfile -ne 'D') {
                Write-Host "Ungültige Option. Bitte geben Sie 'G' oder 'D' ein."
            }
        }

        # 4. Netzwerkkonfiguration
        $Config.NetworkType = ""
        while ($Config.NetworkType -ne 'D' -and $Config.NetworkType -ne 'E') {
            $Config.NetworkType = (Read-Host "4/6: DHCP oder statische IP verwenden? (D für DHCP / E für Statisch)").ToUpper()
            if ($Config.NetworkType -ne 'D' -and $Config.NetworkType -ne 'E') {
                Write-Host "Ungültige Option. Bitte geben Sie 'D' oder 'E' ein."
            }
        }

        if ($Config.NetworkType -eq 'E') {
            $Config.IPAddress  = Read-Host "  - Geben Sie die IP-Adresse ein"
            $Config.SubnetMask = $DefaultSubnetMask
            $Config.Gateway    = $DefaultGateway
            $Config.DnsServers = $DefaultDns
        } else {
            $Config.IPAddress  = $null
            $Config.SubnetMask = $null
            $Config.Gateway    = $null
            $Config.DnsServers = $null
        }

        # 5. Domänenbeitritt
        $Config.JoinDomain = ""
        while ($Config.JoinDomain -ne 'J' -and $Config.JoinDomain -ne 'N') {
            $Config.JoinDomain = (Read-Host "5/6: Möchten Sie der Domäne '$FixedDomainName' jetzt beitreten? (J/N)").ToUpper()
            if ($Config.JoinDomain -ne 'J' -and $Config.JoinDomain -ne 'N') {
                Write-Host "Ungültige Option. Bitte geben Sie 'J' oder 'N' ein."
            }
        }

        if ($Config.JoinDomain -eq 'J') {
            $Config.DomainCredential = Get-Credential -UserName "$FixedDomainName\" -Message "Geben Sie die Anmeldeinformationen mit Berechtigung zum Domänenbeitritt ein"
        } else {
            $Config.DomainCredential = $null
        }

        # 6. Office-Installation (automatisch basierend auf dem Profil)
        $Config.InstallOffice2016 = $false
        if ($Config.MachineProfile -eq 'G') {
            $Config.InstallOffice2016 = $true
            Write-Host "6/6: Profil 'Allgemeine Nutzung' ausgewählt. Office 2016 wird automatisch installiert."
        } else {
            Write-Host "6/6: Profil 'Dedizierter Benutzer' ausgewählt. Office 365 wird automatisch installiert."
        }

        # --- Überprüfung und Bestätigung ---
        Write-Host "--- Überprüfung der Konfiguration ---"

        $MachineProfileText = "Dedizierter Benutzer"
        if ($Config.MachineProfile -eq 'G') { $MachineProfileText = "Allgemeine Nutzung" }

        $NetworkTypeText = "Statisch"
        if ($Config.NetworkType -eq 'D') { $NetworkTypeText = "DHCP" }

        $JoinDomainText = "Nein"
        if ($Config.JoinDomain -eq 'J') { $JoinDomainText = "Ja" }

        Write-Host "Computername: $($Config.ComputerName)"
        Write-Host "Admin-Passwort: (Gesetzt, aber nicht angezeigt)"
        Write-Host "Maschinenprofil: $MachineProfileText"
        Write-Host "Netzwerkkonfiguration: $NetworkTypeText"
        if ($Config.NetworkType -eq 'E') {
            Write-Host "  - IP-Adresse: $($Config.IPAddress)"
            Write-Host "  - Subnetzmaske: $($Config.SubnetMask)"
            Write-Host "  - Gateway: $($Config.Gateway)"
            Write-Host "  - DNS-Server: $($Config.DnsServers)"
        }
        Write-Host "Domänenbeitritt: $JoinDomainText"

        if ($Config.MachineProfile -eq 'G') {
            $Office2016Text = "Nein"
            if ($Config.InstallOffice2016) { $Office2016Text = "Ja (Automatisch)" }
            Write-Host "Office 2016 installieren: $Office2016Text"
        } else {
            Write-Host "Office 365 installieren: Ja (Automatisch)"
        }

        $ConfirmChoice = ""
        while ($ConfirmChoice -ne 'J' -and $ConfirmChoice -ne 'N') {
            $ConfirmChoice = (Read-Host "Diese Einstellungen bestätigen? (J/N)").ToUpper()
            if ($ConfirmChoice -ne 'J' -and $ConfirmChoice -ne 'N') {
                Write-Host "Ungültige Option. Bitte geben Sie 'J' oder 'N' ein."
            }
        }

        if ($ConfirmChoice -eq 'J') {
            $Confirmed = $true
        } else {
            Write-Host "Einstellungen nicht bestätigt. Starte die Informationssammlung neu..."
        }
    }

    return $Config
}

function Get-WiredUpAdapter {
    # bevorzugt: Up + physisch + nicht WLAN + keine virtuellen Adapter
    $adapters = Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object {
        $_.Status -eq 'Up' -and
        ($_.PhysicalMediaType -notmatch '802\.11' -and $_.MediaType -ne 'Wi-Fi') -and
        $_.HardwareInterface -eq $true -and
        $_.Name -notmatch 'vEthernet|Hyper-V|Virtual|TAP|VPN|Loopback|Wi-Fi|WLAN'
    }

    if (-not $adapters) {
        # Fallback: nimm den ersten Up-Adapter, der nicht offensichtlich virtuell ist
        $adapters = Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object {
            $_.Status -eq 'Up' -and $_.Name -notmatch 'vEthernet|Hyper-V|Virtual|TAP|VPN|Loopback'
        }
    }

    return ($adapters | Select-Object -First 1)
}

# --- BEGINN DER AUSFÜHRUNG ---

Ensure-Directory -Path $LogDir

Write-Log "--- Beginn von Phase 1: Maschinenkonfiguration ---"

# Sammelt alle Anfangseinstellungen vom Bediener
$InitialConfig = Get-InitialConfiguration -DefaultSubnetMask $DefaultSubnetMask -DefaultGateway $DefaultGateway -DefaultDns $DefaultDns -FixedDomainName $FixedDomainName

Update-Progress -Etapa 1 -Descricao "Anfangskonfiguration der Maschine" -InitialConfig $InitialConfig

# Schritt 1.5: Administratorpasswort festlegen
Write-Log "Lege Passwort für Benutzer 'admin' fest..."
try {
    Write-Log "Versuche, das Passwort mit Set-LocalUser festzulegen..."
    Set-LocalUser -Name "admin" -Password $InitialConfig.AdminPassword -ErrorAction Stop
    Write-Log "Passwort für 'admin' erfolgreich mit Set-LocalUser festgelegt."
} catch {
    Write-Log "Set-LocalUser ist nicht verfügbar oder fehlgeschlagen. Wechsle zur Fallback-Methode (net user)."
    try {
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($InitialConfig.AdminPassword)
        try {
            $PlainTextPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            net user admin $PlainTextPassword 2>&1 | Out-Null
            # Ausgabe kann Fehler enthalten, aber Passwort nicht loggen
            Write-Log "net user Befehl ausgeführt."
        } finally {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        }
        Write-Log "Passwort für 'admin' erfolgreich mit 'net user' festgelegt."
    } catch {
        Write-Log "FEHLER: Sowohl Set-LocalUser als auch 'net user' sind fehlgeschlagen. $_"
    }
}

# 2. Netzwerkkonfiguration
$NetworkConfigType = "Statisch"
if ($InitialConfig.NetworkType -eq 'D') { $NetworkConfigType = "DHCP" }
Write-Log "Netzwerkkonfiguration: $NetworkConfigType"

$NetAdapter = Get-WiredUpAdapter
if (-not $NetAdapter) {
    Write-Log "FEHLER: Kein aktiver (kabelgebundener) Netzwerkadapter gefunden."
} else {
    if ($InitialConfig.NetworkType -eq 'E') {
        $IPAddress  = $InitialConfig.IPAddress
        $SubnetMask = $InitialConfig.SubnetMask
        $Gateway    = $InitialConfig.Gateway
        $DnsServers = $InitialConfig.DnsServers

        Write-Log "Konfiguriere statische IP: $IPAddress"

        try {
            $PrefixLength = ConvertTo-PrefixLength -SubnetMask $SubnetMask
            if ($null -eq $PrefixLength) { throw "Ungültige Subnetzmaske / PrefixLength konnte nicht berechnet werden." }

            # Entfernt alte IPv4-Adressen (sicherer als "alles löschen")
            Get-NetIPAddress -InterfaceIndex $NetAdapter.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object { $_.IPAddress -and $_.IPAddress -notlike '169.254.*' } |
                Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue

            # Konfiguriert IP und Gateway
            New-NetIPAddress -InterfaceIndex $NetAdapter.InterfaceIndex -IPAddress $IPAddress -PrefixLength $PrefixLength -DefaultGateway $Gateway -ErrorAction Stop | Out-Null

            # Konfiguriert die DNS-Server
            $DnsArray = @()
            if ($DnsServers) {
                $DnsArray = $DnsServers.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            }
            if ($DnsArray.Count -gt 0) {
                Set-DnsClientServerAddress -InterfaceIndex $NetAdapter.InterfaceIndex -ServerAddresses $DnsArray -ErrorAction Stop
            }

            Write-Log "Statische Netzwerkkonfiguration erfolgreich angewendet."
        } catch {
            Write-Log "FEHLER beim Anwenden der Netzwerkkonfiguration: $_. Fallback auf DHCP."
            try {
                Set-NetIPInterface -InterfaceIndex $NetAdapter.InterfaceIndex -Dhcp Enabled -ErrorAction SilentlyContinue
                Set-DnsClientServerAddress -InterfaceIndex $NetAdapter.InterfaceIndex -ResetServerAddresses -ErrorAction SilentlyContinue
            } catch { }
        }
    } else {
        Write-Log "Konfiguriere Netzwerk für DHCP."
        try {
            Set-NetIPInterface -InterfaceIndex $NetAdapter.InterfaceIndex -Dhcp Enabled -ErrorAction Stop
            Set-DnsClientServerAddress -InterfaceIndex $NetAdapter.InterfaceIndex -ResetServerAddresses -ErrorAction Stop
            Write-Log "DHCP-Konfiguration angewendet."
        } catch {
            Write-Log "FEHLER beim Aktivieren von DHCP: $_"
        }
    }
}

# 3. Computer umbenennen / Domäne beitreten
$ComputerName     = $InitialConfig.ComputerName
$JoinDomainChoice = $InitialConfig.JoinDomain

if ($JoinDomainChoice -eq 'J') {
    Write-Log "Führe Domänenbeitritt zu $FixedDomainName durch (inkl. Computernamen: $ComputerName)..."
    $Credential = $InitialConfig.DomainCredential
    try {
        # Add-Computer kann gleichzeitig umbenennen und der Domäne beitreten
        Add-Computer -DomainName $FixedDomainName -Credential $Credential -NewName $ComputerName -Force -ErrorAction Stop
        Write-Log "Domänenbeitritt vorbereitet. Neustart erforderlich."
    } catch {
        Write-Log "FEHLER beim Domänenbeitritt: $_"
        # Fallback: nur umbenennen
        try {
            Rename-Computer -NewName $ComputerName -Force -ErrorAction Stop
            Write-Log "Computer umbenannt (ohne Domäne)."
        } catch {
            Write-Log "FEHLER beim Umbenennen: $_"
        }
    }
} else {
    Write-Log "Benenne den Computer um in $ComputerName (ohne Domäne)..."
    try {
        Rename-Computer -NewName $ComputerName -Force -ErrorAction Stop
        Write-Log "Computer umbenannt."
    } catch {
        Write-Log "FEHLER beim Umbenennen: $_"
    }
}

# 4. Plane die Ausführung von Phase 2 (fortsetzen.ps1) als geplante Aufgabe (robust gegen Reboots)
Write-Log "Plane die Ausführung von Phase 2 (fortsetzen.ps1) als geplante Aufgabe (AtStartup, SYSTEM)."

$TaskName  = "Win11SetupFase2"
$Fase2Path = Join-Path $ScriptRoot "fortsetzen.ps1"

try {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Write-Log "Vorhandene geplante Aufgabe '$TaskName' wird entfernt."
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    }
} catch { }

$TaskAction   = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$Fase2Path`""
$TaskTrigger  = New-ScheduledTaskTrigger -AtStartup
$TaskSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfOnBatteries -MultipleInstances IgnoreNew
$TaskPrincipal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

try {
    Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger -Settings $TaskSettings -Principal $TaskPrincipal -Description "Führt Phase 2 (fortsetzen.ps1) der Windows 11 Bereitstellung aus." | Out-Null
    Write-Log "Geplante Aufgabe '$TaskName' erfolgreich erstellt."
} catch {
    Write-Log "FEHLER beim Erstellen der geplanten Aufgabe '$TaskName': $_"
}

# 5. Fortschritt aktualisieren und neu starten (InitialConfig wird bewusst NICHT erneut übergeben, aber bleibt durch Preserve-Logik erhalten)
Write-Log "Erstkonfiguration abgeschlossen. Der Computer wird in 10 Sekunden neu gestartet."
Update-Progress -Etapa 2 -Descricao "Warte auf Neustart, um Phase 2 zu beginnen"

Start-Sleep -Seconds 10
Restart-Computer -Force