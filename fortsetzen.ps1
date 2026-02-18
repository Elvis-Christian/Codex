# fortsetzen.ps1 - Phase 2: Nach-Installation
# Dieses Skript wird nach dem ersten Neustart ausgeführt und ist für die Installation von Anwendungen
# und die endgültige Konfiguration verantwortlich.
# Kompatibel mit Windows PowerShell 5.1 (kein ternärer Operator).
# Läuft robust weiter nach Reboots (Scheduled Task AtStartup / SYSTEM) und ist idempotent.

# --- ANFANGSKONFIGURATION ---
$ScriptRoot     = Split-Path -Parent $MyInvocation.MyCommand.Definition
$DeploymentRoot = Split-Path $ScriptRoot -Parent
$LocalDeploymentRoot = "C:\SetupScripts\Win11Deploy"
if (Test-Path $LocalDeploymentRoot) { $DeploymentRoot = $LocalDeploymentRoot }
$ConfigsPath  = Join-Path $DeploymentRoot "Konfigurationen"
$ProgramsPath = Join-Path $DeploymentRoot "Installationsprogramme"
$ScriptRoot   = Join-Path $DeploymentRoot "SetupScripts"
$ProgressFile   = Join-Path $ScriptRoot "fortschritt.json"
$LogDir         = Join-Path $ScriptRoot "protokolle"
$LogFile        = Join-Path $LogDir    "setup_log.txt"
$ConfigsPath    = Join-Path $DeploymentRoot "Konfigurationen"

# Optional: Umgebungs-/Kundenparameter aus externer Datei (reduziert Hardcode)
$EnvConfig = $null
$EnvConfigPath = Join-Path $ConfigsPath "Umgebung.json"
if (Test-Path $EnvConfigPath) {
    try { $EnvConfig = Get-Content -Path $EnvConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { $EnvConfig = $null }
}


# Scheduled Task Name (muss zu start.ps1 passen!)
$TaskName       = "Win11SetupFase2"

# Abschluss-Marker (damit wir Fortschritt sicher löschen können, ohne Endlosschleifen zu riskieren)
$CompletedFlag  = Join-Path $ScriptRoot "deployment_completed.flag"

# --- FUNKTIONEN ---

function New-DirectoryIfMissing {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}


function Set-EdgePolicyValue {
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)]$Value
    )
    $EdgeKey = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
    if (-not (Test-Path $EdgeKey)) { New-Item -Path $EdgeKey -Force | Out-Null }

    if ($null -eq $Value) { return }

    if ($Value -is [bool]) {
        New-ItemProperty -Path $EdgeKey -Name $Name -Value ([int]$Value) -PropertyType DWord -Force | Out-Null
    } elseif ($Value -is [int] -or $Value -is [long]) {
        New-ItemProperty -Path $EdgeKey -Name $Name -Value ([int]$Value) -PropertyType DWord -Force | Out-Null
    } elseif ($Value -is [string]) {
        New-ItemProperty -Path $EdgeKey -Name $Name -Value $Value -PropertyType String -Force | Out-Null
    } elseif ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [hashtable]) -and -not ($Value -is [pscustomobject])) {
        $arr = @()
        foreach ($v in $Value) { $arr += [string]$v }
        New-ItemProperty -Path $EdgeKey -Name $Name -Value $arr -PropertyType MultiString -Force | Out-Null
    } else {
        $json = $Value | ConvertTo-Json -Compress -Depth 20
        New-ItemProperty -Path $EdgeKey -Name $Name -Value $json -PropertyType String -Force | Out-Null
    }
}

function Apply-EdgePoliciesFromJsonFile {
    param([Parameter(Mandatory=$true)][string]$JsonPath)
    try {
        $pol = Get-Content -Path $JsonPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        foreach ($p in $pol.PSObject.Properties) {
            Set-EdgePolicyValue -Name $p.Name -Value $p.Value
        }
        Write-Log "Edge-Richtlinien via Registry angewendet: $JsonPath"
    } catch {
        Write-Log "FEHLER beim Anwenden der Edge-Richtlinien ($JsonPath): $_"
    }
}


function Write-Log {
    param([Parameter(Mandatory=$true)][string]$Message)
    New-DirectoryIfMissing -Path $LogDir
    $Timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [fortsetzen.ps1] - $Message"
    Add-Content -Path $LogFile -Value $LogMessage -Encoding UTF8
    Write-Host $Message
}

# Fortschritt schreiben + configuracoesIniciais preservieren
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

    $configToWrite = $null

    if ($InitialConfig) {
        $ConfigForFile = @{}
        foreach ($k in $InitialConfig.Keys) { $ConfigForFile[$k] = $InitialConfig[$k] }

        if ($ConfigForFile.AdminPassword) {
            $ConfigForFile.AdminPassword = ConvertFrom-SecureString -SecureString $ConfigForFile.AdminPassword
        }
        if ($ConfigForFile.DomainCredential) {
            try { $ConfigForFile.DomainCredential = $ConfigForFile.DomainCredential.UserName } catch { }
        }

        $configToWrite = $ConfigForFile
    } elseif ($existing -and $existing.configuracoesIniciais) {
        $configToWrite = $existing.configuracoesIniciais
    }

    $Progress = [ordered]@{
        etapaAtual = $Etapa
        descricao  = $Descricao
        timestamp  = (Get-Date).ToString("o")
    }

    if ($configToWrite) { $Progress.configuracoesIniciais = $configToWrite }

    $Progress | ConvertTo-Json -Depth 10 | Set-Content -Path $ProgressFile -Encoding UTF8
    Write-Log "Fortschritt aktualisiert: Schritt $Etapa - $Descricao"
}

function Test-InternetConnection {
    Write-Log "Teste Internetverbindung..."
    try {
        $response = Invoke-WebRequest -Uri "http://www.msftconnecttest.com/connecttest.txt" -UseBasicParsing -TimeoutSec 10
        if ($response.StatusCode -eq 200) {
            Write-Log "Internetverbindung OK."
            return $true
        }
    } catch {
        Write-Log "WARNUNG: Keine Internetverbindung. $_"
        return $false
    }
    return $false
}

function Remove-FileSecurely {
    param([Parameter(Mandatory=$true)][string]$FilePath)

    if (-not (Test-Path $FilePath)) {
        Write-Log "Sicheres Löschen: Datei $FilePath nicht gefunden."
        return
    }

    try {
        Write-Log "Sicheres Löschen für $FilePath wird vorbereitet..."
        $File   = Get-Item $FilePath -ErrorAction Stop
        $Length = [int64]$File.Length

        if ($Length -gt 0) {
            $RandomBytes = New-Object byte[] $Length
            $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
            $rng.GetBytes($RandomBytes)
            [System.IO.File]::WriteAllBytes($FilePath, $RandomBytes)
            Write-Log "Datei $FilePath mit zufälligen Daten überschrieben."
        }

        Remove-Item $FilePath -Force -ErrorAction Stop
        Write-Log "Datei $FilePath endgültig gelöscht."
    } catch {
        Write-Log "FEHLER beim sicheren Löschen der Datei ${FilePath}: $_"
    }
}

function Remove-Phase2Task {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        try {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
            Write-Log "Geplante Aufgabe '$TaskName' erfolgreich entfernt."
        } catch {
            Write-Log "WARNUNG: Konnte geplante Aufgabe '$TaskName' nicht entfernen: $_"
        }
    } else {
        Write-Log "Geplante Aufgabe '$TaskName' ist nicht vorhanden (ok)."
    }
}

# --- BEGINN DER AUSFÜHRUNG ---

Write-Log "--- Beginn von Phase 2: Installation und Konfiguration ---"

# Wenn schon abgeschlossen (Flag existiert), Task entfernen und raus
if (Test-Path $CompletedFlag) {
    Write-Log "Abschluss-Flag gefunden ($CompletedFlag). Prozess ist bereits abgeschlossen."
    Remove-Phase2Task
    exit 0
}

# Fortschritt lesen
try {
    $CurrentProgress = Get-Content -Path $ProgressFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
} catch {
    Write-Log "FEHLER: Fortschrittsdatei '$ProgressFile' konnte nicht gelesen/analysiert werden. $_"
    exit 1
}

$InitialConfig = $null
if ($CurrentProgress -and $CurrentProgress.configuracoesIniciais) {
    $InitialConfig = $CurrentProgress.configuracoesIniciais
} else {
    Write-Log "FEHLER: 'configuracoesIniciais' fehlt im Fortschritt. Phase 2 kann nicht fortsetzen."
    exit 1
}

# Maschinenprofil auslesen (ohne ternär)
$MachineProfile = $InitialConfig.MachineProfile
$MachineProfileText = "Dedizierter Benutzer"
if ($MachineProfile -eq 'G') { $MachineProfileText = "Allgemeine Nutzung" }
Write-Log "Maschinenprofil aus Fortschritt gelesen: $MachineProfileText"

# Wenn bereits über finale Etappe hinaus: Task weg, Flag setzen, raus
if ($CurrentProgress.etapaAtual -ge 11) {
    Write-Log "Letzte Etappe bereits erreicht (etapaAtual=$($CurrentProgress.etapaAtual)). Abschlussroutine."
    try { Set-Content -Path $CompletedFlag -Value (Get-Date).ToString("o") -Encoding UTF8 } catch { }
    Remove-Phase2Task
    exit 0
}

# --- SCHRITT 3: OFFLINE-SOFTWARE-INSTALLATION ---
if ($CurrentProgress.etapaAtual -lt 3) {
    Update-Progress -Etapa 3 -Descricao "Installiere Offline-Software"

    # 3.1 - Microsoft Office
    if ($MachineProfile -eq 'G') {
        Write-Log "Profil 'Allgemeine Nutzung': Überprüfe Installation von Office 2016."
        $OfficePath = $null
        try { $OfficePath = (Get-Command winword -ErrorAction SilentlyContinue).Source } catch { }

        if ($OfficePath) {
            Write-Log "Microsoft Office gefunden in: $OfficePath. Überspringe Installation."
        } else {
            $InstallOffice = $false
            try { $InstallOffice = [bool]$InitialConfig.InstallOffice2016 } catch { $InstallOffice = $false }

            if ($InstallOffice) {
                
$OfficeInstallerPath = Join-Path $DeploymentRoot "Installationsprogramme\Office2016"

# Robust: suche setup.exe rekursiv (die Office-DVD-Struktur liegt oft in Unterordnern)
$OfficeSetup = $null
try {
    $OfficeSetup = Get-ChildItem -Path $OfficeInstallerPath -Filter "setup.exe" -File -Recurse -ErrorAction SilentlyContinue |
                   Select-Object -First 1 -ExpandProperty FullName
} catch { $OfficeSetup = $null }

if (-not $OfficeSetup) {
    Write-Log "FEHLER: Office 2016 Setup nicht gefunden unter: $OfficeInstallerPath"
}

$OfficeArgs  = "/quiet"

                if (Test-Path $OfficeSetup) {
                    Write-Log "Starte Installation von Office 2016..."
                    try {
                        Start-Process -FilePath $OfficeSetup -ArgumentList $OfficeArgs -Wait -ErrorAction Stop
                        Write-Log "Installation von Office 2016 abgeschlossen. Aktiviere..."

                        # Hinweis: Key hier ist sensibel – falls das real ist, besser extern ablegen.
                        $OfficeKey = $null
                        if ($EnvConfig -and $EnvConfig.Office2016Key) { $OfficeKey = [string]$EnvConfig.Office2016Key }
                        if (-not $OfficeKey -or $OfficeKey.Trim() -eq "") { $OfficeKey = $env:OFFICE2016_KEY }
                        if (-not $OfficeKey -or $OfficeKey.Trim() -eq "") {
                            $keyFile = Join-Path $ConfigsPath "office2016_key.txt"
                            if (Test-Path $keyFile) { $OfficeKey = (Get-Content $keyFile -Raw).Trim() }
                        }
                        $OSPPPath  = "C:\Program Files\Microsoft Office\Office16\OSPP.VBS"
                        if ($OfficeKey -and $OfficeKey.Trim() -ne "" -and (Test-Path $OSPPPath)) {
                            Start-Process cscript.exe -ArgumentList "`"$OSPPPath`" /inpkey:$OfficeKey" -Wait | Out-Null
                            Start-Process cscript.exe -ArgumentList "`"$OSPPPath`" /act" -Wait | Out-Null
                            Write-Log "Aktivierung von Office 2016 versucht."
                        } else {
                            Write-Log "WARNUNG: OSPP.VBS nicht gefunden ($OSPPPath)."
                        }
                    } catch {
                        Write-Log "FEHLER bei Office 2016 Installation/Aktivierung: $_"
                    }
                } else {
                    Write-Log "FEHLER: Office 2016 Setup nicht gefunden: $OfficeSetup"
                }
            } else {
                Write-Log "Office 2016 Installation ist deaktiviert (InstallOffice2016=false)."
            }
        }
    } else {
        # Profil D: Office 365 über ODT (online)
        Write-Log "Profil 'Dedizierter Benutzer': Installation von Office 365 über ODT (Online-Download)."

        if (Test-InternetConnection) {
            $Office365InstallerDir = "C:\Temp\Office365Install"
            $ODTDownloadDir        = "C:\Temp\Office365ODT"
            New-DirectoryIfMissing -Path $Office365InstallerDir
            New-DirectoryIfMissing -Path $ODTDownloadDir

            $ODTSelfExtractingExe = Join-Path $ODTDownloadDir "officedeploymenttool.exe"
            $ODTDownloadUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_19628-20192.exe"

            Write-Log "Lade Office Deployment Tool (ODT) herunter..."
            try {
                Invoke-WebRequest -Uri $ODTDownloadUrl -OutFile $ODTSelfExtractingExe -ErrorAction Stop
                Write-Log "ODT erfolgreich heruntergeladen."
            } catch {
                Write-Log "FEHLER beim Herunterladen des ODT: $_"
            }

            Write-Log "Extrahiere ODT-Dateien..."
            try {
                Start-Process -FilePath $ODTSelfExtractingExe -ArgumentList "/quiet /extract:`"$Office365InstallerDir`"" -Wait -ErrorAction Stop
                Write-Log "ODT nach $Office365InstallerDir extrahiert."
            } catch {
                Write-Log "FEHLER beim Extrahieren der ODT-Dateien: $_"
            }

            $ODTSetupExe = Join-Path $Office365InstallerDir "setup.exe"
            $ConfigurationXmlPath = Join-Path $Office365InstallerDir "configuration.xml"
            $ConfigurationXmlContent = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="de-de" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@
            try {
                $ConfigurationXmlContent | Set-Content -Path $ConfigurationXmlPath -Encoding UTF8
            } catch {
                Write-Log "FEHLER beim Schreiben configuration.xml: $_"
            }

            Write-Log "Starte Installation von Office 365..."
            try {
                Start-Process -FilePath $ODTSetupExe -ArgumentList "/configure `"$ConfigurationXmlPath`"" -Wait -ErrorAction Stop
                Write-Log "Installation von Office 365 abgeschlossen."
            } catch {
                Write-Log "FEHLER bei der Installation von Office 365: $_"
            }
        } else {
            Write-Log "WARNUNG: Keine Internetverbindung. Office 365 Installation übersprungen."
        }
    }

    # 3.2 - Microsoft Edge offline (falls benötigt)
    $EdgePath = $null
    try { $EdgePath = (Get-Command msedge -ErrorAction SilentlyContinue).Source } catch { }
    if ($EdgePath) {
        Write-Log "Microsoft Edge bereits vorhanden. Überspringe."
    } else {
        $EdgeInstallerDir = Join-Path $DeploymentRoot "Installationsprogramme\Edge"
        if (Test-Path $EdgeInstallerDir) {
            $EdgeInstaller = Get-ChildItem -Path $EdgeInstallerDir -Filter "*.msi" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($EdgeInstaller) {
                Write-Log "Installiere Microsoft Edge..."
                try {
                    Start-Process "msiexec.exe" -ArgumentList "/i `"$($EdgeInstaller.FullName)`" /quiet" -Wait -ErrorAction Stop
                    Write-Log "Installation von Edge abgeschlossen."
                } catch {
                    Write-Log "FEHLER bei Edge-Installation: $_"
                }
            } else {
                Write-Log "WARNUNG: Keine Edge MSI in $EdgeInstallerDir gefunden."
            }
        } else {
            Write-Log "WARNUNG: Edge-Installer-Verzeichnis nicht gefunden: $EdgeInstallerDir"
        }
    }
}

# --- SCHRITT 4: WINDOWS UPDATE ---
if ($CurrentProgress.etapaAtual -lt 4) {
    Update-Progress -Etapa 4 -Descricao "Führe Windows Update aus"

    if (Test-InternetConnection) {
        Write-Log "Installiere Modul 'PSWindowsUpdate', falls erforderlich..."
        try {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null
        } catch { }

        try {
            if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
                Install-Module -Name PSWindowsUpdate -Force -AcceptLicense -Scope AllUsers -ErrorAction Stop
            }
            Import-Module PSWindowsUpdate -Force -ErrorAction Stop
        } catch {
            Write-Log "FEHLER: PSWindowsUpdate konnte nicht installiert/importiert werden. $_"
        }

        Write-Log "Suche und installiere Windows-Updates (kann lange dauern, inkl. Reboots)..."
        try {
            Get-WindowsUpdate -Install -AcceptAll -IgnoreReboot | Out-File -FilePath $LogFile -Append -Encoding UTF8
            try {
                $rebootNeeded = $false
                if (Get-Command Get-WURebootStatus -ErrorAction SilentlyContinue) {
                    $rebootNeeded = (Get-WURebootStatus -Silent)
                }
                if ($rebootNeeded) {
                    Write-Log "Windows Update verlangt einen Neustart. Starte neu..."
                    Restart-Computer -Force
                    exit 0
                }
            } catch { }
            Write-Log "Windows Update Durchlauf beendet."
        } catch {
            Write-Log "FEHLER beim Windows Update: $_"
        }
    } else {
        Write-Log "WARNUNG: Überspringe Windows Update (keine Internetverbindung)."
    }
}

# --- SCHRITT 6: ANWENDUNGSINSTALLATION (WINGET, als Benutzer) ---
$AppsDoneFlag = Join-Path $DeploymentRoot "apps_completed.flag"

# Gate: Phase 2 darf erst nach abgeschlossener App-Installation weiterlaufen (winget läuft im User-Kontext).
if ($CurrentProgress.etapaAtual -eq 6) {
    if (-not (Test-Path $AppsDoneFlag)) {
        Write-Log "Warte auf Abschluss der App-Installation (Flag fehlt: $AppsDoneFlag)."
        exit 0
    } else {
        Write-Log "App-Installation abgeschlossen (Flag gefunden). Fahre fort..."
        Update-Progress -Etapa 7 -Descricao "Apps abgeschlossen, starte Nachkonfiguration"
        $CurrentProgress.etapaAtual = 7
    }
}

if ($CurrentProgress.etapaAtual -lt 6) {
    Update-Progress -Etapa 6 -Descricao "Plane App-Installation (Winget) beim Admin-Login"

    $AppsTaskName = "Win11SetupApps"
    $AppsScript   = Join-Path (Join-Path $DeploymentRoot "SetupScripts") "apps.ps1"

    if (-not (Test-Path $AppsScript)) {
        Write-Log "FEHLER: apps.ps1 nicht gefunden: $AppsScript"
    } else {
        try { Unregister-ScheduledTask -TaskName $AppsTaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }

        $action    = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$AppsScript`""
        $trigger   = New-ScheduledTaskTrigger -AtLogOn -User "admin"
        try { $trigger.Delay = "PT45S" } catch { }  # kurzer Delay, damit User-Session/Netzwerk sauber steht
        # Hinweis: Für AtLogOn mit Benutzer reicht der UserId; Passwort wird nicht gespeichert, wenn der Benutzer interaktiv (Auto-Login) angemeldet wird.
        $principal = New-ScheduledTaskPrincipal -UserId "admin" -LogonType InteractiveToken -RunLevel Highest
        $settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfOnBatteries -MultipleInstances IgnoreNew

        Register-ScheduledTask -TaskName $AppsTaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
        Write-Log "Apps-Task erstellt: $AppsTaskName (AtLogOn: admin)."

        # Markiere geplant und starte neu, damit Auto-Login/Apps loslegen kann
        try { New-Item -ItemType File -Path (Join-Path $DeploymentRoot "apps_scheduled.flag") -Force | Out-Null } catch { }
        Write-Log "Neustart, damit der Admin-Login die App-Installation startet..."
        Restart-Computer -Force
        exit 0
    }
}

# --- SCHRITT 7: KONFIGURATION NACH DER APP-INSTALLATION ---
if ($CurrentProgress.etapaAtual -lt 7) {
    Update-Progress -Etapa 7 -Descricao "Konfiguriere installierte Anwendungen"

    # 6.1 - RustDesk
    Write-Log "Konfiguriere öffentlichen Schlüssel von RustDesk..."
    $RustDeskKey = "LjpIRqXwDB3m8Zvvan1Th7KF3A0F6R9cgLMCG69KtYQ="
    if ($EnvConfig -and $EnvConfig.RustDeskKey) {
        try { $RustDeskKey = [string]$EnvConfig.RustDeskKey } catch { }
    }
    $RustDeskConfigPath = "C:\Windows\ServiceProfiles\LocalService\AppData\Roaming\RustDesk\config\RustDesk.toml"
    try {
        $RustDeskDir = Split-Path $RustDeskConfigPath -Parent
        if (Test-Path $RustDeskDir) {
            $content = ""
            if (Test-Path $RustDeskConfigPath) { $content = Get-Content $RustDeskConfigPath -Raw -ErrorAction SilentlyContinue }
            if ($content -match '(?m)^\s*key\s*=') {
                $content = [regex]::Replace($content, '(?m)^\s*key\s*=.*$', 'key = "' + $RustDeskKey + '"')
            } else {
                if (-not $content.EndsWith("`n")) { $content += "`n" }
                $content += 'key = "' + $RustDeskKey + '"' + "`n"
            }
            Set-Content -Path $RustDeskConfigPath -Value $content -Encoding UTF8
            Write-Log "RustDesk-Schlüssel konfiguriert."
        } else {
            Write-Log "WARNUNG: RustDesk-Verzeichnis nicht gefunden: $RustDeskDir"
        }
    } catch {
        Write-Log "FEHLER bei RustDesk-Konfiguration: $_"
    }

    # 6.2 - Zabbix Agent
    Write-Log "Konfiguriere Zabbix Agent..."
    $ZabbixServerIP   = "192.168.2.165"
    if ($EnvConfig -and $EnvConfig.ZabbixServerIP) {
        try { $ZabbixServerIP = [string]$EnvConfig.ZabbixServerIP } catch { }
    }
    $ZabbixConfigPath = "C:\Program Files\Zabbix Agent\zabbix_agentd.conf"
    if (Test-Path $ZabbixConfigPath) {
        try {
            (Get-Content $ZabbixConfigPath) |
                ForEach-Object { $_ -replace "^Server=127.0.0.1",       "Server=${ZabbixServerIP}" } |
                ForEach-Object { $_ -replace "^ServerActive=127.0.0.1", "ServerActive=${ZabbixServerIP}" } |
                Set-Content -Path $ZabbixConfigPath -Encoding UTF8

            Write-Log "Zabbix-Konfigurationsdatei aktualisiert (Server=${ZabbixServerIP})."

            Restart-Service "Zabbix Agent" -ErrorAction SilentlyContinue
            if ($?) { Write-Log "Zabbix Agent-Dienst neu gestartet." } else { Write-Log "WARNUNG: Zabbix Agent-Dienst konnte nicht neu gestartet werden." }
        } catch {
            Write-Log "FEHLER bei Zabbix-Konfiguration: $_"
        }
    } else {
        Write-Log "WARNUNG: Zabbix-Konfigurationsdatei nicht gefunden: $ZabbixConfigPath"
    }

    # 6.3 - Microsoft Edge Policies
    Write-Log "Konfiguriere Microsoft Edge entsprechend dem Maschinenprofil..."
    $EdgePolicySourceFile = $null
    if ($MachineProfile -eq 'G') {
        $EdgePolicySourceFile = Join-Path $ConfigsPath "EdgeRichtlinien_Geral.json"
        Write-Log "Wende Edge-Richtlinien für Profil 'Allgemeine Nutzung' an."
    } else {
        $EdgePolicySourceFile = Join-Path $ConfigsPath "EdgeRichtlinien_Dedicado.json"
        Write-Log "Wende Edge-Richtlinien für Profil 'Dedizierter Benutzer' an."
    }

    if ($EdgePolicySourceFile -and (Test-Path $EdgePolicySourceFile)) {
        Apply-EdgePoliciesFromJsonFile -JsonPath $EdgePolicySourceFile
    } else {
        Write-Log "WARNUNG: Edge-Richtlinien-Datei nicht gefunden: $EdgePolicySourceFile"
    }
}

# --- SCHRITT 8: ERSTELLUNG VON VERKNÜPFUNGEN ---
if ($CurrentProgress.etapaAtual -lt 8) {
    Update-Progress -Etapa 8 -Descricao "Erstelle Verknüpfungen auf dem Desktop"

    $ShortcutsFile = Join-Path $ConfigsPath "Shortcuts.txt"
    if (Test-Path $ShortcutsFile) {
        $PublicDesktop = [Environment]::GetFolderPath("CommonDesktopDirectory")
        $shell = $null
        try { $shell = New-Object -ComObject WScript.Shell } catch { }

        Get-Content $ShortcutsFile |
            Where-Object { $_ -notmatch '^\s*#' -and $_.Trim() -ne '' } |
            ForEach-Object {
                $parts = $_.Split(';')
                $name = $parts[0]
                $target = $parts[1]
                $arguments = $null
                if ($parts.Count -ge 3) { $arguments = $parts[2] }

                Write-Log "Erstelle Verknüpfung: $name"
                try {
                    if (-not $shell) { throw "COM WScript.Shell nicht verfügbar." }
                    $shortcut = $shell.CreateShortcut("$PublicDesktop\$name.lnk")
                    $shortcut.TargetPath = $target
                    if ($arguments) { $shortcut.Arguments = $arguments }
                    $shortcut.Save()
                } catch {
                    Write-Log "FEHLER beim Erstellen der Verknüpfung '$name': $_"
                }
            }
    } else {
        Write-Log "WARNUNG: Shortcuts-Datei nicht gefunden: $ShortcutsFile"
    }
}

# --- SCHRITT 9: ENTFERNT ---
# Der Schritt zur Vorbereitung des HTA-Druckerinstallationsprogramms wurde entfernt.

# --- SCHRITT 10: FERNZUGRIFF (RDP) AKTIVIEREN ---
if ($CurrentProgress.etapaAtual -lt 10) {
    Update-Progress -Etapa 10 -Descricao "Aktiviere Fernzugriff (RDP)"

    Write-Log "Aktiviere Remotedesktop..."
    try {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 0 -Force -ErrorAction Stop
        Write-Log "RDP in der Registrierung aktiviert."
    } catch {
        Write-Log "FEHLER beim Aktivieren von RDP in der Registry: $_"
    }

    Write-Log "Aktiviere Firewall-Regel für Remotedesktop..."
    try {
        Enable-NetFirewallRule -DisplayGroup "Remote Desktop" -ErrorAction Stop
        Write-Log "Firewall-Regeln für RDP aktiviert."
    } catch {
        Write-Log "FEHLER beim Aktivieren der Firewall-Regeln für RDP: $_"
    }
}

# --- SCHRITT 11: ABSCHLUSS ---
if ($CurrentProgress.etapaAtual -lt 11) {
    Update-Progress -Etapa 11 -Descricao "Bereitstellung abgeschlossen"
    Write-Log "--- BEREITSTELLUNGSPROZESS ABGESCHLOSSEN ---"
}

# --- SCHRITT 12: AUFRÄUMEN + TASK ENTFERNEN (nicht interaktiv) ---
Update-Progress -Etapa 12 -Descricao "Bereinigung und Abschluss"

try {
    Set-Content -Path $CompletedFlag -Value (Get-Date).ToString("o") -Encoding UTF8
    Write-Log "Abschluss-Flag gesetzt: $CompletedFlag"
} catch {
    Write-Log "WARNUNG: Konnte Abschluss-Flag nicht schreiben: $_"
}

Write-Log "Entferne geplante Aufgabe '$TaskName' nach erfolgreichem Abschluss."
Remove-Phase2Task

Write-Log "Beginne mit der sicheren Bereinigung von temporären Konfigurationsdateien..."
Remove-FileSecurely -FilePath $ProgressFile
Write-Log "Sichere Bereinigung abgeschlossen."

Write-Log "Phase 2 beendet. Kein Prompt / keine Interaktion erforderlich."
exit 0