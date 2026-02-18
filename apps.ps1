# apps.ps1 - App-Installation (Winget) im Benutzerkontext (admin)
# Läuft als Scheduled Task "Win11SetupApps" beim Login des Users "admin"

$DeploymentRoot = "C:\SetupScripts\Win11Deploy"
$ProgressFile   = Join-Path $DeploymentRoot "fortschritt.json"
$LogDir         = Join-Path $DeploymentRoot "protokolle"
$LogFile        = Join-Path $LogDir "setup_log.txt"
$ConfigsPath    = Join-Path $DeploymentRoot "Konfigurationen"
$TaskName       = "Win11SetupApps"

function Ensure-Directory {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

function Write-Log {
    param([Parameter(Mandatory=$true)][string]$Message)
    Ensure-Directory -Path $LogDir
    $Timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [apps.ps1] - $Message"
    Add-Content -Path $LogFile -Value $LogMessage -Encoding UTF8
    Write-Host $Message
}

function Update-Progress {
    param(
        [Parameter(Mandatory=$true)][int]$Etapa,
        [Parameter(Mandatory=$true)][string]$Descricao
    )

    $existing = $null
    if (Test-Path $ProgressFile) {
        try { $existing = Get-Content -Path $ProgressFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { $existing = $null }
    }

    $toWrite = [ordered]@{
        etapaAtual = $Etapa
        descricao  = $Descricao
        timestamp  = (Get-Date).ToString("o")
        configuracoesIniciais = $null
    }

    if ($existing -and $existing.configuracoesIniciais) {
        $toWrite.configuracoesIniciais = $existing.configuracoesIniciais
    }

    $toWrite | ConvertTo-Json -Depth 10 | Set-Content -Path $ProgressFile -Encoding UTF8
}

function Test-InternetConnection {
    try {
        return (Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue)
    } catch { return $false }
}

Write-Log "--- Beginn Apps (Winget) ---"

if (-not (Test-Path $DeploymentRoot)) {
    Write-Log "FEHLER: DeploymentRoot nicht gefunden: $DeploymentRoot"
    exit 1
}

# winget verfügbar?
$wingetCmd = $null
try { $wingetCmd = Get-Command winget -ErrorAction SilentlyContinue } catch { }
if (-not $wingetCmd) {
    Write-Log "FEHLER: winget wurde im Benutzerkontext nicht gefunden. Abbruch."
    exit 2
}

if (-not (Test-InternetConnection)) {
    Write-Log "WARNUNG: Keine Internetverbindung. Überspringe Winget."
    Update-Progress -Etapa 6 -Descricao "Winget übersprungen (kein Internet)"
} else {
    Update-Progress -Etapa 5 -Descricao "Installiere Anwendungen über Winget (admin)"

    try {
        Write-Log "winget --version: $((winget --version) | Out-String)"
    } catch { }

    try {
        Write-Log "Winget Upgrade aller Pakete..."
        winget upgrade --all --silent --accept-source-agreements --accept-package-agreements | Out-Null
    } catch {
        Write-Log "WARNUNG: winget upgrade Problem: $_"
    }

    $AppsFile = Join-Path $ConfigsPath "anwendungen.txt"
    if (Test-Path $AppsFile) {
        $AppsToInstall = Get-Content $AppsFile | Where-Object { $_ -notmatch '^\s*#' -and $_.Trim() -ne '' }
        foreach ($app in $AppsToInstall) {
            Write-Log "Installiere '$app'..."
            try {
                # Edge: Tiny11 hat häufig keinen Store-Stack; die msstore-ID kann dann scheitern.
                if ($app -match "^XPFFTQ037JWMHS$") {
                    Write-Log "Spezialfall Edge: Versuche zuerst Winget-Repo ID Microsoft.Edge..."
                    winget install --id Microsoft.Edge --source winget --silent --accept-source-agreements --accept-package-agreements | Out-Null
                    if ($LASTEXITCODE -ne 0) {
                        Write-Log "WARNUNG: Winget-Repo Edge fehlgeschlagen (ExitCode=$LASTEXITCODE). Versuche msstore-ID $app..."
                        winget install --id $app --silent --accept-source-agreements --accept-package-agreements | Out-Null
                    }

                    if ($LASTEXITCODE -ne 0) {
                        $EdgeMsi = Join-Path $DeploymentRoot "Installationsprogramme\Edge\MicrosoftEdgeEnterpriseX64.msi"
                        if (Test-Path $EdgeMsi) {
                            Write-Log "Fallback: Installiere Edge offline via MSI: $EdgeMsi"
                            Start-Process msiexec.exe -ArgumentList "/i `"$EdgeMsi`" /qn /norestart" -Wait | Out-Null
                        } else {
                            Write-Log "FEHLER: Edge konnte weder via Winget noch offline installiert werden (MSI fehlt).";
                        }
                    }
                } else {
                    winget install --id $app --silent --accept-source-agreements --accept-package-agreements | Out-Null
                }

                if ($LASTEXITCODE -ne 0) {
                    Write-Log "FEHLER: winget install ExitCode=$LASTEXITCODE für '$app'."
                }
            } catch {
                Write-Log "FEHLER: Ausnahme bei winget install '$app': $_"
            }
        }
        Update-Progress -Etapa 6 -Descricao "Winget App-Installation abgeschlossen"
    } else {
        Write-Log "WARNUNG: Anwendungen-Datei nicht gefunden: $AppsFile"
        Update-Progress -Etapa 6 -Descricao "Winget abgeschlossen (keine Anwendungen-Datei)"
    }
}

# Task entfernen
try {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    Write-Log "Apps-Task entfernt: $TaskName"
} catch { }


# Abschluss-Flag setzen, damit fortsetzen.ps1 weiterlaufen darf
try {
    New-Item -ItemType File -Path (Join-Path $DeploymentRoot "apps_completed.flag") -Force | Out-Null
    Write-Log "Apps-Abschluss-Flag geschrieben: $(Join-Path $DeploymentRoot "apps_completed.flag")"
} catch {
    Write-Log "WARNUNG: Konnte apps_completed.flag nicht schreiben: $_"
}

Write-Log "Neustart, um mit Phase 2 fortzusetzen..."
Restart-Computer -Force