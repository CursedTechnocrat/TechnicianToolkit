<#
.SYNOPSIS
    C.I.P.H.E.R. — Configures & Implements Policy-based Hardware Encryption & Recovery
    BitLocker Drive Encryption Tool for PowerShell 5.1+

.DESCRIPTION
    Manages BitLocker drive encryption across all volumes on the local machine.
    Check encryption status, enable or disable encryption, back up recovery keys
    to Active Directory or Entra ID, view recovery key IDs, and suspend or resume
    BitLocker protection.

.USAGE
    PS C:\> .\cipher.ps1                                           # Must be run as Administrator
    PS C:\> .\cipher.ps1 -WhatIf                                   # Preview actions without making changes
    PS C:\> .\cipher.ps1 -Unattended -Action Status                # Show drive status and exit
    PS C:\> .\cipher.ps1 -Unattended -Action Disable -Drive C      # Disable BitLocker on C:
    PS C:\> .\cipher.ps1 -Unattended -Action Suspend -Drive C      # Suspend BitLocker on C:
    PS C:\> .\cipher.ps1 -Unattended -Action BackupAD -Drive C     # Backup recovery key to AD

.NOTES
    Version : 1.0

    Tools Available
    ─────────────────────────────────────────────────────────────────
    G.R.I.M.O.I.R.E.       — Technician Toolkit hub and central launcher
    R.U.N.E.P.R.E.S.S.     — Printer driver installation & configuration
    R.E.S.T.O.R.A.T.I.O.N. — Windows Update management
    C.O.N.J.U.R.E.         — Software deployment via winget / Chocolatey
    A.U.S.P.E.X.           — System diagnostics & HTML report generation
    C.O.V.E.N.A.N.T.       — Machine onboarding & Entra ID domain join
    R.E.V.E.N.A.N.T.       — Profile migration & data transfer
    C.I.P.H.E.R.           — BitLocker drive encryption management
    W.A.R.D.               — User account & local security audit
    A.R.C.H.I.V.E.         — Pre-reimaging profile backup
    A.R.T.I.F.A.C.T.       — Certificate health & SSL expiry monitoring
    H.E.A.R.T.H.           — Toolkit setup & configuration wizard

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success messages
    Yellow   Warnings and cautions
    Red      Critical errors
    Gray     Information and details
#>

param(
    [switch]$Unattended,
    [switch]$WhatIf,
    [ValidateSet('Status','Enable','Disable','Suspend','Resume','BackupAD','BackupEntraID')]
    [string]$Action = "Status",
    [string]$Drive  = "C",
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

# ===========================
# SHARED MODULE BOOTSTRAP
# ===========================
$TKModulePath = Join-Path $PSScriptRoot 'TechnicianToolkit.psm1'
if (-not (Test-Path $TKModulePath)) {
    $TKModuleUrl = 'https://raw.githubusercontent.com/CursedTechnocrat/TechnicianToolkit/main/TechnicianToolkit.psm1'
    Write-Host "  [*] Shared module TechnicianToolkit.psm1 not found - downloading from GitHub..." -ForegroundColor Magenta
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-RestMethod -Uri $TKModuleUrl -OutFile $TKModulePath -ErrorAction Stop
        $parseErrors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile($TKModulePath, [ref]$null, [ref]$parseErrors)
        if ($parseErrors.Count -gt 0) {
            Remove-Item -Path $TKModulePath -Force -ErrorAction SilentlyContinue
            Write-Host "  [!!] Downloaded module failed syntax validation - file removed." -ForegroundColor Red
            Write-Host "       $($parseErrors[0].Message)" -ForegroundColor Red
            exit 1
        }
        Write-Host "  [+] Module downloaded and verified." -ForegroundColor Green
    } catch {
        Write-Host "  [!!] Could not download TechnicianToolkit.psm1:" -ForegroundColor Red
        Write-Host "       $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "       Place the module manually next to this script from:" -ForegroundColor Yellow
        Write-Host "       $TKModuleUrl" -ForegroundColor Yellow
        exit 1
    }
}
Import-Module $TKModulePath -Force -ErrorAction Stop
Assert-AdminPrivilege

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $PSScriptRoot) }

# ─────────────────────────────────────────────────────────────────────────────
# COLOR SCHEMA
# ─────────────────────────────────────────────────────────────────────────────

$ColorSchema = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
    Accent   = 'Blue'
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-CipherBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

   ██████╗██╗██████╗ ██╗  ██╗███████╗██████╗
  ██╔════╝██║██╔══██╗██║  ██║██╔════╝██╔══██╗
  ██║     ██║██████╔╝███████║█████╗  ██████╔╝
  ██║     ██║██╔═══╝ ██╔══██║██╔══╝  ██╔══██╗
  ╚██████╗██║██║     ██║  ██║███████╗██║  ██║
   ╚═════╝╚═╝╚═╝     ╚═╝  ╚═╝╚══════╝╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    C.I.P.H.E.R. — Configures & Implements Policy-based Hardware Encryption & Recovery" -ForegroundColor Cyan
    Write-Host "    BitLocker Drive Encryption Management Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# DRIVE STATUS DISPLAY
# ─────────────────────────────────────────────────────────────────────────────

function Show-DriveStatus {
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  DRIVE ENCRYPTION STATUS" -ForegroundColor $ColorSchema.Header
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host ""

    try {
        $volumes = Get-BitLockerVolume -ErrorAction Stop

        foreach ($vol in $volumes) {
            $statusColor = switch ($vol.VolumeStatus) {
                "FullyEncrypted"       { $ColorSchema.Success  }
                "FullyDecrypted"       { $ColorSchema.Warning  }
                "EncryptionInProgress" { $ColorSchema.Progress }
                "DecryptionInProgress" { $ColorSchema.Progress }
                default                { $ColorSchema.Info     }
            }
            $protColor = if ($vol.ProtectionStatus -eq "On") { $ColorSchema.Success } else { $ColorSchema.Warning }
            $keyTypes  = if ($vol.KeyProtector.Count -gt 0) {
                ($vol.KeyProtector | ForEach-Object { $_.KeyProtectorType }) -join ', '
            } else { "None" }

            Write-Host "  Drive $($vol.MountPoint)" -ForegroundColor $ColorSchema.Header
            Write-Host ("    Status      : {0}" -f $vol.VolumeStatus) -ForegroundColor $statusColor
            Write-Host ("    Protection  : {0}" -f $vol.ProtectionStatus) -ForegroundColor $protColor
            Write-Host ("    Encryption  : {0}%" -f $vol.EncryptionPercentage) -ForegroundColor $ColorSchema.Info
            Write-Host ("    Key Types   : {0}" -f $keyTypes) -ForegroundColor $ColorSchema.Info
            Write-Host ""
        }
    }
    catch {
        Write-Host "  [-] Unable to retrieve BitLocker information: $_" -ForegroundColor $ColorSchema.Error
        Write-Host "  [!!] BitLocker may not be available on this edition of Windows." -ForegroundColor $ColorSchema.Warning
        Write-Host ""
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: SELECT A DRIVE
# ─────────────────────────────────────────────────────────────────────────────

function Select-Drive {
    param([string]$Prompt = "Enter drive letter (e.g. C)")
    Write-Host ""
    Write-Host -NoNewline "  $Prompt`: " -ForegroundColor $ColorSchema.Header
    $userInput = (Read-Host).Trim().ToUpper().TrimEnd(':')
    $mountPoint = "$userInput`:"

    try {
        return Get-BitLockerVolume -MountPoint $mountPoint -ErrorAction Stop
    }
    catch {
        Write-Host ""
        Write-Host "  [-] Drive $mountPoint not found or BitLocker unavailable." -ForegroundColor $ColorSchema.Error
        return $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# ACTION FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Enable-DriveEncryption {
    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  ENABLE BITLOCKER" -ForegroundColor $ColorSchema.Header
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host ""
    Write-Host "  [!!] Enabling BitLocker will begin encrypting the selected drive." -ForegroundColor $ColorSchema.Warning
    Write-Host "       A recovery password will always be generated — save it." -ForegroundColor $ColorSchema.Warning

    $vol = Select-Drive -Prompt "Drive letter to encrypt"
    if (-not $vol) { return }

    if ($vol.VolumeStatus -eq "FullyEncrypted") {
        Write-Host ""
        Write-Host "  [!!] Drive $($vol.MountPoint) is already fully encrypted." -ForegroundColor $ColorSchema.Warning

        $existingKey = $vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } | Select-Object -Last 1
        if ($existingKey) {
            Write-Host ""
            Write-Host "  [+] Recovery key is already present." -ForegroundColor $ColorSchema.Success
            Write-Host "  ID  : $($existingKey.KeyProtectorId)" -ForegroundColor $ColorSchema.Warning
            Write-Host "  Key : $($existingKey.RecoveryPassword)" -ForegroundColor $ColorSchema.Warning
        } else {
            Write-Host "  [!!] No recovery key found — adding one now." -ForegroundColor $ColorSchema.Warning
            try {
                $vol = Add-BitLockerKeyProtector -MountPoint $vol.MountPoint -RecoveryPasswordProtector -ErrorAction Stop
                $newKey = $vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } | Select-Object -Last 1
                if ($newKey) {
                    Write-Host ""
                    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Warning
                    Write-Host "  RECOVERY KEY — SAVE THIS BEFORE CONTINUING" -ForegroundColor $ColorSchema.Warning
                    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Warning
                    Write-Host ""
                    Write-Host "  ID  : $($newKey.KeyProtectorId)" -ForegroundColor $ColorSchema.Warning
                    Write-Host "  Key : $($newKey.RecoveryPassword)" -ForegroundColor $ColorSchema.Warning
                    Write-Host ""
                    Read-Host "  Press Enter once you have saved the recovery key"
                    Write-Host "  [+] Recovery key added successfully." -ForegroundColor $ColorSchema.Success
                }
            } catch {
                Write-Host "  [-] Failed to add recovery key: $_" -ForegroundColor $ColorSchema.Error
            }
        }
        if ($vol.ProtectionStatus -ne "On") {
            Write-Host ""
            Write-Host "  [!!] BitLocker protection is OFF — encryption exists but keys are exposed." -ForegroundColor $ColorSchema.Warning
            Write-Host "  [*] Activating BitLocker protection..." -ForegroundColor $ColorSchema.Progress
            try {
                Resume-BitLocker -MountPoint $vol.MountPoint -ErrorAction Stop | Out-Null
                Write-Host "  [+] BitLocker protection is now ON." -ForegroundColor $ColorSchema.Success
            } catch {
                Write-Host "  [-] Failed to activate protection: $_" -ForegroundColor $ColorSchema.Error
            }
        } else {
            Write-Host "  [+] BitLocker protection is ON." -ForegroundColor $ColorSchema.Success
        }

        Write-Host ""
        return
    }

    Write-Host ""
    Write-Host "  Additional key protector:" -ForegroundColor $ColorSchema.Info
    Write-Host "  [1] TPM only  (no PIN, transparent to user)" -ForegroundColor $ColorSchema.Info
    Write-Host "  [2] TPM + PIN (recommended for high security)" -ForegroundColor $ColorSchema.Info
    Write-Host "  [3] Recovery password only  (no TPM required)" -ForegroundColor $ColorSchema.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
    $protChoice = (Read-Host).Trim()

    if ($WhatIf) {
        $protName = switch ($protChoice) {
            '1' { 'TPM' }; '2' { 'TPM + PIN' }; '3' { 'Recovery password only' }
            default { 'Recovery password only' }
        }
        Write-Host ""
        Write-Host "  [~] Would add recovery password protector to $($vol.MountPoint)" -ForegroundColor Cyan
        Write-Host "  [~] Would add $protName protector to $($vol.MountPoint)" -ForegroundColor Cyan
        Write-Host "  [~] Would start XtsAes256 encryption (used space only) on $($vol.MountPoint)" -ForegroundColor Cyan
        Write-Host ""
        return
    }

    try {
        Write-Host ""
        Write-Host "  [*] Adding recovery password protector..." -ForegroundColor $ColorSchema.Progress
        $vol = Add-BitLockerKeyProtector -MountPoint $vol.MountPoint -RecoveryPasswordProtector -ErrorAction Stop

        $recoveryKey = $vol.KeyProtector |
            Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } |
            Select-Object -Last 1

        if ($recoveryKey) {
            Write-Host ""
            Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Warning
            Write-Host "  RECOVERY KEY — SAVE THIS BEFORE CONTINUING" -ForegroundColor $ColorSchema.Warning
            Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Warning
            Write-Host ""
            Write-Host "  Key ID : $($recoveryKey.KeyProtectorId)" -ForegroundColor $ColorSchema.Warning
            Write-Host "  Key    : $($recoveryKey.RecoveryPassword)" -ForegroundColor $ColorSchema.Warning
            Write-Host ""
            Read-Host "  Press Enter once you have saved the recovery key"
        }

        switch ($protChoice) {
            "1" {
                Write-Host "  [*] Adding TPM protector..." -ForegroundColor $ColorSchema.Progress
                Add-BitLockerKeyProtector -MountPoint $vol.MountPoint -TpmProtector -ErrorAction Stop | Out-Null
            }
            "2" {
                Write-Host ""
                Write-Host -NoNewline "  Enter PIN (6-20 digits): " -ForegroundColor $ColorSchema.Header
                $pin = Read-Host -AsSecureString
                Write-Host "  [*] Adding TPM + PIN protector..." -ForegroundColor $ColorSchema.Progress
                Add-BitLockerKeyProtector -MountPoint $vol.MountPoint -TpmAndPinProtector -Pin $pin -ErrorAction Stop | Out-Null
            }
            "3" {
                Write-Host "  [*] Using recovery password only — no TPM protector added." -ForegroundColor $ColorSchema.Info
            }
            default {
                Write-Host "  [!!] Invalid selection — recovery password protector added only." -ForegroundColor $ColorSchema.Warning
            }
        }

        Write-Host "  [*] Starting encryption on $($vol.MountPoint)..." -ForegroundColor $ColorSchema.Progress
        Enable-BitLocker -MountPoint $vol.MountPoint -EncryptionMethod XtsAes256 -UsedSpaceOnly -ErrorAction Stop | Out-Null

        $volCheck = Get-BitLockerVolume -MountPoint $vol.MountPoint -ErrorAction SilentlyContinue
        if ($volCheck -and $volCheck.ProtectionStatus -ne "On") {
            Write-Host "  [*] Activating BitLocker protection..." -ForegroundColor $ColorSchema.Progress
            Resume-BitLocker -MountPoint $vol.MountPoint -ErrorAction SilentlyContinue | Out-Null
        }

        Write-Host "  [+] Encryption started on $($vol.MountPoint). BitLocker protection is ON." -ForegroundColor $ColorSchema.Success
    }
    catch {
        Write-Host "  [-] Failed to enable BitLocker: $_" -ForegroundColor $ColorSchema.Error
    }

    Write-Host ""
}

function Disable-DriveEncryption {
    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  DISABLE BITLOCKER" -ForegroundColor $ColorSchema.Header
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host ""
    Write-Host "  [!!] This will fully decrypt the drive. Decryption cannot be undone quickly." -ForegroundColor $ColorSchema.Warning
    Write-Host ""
    Write-Host -NoNewline "  Are you sure? (Y/N): " -ForegroundColor $ColorSchema.Warning
    $confirm = (Read-Host).Trim().ToUpper()

    if ($confirm -ne "Y") {
        Write-Host "  [*] Operation cancelled." -ForegroundColor $ColorSchema.Info
        return
    }

    $vol = Select-Drive -Prompt "Drive letter to decrypt"
    if (-not $vol) { return }

    if ($WhatIf) {
        Write-Host ""
        Write-Host "  [~] Would start decryption (disable BitLocker) on $($vol.MountPoint)" -ForegroundColor Cyan
        Write-Host ""
        return
    }

    try {
        Write-Host ""
        Write-Host "  [*] Starting decryption on $($vol.MountPoint)..." -ForegroundColor $ColorSchema.Progress
        Disable-BitLocker -MountPoint $vol.MountPoint -ErrorAction Stop | Out-Null
        Write-Host "  [+] Decryption started on $($vol.MountPoint). This runs in the background." -ForegroundColor $ColorSchema.Success
    }
    catch {
        Write-Host "  [-] Failed to disable BitLocker: $_" -ForegroundColor $ColorSchema.Error
    }

    Write-Host ""
}

function Backup-RecoveryKey {
    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  BACKUP RECOVERY KEY" -ForegroundColor $ColorSchema.Header
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host ""
    Write-Host "  [1] Active Directory" -ForegroundColor $ColorSchema.Info
    Write-Host "  [2] Entra ID (Azure AD)" -ForegroundColor $ColorSchema.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
    $backupChoice = (Read-Host).Trim()

    $vol = Select-Drive -Prompt "Drive letter"
    if (-not $vol) { return }

    $keyProtector = $vol.KeyProtector |
        Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } |
        Select-Object -First 1

    if (-not $keyProtector) {
        Write-Host ""
        Write-Host "  [-] No recovery password protector found on $($vol.MountPoint)." -ForegroundColor $ColorSchema.Error
        Write-Host "  [!!] Enable BitLocker with a recovery password first." -ForegroundColor $ColorSchema.Warning
        return
    }

    try {
        if ($backupChoice -eq "1") {
            Write-Host ""
            Write-Host "  [*] Backing up to Active Directory..." -ForegroundColor $ColorSchema.Progress
            Backup-BitLockerKeyProtector -MountPoint $vol.MountPoint -KeyProtectorId $keyProtector.KeyProtectorId -ErrorAction Stop | Out-Null
            Write-Host "  [+] Recovery key backed up to Active Directory." -ForegroundColor $ColorSchema.Success
        }
        elseif ($backupChoice -eq "2") {
            Write-Host ""
            Write-Host "  [*] Backing up to Entra ID (Azure AD)..." -ForegroundColor $ColorSchema.Progress
            BackupToAAD-BitLockerKeyProtector -MountPoint $vol.MountPoint -KeyProtectorId $keyProtector.KeyProtectorId -ErrorAction Stop | Out-Null
            Write-Host "  [+] Recovery key backed up to Entra ID." -ForegroundColor $ColorSchema.Success
        }
        else {
            Write-Host ""
            Write-Host "  [-] Invalid selection." -ForegroundColor $ColorSchema.Error
        }
    }
    catch {
        Write-Host "  [-] Backup failed: $_" -ForegroundColor $ColorSchema.Error
        Write-Host "  [!!] Ensure this machine is domain-joined or Entra ID-joined and connected." -ForegroundColor $ColorSchema.Warning
    }

    Write-Host ""
}

function Show-RecoveryKey {
    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  SHOW RECOVERY KEY" -ForegroundColor $ColorSchema.Header
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header

    $vol = Select-Drive -Prompt "Drive letter"
    if (-not $vol) { return }

    $recoveryKeys = $vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }

    if (-not $recoveryKeys) {
        Write-Host ""
        Write-Host "  [-] No recovery password found on $($vol.MountPoint)." -ForegroundColor $ColorSchema.Error
        Write-Host ""
        return
    }

    Write-Host ""
    Write-Host "  Recovery key(s) for $($vol.MountPoint):" -ForegroundColor $ColorSchema.Warning
    Write-Host ""

    foreach ($key in $recoveryKeys) {
        Write-Host "  ID  : $($key.KeyProtectorId)" -ForegroundColor $ColorSchema.Warning
        Write-Host "  Key : $($key.RecoveryPassword)" -ForegroundColor $ColorSchema.Warning
        Write-Host ""
    }
}

function Suspend-DriveProtection {
    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  SUSPEND BITLOCKER" -ForegroundColor $ColorSchema.Header
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host ""
    Write-Host "  Suspends protection without decrypting. Use before BIOS or" -ForegroundColor $ColorSchema.Info
    Write-Host "  firmware updates to avoid unexpected recovery key prompts." -ForegroundColor $ColorSchema.Info
    Write-Host "  Protection automatically resumes after 1 reboot." -ForegroundColor $ColorSchema.Info

    $vol = Select-Drive -Prompt "Drive letter to suspend"
    if (-not $vol) { return }

    if ($WhatIf) {
        Write-Host ""
        Write-Host "  [~] Would suspend BitLocker protection on $($vol.MountPoint) (resumes after 1 reboot)" -ForegroundColor Cyan
        Write-Host ""
        return
    }

    try {
        Suspend-BitLocker -MountPoint $vol.MountPoint -RebootCount 1 -ErrorAction Stop | Out-Null
        Write-Host ""
        Write-Host "  [+] BitLocker suspended on $($vol.MountPoint) — resumes after next reboot." -ForegroundColor $ColorSchema.Success
    }
    catch {
        Write-Host "  [-] Suspend failed: $_" -ForegroundColor $ColorSchema.Error
    }

    Write-Host ""
}

function Resume-DriveProtection {
    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  RESUME BITLOCKER" -ForegroundColor $ColorSchema.Header
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header

    $vol = Select-Drive -Prompt "Drive letter to resume"
    if (-not $vol) { return }

    if ($WhatIf) {
        Write-Host ""
        Write-Host "  [~] Would resume BitLocker protection on $($vol.MountPoint)" -ForegroundColor Cyan
        Write-Host ""
        return
    }

    try {
        Resume-BitLocker -MountPoint $vol.MountPoint -ErrorAction Stop | Out-Null
        Write-Host ""
        Write-Host "  [+] BitLocker protection resumed on $($vol.MountPoint)." -ForegroundColor $ColorSchema.Success
    }
    catch {
        Write-Host "  [-] Resume failed: $_" -ForegroundColor $ColorSchema.Error
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-DriveStatus

    $mountPoint = "$($Drive.ToUpper().TrimEnd(':')):"

    switch ($Action) {
        "Status"      { Write-Host "[OK] Status displayed above." -ForegroundColor $ColorSchema.Success }
        "Enable"      {
            Write-Host "  [*] Unattended enable not supported — requires key protector selection and key confirmation." -ForegroundColor $ColorSchema.Warning
            Write-Host "  [!!] Run cipher.ps1 interactively to enable BitLocker." -ForegroundColor $ColorSchema.Warning
        }
        "Disable"     {
            try {
                Disable-BitLocker -MountPoint $mountPoint -ErrorAction Stop | Out-Null
                Write-Host "  [+] Decryption started on $mountPoint." -ForegroundColor $ColorSchema.Success
            } catch {
                Write-Host "  [-] Failed: $_" -ForegroundColor $ColorSchema.Error
            }
        }
        "Suspend"     {
            try {
                Suspend-BitLocker -MountPoint $mountPoint -RebootCount 1 -ErrorAction Stop | Out-Null
                Write-Host "  [+] BitLocker suspended on $mountPoint — resumes after next reboot." -ForegroundColor $ColorSchema.Success
            } catch {
                Write-Host "  [-] Failed: $_" -ForegroundColor $ColorSchema.Error
            }
        }
        "Resume"      {
            try {
                Resume-BitLocker -MountPoint $mountPoint -ErrorAction Stop | Out-Null
                Write-Host "  [+] BitLocker resumed on $mountPoint." -ForegroundColor $ColorSchema.Success
            } catch {
                Write-Host "  [-] Failed: $_" -ForegroundColor $ColorSchema.Error
            }
        }
        "BackupAD"    {
            try {
                $vol = Get-BitLockerVolume -MountPoint $mountPoint -ErrorAction Stop
                $kp  = $vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } | Select-Object -First 1
                if (-not $kp) { Write-Host "  [-] No recovery password found." -ForegroundColor $ColorSchema.Error; break }
                Backup-BitLockerKeyProtector -MountPoint $mountPoint -KeyProtectorId $kp.KeyProtectorId -ErrorAction Stop | Out-Null
                Write-Host "  [+] Recovery key backed up to Active Directory." -ForegroundColor $ColorSchema.Success
            } catch {
                Write-Host "  [-] Backup failed: $_" -ForegroundColor $ColorSchema.Error
            }
        }
        "BackupEntraID" {
            try {
                $vol = Get-BitLockerVolume -MountPoint $mountPoint -ErrorAction Stop
                $kp  = $vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } | Select-Object -First 1
                if (-not $kp) { Write-Host "  [-] No recovery password found." -ForegroundColor $ColorSchema.Error; break }
                BackupToAAD-BitLockerKeyProtector -MountPoint $mountPoint -KeyProtectorId $kp.KeyProtectorId -ErrorAction Stop | Out-Null
                Write-Host "  [+] Recovery key backed up to Entra ID." -ForegroundColor $ColorSchema.Success
            } catch {
                Write-Host "  [-] Backup failed: $_" -ForegroundColor $ColorSchema.Error
            }
        }
    }
} else {
    $choice = ""

    do {
        Show-CipherBanner
        if ($WhatIf) {
            Write-Host "  ─────────────────────────────────────────────────────────────" -ForegroundColor Cyan
            Write-Host "  [~] DRY RUN MODE — No changes will be made to this system." -ForegroundColor Cyan
            Write-Host "  ─────────────────────────────────────────────────────────────" -ForegroundColor Cyan
            Write-Host ""
        }
        Show-DriveStatus

        Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
        Write-Host "  ACTIONS" -ForegroundColor $ColorSchema.Header
        Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
        Write-Host ""
        Write-Host "  [1] Enable BitLocker on a drive" -ForegroundColor $ColorSchema.Info
        Write-Host "  [2] Disable BitLocker on a drive" -ForegroundColor $ColorSchema.Info
        Write-Host "  [3] Backup recovery key  (AD / Entra ID)" -ForegroundColor $ColorSchema.Info
        Write-Host "  [4] Show recovery key" -ForegroundColor $ColorSchema.Info
        Write-Host "  [5] Suspend BitLocker protection" -ForegroundColor $ColorSchema.Info
        Write-Host "  [6] Resume BitLocker protection" -ForegroundColor $ColorSchema.Info
        Write-Host "  [R] Refresh drive status" -ForegroundColor $ColorSchema.Info
        Write-Host "  [Q] Quit" -ForegroundColor $ColorSchema.Info
        Write-Host ""
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $ColorSchema.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            "1" { Enable-DriveEncryption }
            "2" { Disable-DriveEncryption }
            "3" { Backup-RecoveryKey }
            "4" { Show-RecoveryKey }
            "5" { Suspend-DriveProtection }
            "6" { Resume-DriveProtection }
            "R" { }
            "Q" {
                Write-Host ""
                Write-Host "  Closing C.I.P.H.E.R." -ForegroundColor $ColorSchema.Header
                Write-Host ""
            }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1-6, R, or Q." -ForegroundColor $ColorSchema.Warning
                Start-Sleep -Seconds 1
            }
        }

        if ($choice -notin @("Q", "R")) {
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $ColorSchema.Info
            Read-Host | Out-Null
        }

    } while ($choice -ne "Q")
}
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
