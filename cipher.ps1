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
    PS C:\> .\cipher.ps1      # Must be run as Administrator

.NOTES
    Version : 1.0

    Tools Available
    ─────────────────────────────────────────────────────────────────
    G.R.I.M.O.I.R.E.       — Technician Toolkit hub and central launcher
    R.U.N.E.P.R.E.S.S.     — Printer driver installation & configuration
    R.E.S.T.O.R.A.T.I.O.N. — Windows Update management
    C.O.N.J.U.R.E.         — Software deployment via winget / Chocolatey
    O.R.A.C.L.E.           — System diagnostics & HTML report generation
    C.O.V.E.N.A.N.T.       — Machine onboarding & Entra ID domain join
    P.H.A.N.T.O.M.         — Profile migration & data transfer
    C.I.P.H.E.R.           — BitLocker drive encryption management
    W.A.R.D.               — User account & local security audit
    A.R.C.H.I.V.E.         — Pre-reimaging profile backup

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success messages
    Yellow   Warnings and cautions
    Red      Critical errors
    Gray     Information and details
#>

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script must be run as Administrator!" -ForegroundColor Red
    exit 1
}

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

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
    Clear-Host
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
    $input = (Read-Host).Trim().ToUpper().TrimEnd(':')
    $mountPoint = "$input`:"

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
        Write-Host "  [+] Encryption started on $($vol.MountPoint). This runs in the background." -ForegroundColor $ColorSchema.Success
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
# MAIN MENU LOOP
# ─────────────────────────────────────────────────────────────────────────────

$choice = ""

do {
    Show-CipherBanner
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
