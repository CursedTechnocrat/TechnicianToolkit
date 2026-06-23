<#
.SYNOPSIS
    C.I.P.H.E.R. — Configures & Implements Policy-based Hardware Encryption & Recovery
    BitLocker Drive Encryption Tool for PowerShell 5.1+

.DESCRIPTION
    Manages BitLocker drive encryption across all volumes on the local machine.
    Check encryption status, enable or disable encryption, back up recovery keys
    to Active Directory or Entra ID, view recovery key IDs, suspend or resume
    BitLocker protection, and export a status + recovery-key report to PDF.

.USAGE
    PS C:\> .\cipher.ps1                                           # Must be run as Administrator
    PS C:\> .\cipher.ps1 -WhatIf                                   # Preview actions without making changes
    PS C:\> .\cipher.ps1 -Unattended -Action Status                # Show drive status and exit
    PS C:\> .\cipher.ps1 -Unattended -Action Disable -Drive C      # Disable BitLocker on C:
    PS C:\> .\cipher.ps1 -Unattended -Action Suspend -Drive C      # Suspend BitLocker on C:
    PS C:\> .\cipher.ps1 -Unattended -Action BackupAD -Drive C     # Backup recovery key to AD
    PS C:\> .\cipher.ps1 -Unattended -Action Export                # Export status + recovery keys to PDF
    PS C:\> .\cipher.ps1 -Unattended -Action Export -OutputPath D:\Reports

.NOTES
    Version : 4.0

#>

param(
    [switch]$Unattended,
    [switch]$WhatIf,
    [ValidateSet('Status','Enable','Disable','Suspend','Resume','BackupAD','BackupEntraID','Export')]
    [string]$Action = "Status",
    [ValidatePattern('^[A-Za-z]:?$')]
    [string]$Drive  = "C",
    [string]$OutputPath,
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

# Turns BitLocker protection ON for a volume that is already encrypted but has
# Turns BitLocker protection ON for a volume that is already encrypted but has
# ProtectionStatus = Off — typically an OEM "device encryption waiting" state
# where the volume carries a TPM protector plus an unsecured clear key, so the
# volume key is exposed. Resume-BitLocker only handles volumes suspended via
# Suspend-BitLocker; on these it throws FVE_E_KEY_REQUIRED (0x8031001D, "you
# cannot delete the last key"), so we fall back to manage-bde -protectors -enable,
# which removes the unsecured key and enforces the remaining protectors.
#
# Verification is deliberate: Get-BitLockerVolume can briefly report a stale
# ProtectionStatus right after an external manage-bde process flips it, so we
# treat manage-bde's exit code as authoritative and also poll the status. Returns
# $true if protection ends up On. On genuine failure the manage-bde output and
# exit code are left in $script:LastEnableOutput / $script:LastEnableExit so the
# caller can surface them.
function Enable-DriveProtection {
    param([Parameter(Mandatory)][string]$MountPoint)

    $script:LastEnableOutput = ''
    $script:LastEnableExit   = $null

    # Resume-BitLocker cleanly resumes a volume suspended via Suspend-BitLocker.
    try {
        Resume-BitLocker -MountPoint $MountPoint -ErrorAction Stop | Out-Null
        if (Test-ProtectionOn -MountPoint $MountPoint) { return $true }
    } catch { }

    # Otherwise enforce protectors / remove the unsecured clear key via manage-bde.
    $script:LastEnableOutput = (& manage-bde.exe -protectors -enable $MountPoint 2>&1 | Out-String).Trim()
    $script:LastEnableExit   = $LASTEXITCODE

    if (Test-ProtectionOn -MountPoint $MountPoint) { return $true }
    return ($script:LastEnableExit -eq 0)
}

# Returns $true if the volume's ProtectionStatus is On. ProtectionStatus can lag
# the actual FVE change by a moment (especially after a separate manage-bde
# process), so poll briefly rather than reading it once.
function Test-ProtectionOn {
    param([Parameter(Mandatory)][string]$MountPoint)
    for ($i = 0; $i -lt 6; $i++) {
        $v = Get-BitLockerVolume -MountPoint $MountPoint -ErrorAction SilentlyContinue
        if ($v -and $v.ProtectionStatus -eq 'On') { return $true }
        Start-Sleep -Milliseconds 500
    }
    return $false
}

# Starts encryption on a not-yet-encrypted volume, working around two snags seen
# in the field:
#   - Virtual machines, where the pre-encryption hardware test and used-space-only
#     conversion are unreliable — encryption is started full-volume with the
#     hardware test skipped.
#   - Physical disks that reject used-space-only conversion with 0x803100a5 — we
#     retry full-volume.
# manage-bde is used so the volume's existing protectors are honoured and
# -SkipHardwareTest is available (it begins encrypting immediately instead of
# waiting for a reboot-time hardware test). Returns $true on success; the command
# output / exit code are left in $script:LastEncryptOutput / $script:LastEncryptExit.
function Start-DriveEncryption {
    param(
        [Parameter(Mandatory)][string]$MountPoint,
        [string]$EncryptionMethod = 'XtsAes256'
    )

    $model = (Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).Model
    $isVM  = $model -match 'Virtual|VMware|QEMU|KVM|Hyper-V|Xen'

    if ($isVM) {
        $out  = & manage-bde.exe -on $MountPoint -EncryptionMethod $EncryptionMethod -SkipHardwareTest 2>&1 | Out-String
        $exit = $LASTEXITCODE
    } else {
        $out  = & manage-bde.exe -on $MountPoint -EncryptionMethod $EncryptionMethod -UsedSpaceOnly -SkipHardwareTest 2>&1 | Out-String
        $exit = $LASTEXITCODE
        if ($exit -ne 0 -and $out -match '0x803100a5') {
            # Used-space-only not accepted on this volume — retry full-volume.
            $out  = & manage-bde.exe -on $MountPoint -EncryptionMethod $EncryptionMethod -SkipHardwareTest 2>&1 | Out-String
            $exit = $LASTEXITCODE
        }
    }

    $script:LastEncryptOutput = $out.Trim()
    $script:LastEncryptExit   = $exit
    return ($exit -eq 0)
}

# Renders an HTML file to PDF using headless Microsoft Edge (or Chrome as a
# fallback) — both ship Chromium's --print-to-pdf, so no third-party tooling is
# required. Returns $true if the PDF was produced, $false if no browser is
# available (the caller keeps the HTML in that case).
function Convert-HtmlToPdf {
    param(
        [Parameter(Mandatory)][string]$HtmlPath,
        [Parameter(Mandatory)][string]$PdfPath
    )

    $bases = @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { $_ }
    $candidates = foreach ($b in $bases) {
        Join-Path $b 'Microsoft\Edge\Application\msedge.exe'
        Join-Path $b 'Google\Chrome\Application\chrome.exe'
    }
    $browser = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $browser) { return $false }

    $uri = 'file:///' + ($HtmlPath -replace '\\', '/')
    $browserArgs = @(
        '--headless'
        '--disable-gpu'
        '--no-pdf-header-footer'
        "--print-to-pdf=`"$PdfPath`""
        "`"$uri`""
    )
    try {
        Start-Process -FilePath $browser -ArgumentList $browserArgs -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop | Out-Null
    } catch {
        return $false
    }
    # The PDF is flushed to disk a moment after the process exits — give it time.
    for ($i = 0; $i -lt 10 -and -not (Test-Path $PdfPath); $i++) { Start-Sleep -Milliseconds 300 }
    return (Test-Path $PdfPath)
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

        if ($vol.ProtectionStatus -eq "On") {
            $existingKey = $vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } | Select-Object -Last 1
            Write-Host ""
            Write-Host "  [+] BitLocker protection is already ON — nothing to do." -ForegroundColor $ColorSchema.Success
            if ($existingKey) {
                Write-Host "  ID  : $($existingKey.KeyProtectorId)" -ForegroundColor $ColorSchema.Warning
                Write-Host "  Key : $($existingKey.RecoveryPassword)" -ForegroundColor $ColorSchema.Warning
            }
            Write-Host ""
            return
        }

        # Protection is OFF on an already-encrypted volume. Two very different cases:
        #   (a) the volume is suspended / already carries a usable protector (TPM,
        #       recovery password, etc.) — it just needs RESUMING, and adding another
        #       recovery password would only pile up duplicates;
        #   (b) the only thing holding the volume key is an unsecured clear key — a
        #       usable protector must be added before protection can be enabled.
        $usableTypes  = @('Tpm','TpmPin','TpmStartupKey','TpmPinStartupKey','RecoveryPassword','Password')
        $hasProtector = [bool]($vol.KeyProtector | Where-Object { $_.KeyProtectorType -in $usableTypes })

        if ($WhatIf) {
            Write-Host ""
            if ($hasProtector) {
                Write-Host "  [~] Would resume BitLocker protection on $($vol.MountPoint) (suspended; protectors already present)." -ForegroundColor Cyan
            } else {
                Write-Host "  [~] Would add a recovery password to $($vol.MountPoint), then enable protection." -ForegroundColor Cyan
            }
            Write-Host ""
            return
        }

        if ($hasProtector) {
            Write-Host ""
            Write-Host "  [!!] BitLocker is suspended (protection off) but usable key protectors exist." -ForegroundColor $ColorSchema.Warning
            Write-Host "  [*] Resuming BitLocker protection..." -ForegroundColor $ColorSchema.Progress
        } else {
            Write-Host ""
            Write-Host "  [!!] No usable key protector found — only an unsecured (clear) key is present." -ForegroundColor $ColorSchema.Warning
            Write-Host "  [*] Adding a recovery password before enabling protection..." -ForegroundColor $ColorSchema.Progress
            try {
                $null = Add-BitLockerKeyProtector -MountPoint $vol.MountPoint -RecoveryPasswordProtector -ErrorAction Stop
                # Re-query to confirm the protector actually persisted to the volume —
                # the returned object can report success even when a policy blocks the commit.
                $vol    = Get-BitLockerVolume -MountPoint $vol.MountPoint -ErrorAction Stop
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
                } else {
                    Write-Host "  [-] The recovery password did not persist to the volume — cannot enable protection." -ForegroundColor $ColorSchema.Error
                    Write-Host "      This usually means a policy requires the key be escrowed first, or the drive is" -ForegroundColor $ColorSchema.Warning
                    Write-Host "      in Windows Device Encryption 'waiting' state. Resolve that, then retry." -ForegroundColor $ColorSchema.Warning
                    Write-TKError -ScriptName 'cipher' -Message "Recovery password did not persist on '$($vol.MountPoint)' after Add-BitLockerKeyProtector." -Category 'BitLocker Enable'
                    Write-Host ""
                    return
                }
            } catch {
                Write-Host "  [-] Failed to add recovery key: $_" -ForegroundColor $ColorSchema.Error
                Write-Host ""
                return
            }
        }

        if (Enable-DriveProtection -MountPoint $vol.MountPoint) {
            Write-Host "  [+] BitLocker protection is now ON." -ForegroundColor $ColorSchema.Success
        } else {
            Write-Host "  [-] Failed to activate protection on $($vol.MountPoint)." -ForegroundColor $ColorSchema.Error
            if ($script:LastEnableOutput) {
                Write-Host "      manage-bde (exit $($script:LastEnableExit)): $($script:LastEnableOutput)" -ForegroundColor $ColorSchema.Warning
            }
            Write-Host "      Run manually to inspect: manage-bde -protectors -enable $($vol.MountPoint)" -ForegroundColor $ColorSchema.Warning
            Write-TKError -ScriptName 'cipher' -Message "Activate protection failed on '$($vol.MountPoint)' (manage-bde exit $($script:LastEnableExit)): $($script:LastEnableOutput)" -Category 'BitLocker Enable'
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
        Write-Host "  [~] Would start XtsAes256 encryption on $($vol.MountPoint) (used space only; full-volume on VMs or if the disk rejects it)" -ForegroundColor Cyan
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
        if (-not (Start-DriveEncryption -MountPoint $vol.MountPoint -EncryptionMethod 'XtsAes256')) {
            Write-Host "  [-] Failed to start encryption on $($vol.MountPoint)." -ForegroundColor $ColorSchema.Error
            if ($script:LastEncryptOutput) {
                Write-Host "      manage-bde (exit $($script:LastEncryptExit)): $($script:LastEncryptOutput)" -ForegroundColor $ColorSchema.Warning
            }
            Write-TKError -ScriptName 'cipher' -Message "Start encryption failed on '$($vol.MountPoint)' (manage-bde exit $($script:LastEncryptExit)): $($script:LastEncryptOutput)" -Category 'BitLocker Enable'
            Write-Host ""
            return
        }

        $volCheck = Get-BitLockerVolume -MountPoint $vol.MountPoint -ErrorAction SilentlyContinue
        if ($volCheck -and $volCheck.ProtectionStatus -ne "On") {
            Write-Host "  [*] Activating BitLocker protection..." -ForegroundColor $ColorSchema.Progress
            [void](Enable-DriveProtection -MountPoint $vol.MountPoint)
        }

        Write-Host "  [+] Encryption started on $($vol.MountPoint). BitLocker protection is ON." -ForegroundColor $ColorSchema.Success
    }
    catch {
        Write-Host "  [-] Failed to enable BitLocker: $_" -ForegroundColor $ColorSchema.Error
        Write-TKError -ScriptName 'cipher' -Message "Enable-BitLocker failed on '$($vol.MountPoint)': $($_.Exception.Message)" -Category 'BitLocker Enable'
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
        Write-TKError -ScriptName 'cipher' -Message "Disable-BitLocker failed on '$($vol.MountPoint)': $($_.Exception.Message)" -Category 'BitLocker Disable'
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
        Write-TKError -ScriptName 'cipher' -Message "Recovery key backup failed on '$($vol.MountPoint)': $($_.Exception.Message)" -Category 'BitLocker KeyBackup'
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

function Export-EncryptionReport {
    param([switch]$SkipConfirm)

    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host "  EXPORT ENCRYPTION REPORT (PDF)" -ForegroundColor $ColorSchema.Header
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
    Write-Host ""
    Write-Host "  [!!] This report includes the 48-digit RECOVERY PASSWORDS for every" -ForegroundColor $ColorSchema.Warning
    Write-Host "       drive that has one. Anyone who opens the PDF can unlock those" -ForegroundColor $ColorSchema.Warning
    Write-Host "       drives — store it somewhere secure (not on the encrypted drive)." -ForegroundColor $ColorSchema.Warning

    if ($WhatIf) {
        Write-Host ""
        Write-Host "  [~] Would build an HTML + PDF encryption report (including recovery keys)." -ForegroundColor Cyan
        Write-Host ""
        return
    }

    if (-not $SkipConfirm) {
        Write-Host ""
        Write-Host -NoNewline "  Generate report with recovery keys? (Y/N): " -ForegroundColor $ColorSchema.Warning
        if ((Read-Host).Trim().ToUpper() -ne "Y") {
            Write-Host "  [*] Export cancelled." -ForegroundColor $ColorSchema.Info
            Write-Host ""
            return
        }
    }

    try {
        $volumes = @(Get-BitLockerVolume -ErrorAction Stop)
    } catch {
        Write-Host ""
        Write-Host "  [-] Unable to retrieve BitLocker information: $_" -ForegroundColor $ColorSchema.Error
        Write-TKError -ScriptName 'cipher' -Message "Export report: Get-BitLockerVolume failed: $($_.Exception.Message)" -Category 'BitLocker Export'
        Write-Host ""
        return
    }

    $reportDir = if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath }
                 else { Resolve-LogDirectory -FallbackPath $PSScriptRoot }
    if (-not (Test-Path $reportDir)) { New-Item -ItemType Directory -Path $reportDir -Force | Out-Null }

    $stamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
    $htmlPath = Join-Path $reportDir "CIPHER_Report_$stamp.html"
    $pdfPath  = Join-Path $reportDir "CIPHER_Report_$stamp.pdf"

    Write-Host ""
    Write-Host "  [*] Building report..." -ForegroundColor $ColorSchema.Progress

    $cfg       = Get-TKConfig
    $orgPrefix = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$(EscHtml $cfg.OrgName) — " } else { '' }

    $html = Get-TKHtmlHead -Title 'BitLocker Encryption Report' `
        -ScriptName 'C.I.P.H.E.R.' `
        -Subtitle "$orgPrefix$env:COMPUTERNAME" `
        -MetaItems ([ordered]@{
            'Generated' = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            'Run As'    = "$env:USERDOMAIN\$env:USERNAME"
            'Volumes'   = $volumes.Count
        }) `
        -NavItems @('Drive Status', 'Recovery Keys')

    $html += @"
<div class="tk-section" id="s01">
  <div class="tk-section-title"><span class="tk-section-num">01</span> Drive Status</div>
  <div class="tk-info-box"><span class="tk-info-label">SENSITIVE</span> This document contains BitLocker recovery passwords. Anyone who reads it can unlock the listed drives — handle it as a secret.</div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>Drive</th><th>Volume Status</th><th>Protection</th><th>Encryption</th><th>Key Protectors</th></tr></thead>
      <tbody>
"@
    foreach ($vol in $volumes) {
        $protBadge = if ($vol.ProtectionStatus -eq 'On') { "<span class='tk-badge-ok'>On</span>" } else { "<span class='tk-badge-warn'>Off</span>" }
        $keyTypes  = if ($vol.KeyProtector.Count -gt 0) { ($vol.KeyProtector | ForEach-Object { $_.KeyProtectorType }) -join ', ' } else { 'None' }
        $html += "<tr><td class='tk-mono'>$(EscHtml $vol.MountPoint)</td><td>$(EscHtml $vol.VolumeStatus)</td><td>$protBadge</td><td>$(EscHtml ([string]$vol.EncryptionPercentage))%</td><td>$(EscHtml $keyTypes)</td></tr>"
    }
    $html += @"
      </tbody>
    </table>
  </div>
</div>
<div class="tk-section" id="s02">
  <div class="tk-section-title"><span class="tk-section-num">02</span> Recovery Keys</div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>Drive</th><th>Key Protector ID</th><th>Recovery Password</th></tr></thead>
      <tbody>
"@
    $anyKeys = $false
    foreach ($vol in $volumes) {
        foreach ($key in ($vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' })) {
            $anyKeys = $true
            $html += "<tr><td class='tk-mono'>$(EscHtml $vol.MountPoint)</td><td class='tk-mono'>$(EscHtml $key.KeyProtectorId)</td><td class='tk-mono'>$(EscHtml $key.RecoveryPassword)</td></tr>"
        }
    }
    if (-not $anyKeys) {
        $html += "<tr><td colspan='3'>No recovery password protectors found on any volume.</td></tr>"
    }
    $html += @"
      </tbody>
    </table>
  </div>
</div>
"@
    $html += Get-TKHtmlFoot -ScriptName 'C.I.P.H.E.R. v4.0'

    try {
        $html | Out-File -FilePath $htmlPath -Encoding UTF8 -ErrorAction Stop
    } catch {
        Write-Host "  [-] Failed to write report: $_" -ForegroundColor $ColorSchema.Error
        Write-TKError -ScriptName 'cipher' -Message "Export report: writing HTML failed: $($_.Exception.Message)" -Category 'BitLocker Export'
        Write-Host ""
        return
    }

    Write-Host "  [*] Converting to PDF..." -ForegroundColor $ColorSchema.Progress
    if (Convert-HtmlToPdf -HtmlPath $htmlPath -PdfPath $pdfPath) {
        Remove-Item -Path $htmlPath -Force -ErrorAction SilentlyContinue
        Write-Host "  [+] PDF report saved: $pdfPath" -ForegroundColor $ColorSchema.Success
    } else {
        Write-Host "  [!!] Microsoft Edge / Chrome not found — could not render PDF." -ForegroundColor $ColorSchema.Warning
        Write-Host "  [+] HTML report saved: $htmlPath" -ForegroundColor $ColorSchema.Success
        Write-Host "      Open it and choose Print -> Save as PDF to produce a PDF." -ForegroundColor $ColorSchema.Info
    }

    Write-Host ""
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
        Write-TKError -ScriptName 'cipher' -Message "Resume-BitLocker failed on '$($vol.MountPoint)': $($_.Exception.Message)" -Category 'BitLocker Resume'
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
        "Export"      { Export-EncryptionReport -SkipConfirm }
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
        Write-Host "  [7] Export encryption report  (PDF)" -ForegroundColor $ColorSchema.Info
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
            "7" { Export-EncryptionReport }
            "R" { }
            "Q" {
                Write-Host ""
                Write-Host "  Closing C.I.P.H.E.R." -ForegroundColor $ColorSchema.Header
                Write-Host ""
            }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1-7, R, or Q." -ForegroundColor $ColorSchema.Warning
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
