# cipher.ps1 - C.I.P.H.E.R. — Configures & Implements Policy-based Hardware Encryption & Recovery
# Part of the Technician Toolkit - https://github.com/CursedTechnocrat/TechnicianToolkit
#
# Copyright (C) 2026 John Joseph Bejarana (CursedTechnocrat) and the Technician Toolkit contributors
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
# SPDX-License-Identifier: GPL-3.0-or-later

<#
.SYNOPSIS
    C.I.P.H.E.R. — Configures & Implements Policy-based Hardware Encryption & Recovery
    BitLocker Drive Encryption Tool for PowerShell 5.1+

.DESCRIPTION
    Manages BitLocker on the local machine: show drive status, enable or
    disable encryption, suspend or resume protection, show recovery keys, back
    them up to Active Directory or Entra ID, and export a status + recovery-key
    report.

    Every change goes through manage-bde.exe and every read through the
    Win32_EncryptableVolume CIM class. Neither depends on the BitLocker
    PowerShell module, so the tool behaves the same under Windows PowerShell
    5.1 and PowerShell 7.

    Enable uses one protector set: TPM + recovery password on the operating
    system drive, recovery password + auto-unlock on any other drive.

.USAGE
    PS C:\> .\cipher.ps1                                           # Must be run as Administrator
    PS C:\> .\cipher.ps1 -WhatIf                                   # Preview actions without making changes
    PS C:\> .\cipher.ps1 -Unattended -Action Status                # Show drive status and exit
    PS C:\> .\cipher.ps1 -Unattended -Action Enable -Drive C       # Encrypt C: (TPM + recovery password)
    PS C:\> .\cipher.ps1 -Unattended -Action Disable -Drive C      # Decrypt C:
    PS C:\> .\cipher.ps1 -Unattended -Action Suspend -Drive C      # Suspend BitLocker on C: for one reboot
    PS C:\> .\cipher.ps1 -Unattended -Action BackupAD -Drive C     # Backup recovery key to AD
    PS C:\> .\cipher.ps1 -Unattended -Action Export                # Export status + recovery keys to HTML
    PS C:\> .\cipher.ps1 -Unattended -Action Export -OutputPath D:\Reports

.NOTES
    Version : 5.1

    Credits : Thanks to Steve the Killer for help and letting me use his
              script BERET: https://tools.thekiller.net/killer-scripts

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
# BITLOCKER PLUMBING
# ─────────────────────────────────────────────────────────────────────────────

# Sysnative reaches the real System32 from a 32-bit host, where manage-bde
# would otherwise not be found at all.
$ManageBde = Join-Path $env:windir 'Sysnative\manage-bde.exe'
if (-not (Test-Path $ManageBde)) { $ManageBde = Join-Path $env:windir 'System32\manage-bde.exe' }

$BdeNamespace = 'root/CIMV2/Security/MicrosoftVolumeEncryption'

$ConversionStatusNames = @{
    0 = 'FullyDecrypted'; 1 = 'FullyEncrypted'; 2 = 'EncryptionInProgress'
    3 = 'DecryptionInProgress'; 4 = 'EncryptionPaused'; 5 = 'DecryptionPaused'
}
$ProtectorTypeNames = @{
    0 = 'Unknown'; 1 = 'Tpm'; 2 = 'ExternalKey'; 3 = 'RecoveryPassword'; 4 = 'TpmPin'
    5 = 'TpmStartupKey'; 6 = 'TpmPinStartupKey'; 7 = 'PublicKey'; 8 = 'Password'
    9 = 'TpmNetworkKey'; 10 = 'AdAccountOrGroup'
}

# Runs manage-bde, echoes its output, and returns $true on exit code 0. Under
# -WhatIf it only prints the command. Every change the tool makes goes through
# here.
function Invoke-ManageBde {
    param([Parameter(Mandatory)][string[]]$Arguments)

    if ($WhatIf) {
        Write-Host "  [~] Would run: manage-bde $($Arguments -join ' ')" -ForegroundColor Cyan
        return $true
    }

    $output = & $ManageBde @Arguments 2>&1
    $exit   = $LASTEXITCODE
    foreach ($line in $output) {
        if ("$line".Trim()) { Write-Host "      $line" -ForegroundColor $ColorSchema.Info }
    }
    if ($exit -ne 0) {
        Write-TKError -ScriptName 'cipher' -Message "manage-bde $($Arguments -join ' ') exited $exit" -Category 'BitLocker'
    }
    return ($exit -eq 0)
}

# Returns one object per lettered volume: MountPoint, Status, Protection,
# Percent and Protectors (Id / Type / RecoveryPassword).
function Get-CipherVolume {
    param([string]$MountPoint)

    $volumes = Get-CimInstance -Namespace $BdeNamespace -ClassName Win32_EncryptableVolume -ErrorAction Stop |
        Where-Object { $_.DriveLetter -and (-not $MountPoint -or $_.DriveLetter -eq $MountPoint) }

    foreach ($v in $volumes) {
        $conv = Invoke-CimMethod -InputObject $v -MethodName GetConversionStatus
        $ids  = (Invoke-CimMethod -InputObject $v -MethodName GetKeyProtectors).VolumeKeyProtectorID

        $protectors = foreach ($id in $ids) {
            $idArg    = @{ VolumeKeyProtectorID = $id }
            $typeCode = [int](Invoke-CimMethod -InputObject $v -MethodName GetKeyProtectorType -Arguments $idArg).KeyProtectorType
            $password = $null
            if ($typeCode -eq 3) {
                $password = (Invoke-CimMethod -InputObject $v -MethodName GetKeyProtectorNumericalPassword -Arguments $idArg).NumericalPassword
            }
            [PSCustomObject]@{ Id = $id; Type = $ProtectorTypeNames[$typeCode]; RecoveryPassword = $password }
        }

        # A locked volume fails GetConversionStatus; its zeroed output must not
        # read as FullyDecrypted.
        $status = if ($conv.ReturnValue -eq 0) { $ConversionStatusNames[[int]$conv.ConversionStatus] }
        [PSCustomObject]@{
            MountPoint = $v.DriveLetter
            Status     = if ($status) { $status } else { 'Unknown' }
            Protection = switch ([int]$v.ProtectionStatus) { 0 { 'Off' } 1 { 'On' } default { 'Unknown' } }
            Percent    = [int]$conv.EncryptionPercentage
            Protectors = @($protectors)
        }
    }
}

# Returns the volume for one drive, or $null after saying why.
function Get-TargetVolume {
    param([Parameter(Mandatory)][string]$MountPoint)
    $vol = Get-CipherVolume -MountPoint $MountPoint | Select-Object -First 1
    if (-not $vol) { Write-Fail "Drive $MountPoint was not found, or BitLocker cannot manage it." }
    return $vol
}

# ─────────────────────────────────────────────────────────────────────────────
# ACTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Show-DriveStatus {
    Write-Section 'DRIVE ENCRYPTION STATUS'
    foreach ($vol in Get-CipherVolume) {
        $statusColor = switch ($vol.Status) {
            'FullyEncrypted' { $ColorSchema.Success }
            'FullyDecrypted' { $ColorSchema.Warning }
            default          { $ColorSchema.Progress }
        }
        $protColor = if ($vol.Protection -eq 'On') { $ColorSchema.Success } else { $ColorSchema.Warning }
        $types     = if ($vol.Protectors.Count) { ($vol.Protectors.Type) -join ', ' } else { 'None' }

        Write-Host "  Drive $($vol.MountPoint)" -ForegroundColor $ColorSchema.Header
        Write-Host ("    Status      : {0}" -f $vol.Status) -ForegroundColor $statusColor
        Write-Host ("    Protection  : {0}" -f $vol.Protection) -ForegroundColor $protColor
        Write-Host ("    Encrypted   : {0}%" -f $vol.Percent) -ForegroundColor $ColorSchema.Info
        Write-Host ("    Protectors  : {0}" -f $types) -ForegroundColor $ColorSchema.Info
        Write-Host ""
    }
}

function Show-RecoveryKey {
    param([Parameter(Mandatory)][string]$MountPoint)

    $vol = Get-TargetVolume -MountPoint $MountPoint
    if (-not $vol) { return }

    $keys = @($vol.Protectors | Where-Object { $_.Type -eq 'RecoveryPassword' })
    if (-not $keys) {
        Write-Warn "No recovery password on $MountPoint."
        return
    }
    Write-Host ""
    Write-Host "  RECOVERY KEY(S) FOR $MountPoint" -ForegroundColor $ColorSchema.Warning
    foreach ($key in $keys) {
        Write-Host "    ID  : $($key.Id)" -ForegroundColor $ColorSchema.Warning
        Write-Host "    Key : $($key.RecoveryPassword)" -ForegroundColor $ColorSchema.Warning
    }
    Write-Host ""
}

# Encrypts a decrypted drive, or turns protection back on for an encrypted one.
# Either way the drive ends up with a recovery password plus TPM (OS drive) or
# auto-unlock (any other drive).
function Enable-DriveEncryption {
    param([Parameter(Mandatory)][string]$MountPoint)

    Write-Section "ENABLE BITLOCKER ON $MountPoint"
    $vol = Get-TargetVolume -MountPoint $MountPoint
    if (-not $vol) { return }

    if ($vol.Status -eq 'Unknown') {
        Write-Fail "$MountPoint is locked or unreadable. Unlock it first (manage-bde -unlock)."
        return
    }
    if ($vol.Status -in 'DecryptionInProgress', 'DecryptionPaused') {
        Write-Warn "$MountPoint is still decrypting ($($vol.Percent)% encrypted). Run Enable again once it reads FullyDecrypted."
        return
    }
    if ($vol.Status -ne 'FullyDecrypted' -and $vol.Protection -eq 'On') {
        Write-Ok "BitLocker is already on for $MountPoint - nothing to do."
        Show-RecoveryKey -MountPoint $MountPoint
        return
    }

    $isOsDrive = $MountPoint -eq $env:SystemDrive
    $types     = @($vol.Protectors.Type)

    if ($types -notcontains 'RecoveryPassword') {
        Write-Step 'Adding a recovery password...'
        if (-not (Invoke-ManageBde '-protectors', '-add', $MountPoint, '-RecoveryPassword')) {
            Write-Fail 'Could not add a recovery password. Nothing else was changed.'
            return
        }
    }
    if ($isOsDrive -and -not ($types | Where-Object { $_ -like 'Tpm*' })) {
        Write-Step 'Adding a TPM protector...'
        if (-not (Invoke-ManageBde '-protectors', '-add', $MountPoint, '-TPM')) {
            Write-Fail 'Could not add a TPM protector. Check that the TPM is present and ready (tpm.msc).'
            return
        }
    }

    if ($vol.Status -eq 'FullyDecrypted') {
        Write-Step "Starting XTS-AES 256 encryption on $MountPoint..."
        $started = Invoke-ManageBde '-on', $MountPoint, '-EncryptionMethod', 'XtsAes256', '-UsedSpaceOnly', '-SkipHardwareTest'
        if (-not $started) {
            Write-Step 'Retrying as full-volume encryption...'
            $started = Invoke-ManageBde '-on', $MountPoint, '-EncryptionMethod', 'XtsAes256', '-SkipHardwareTest'
        }
        if (-not $started) {
            Write-Fail "Encryption did not start on $MountPoint."
            return
        }
    } else {
        # Encrypted but protection is off: suspended, or an OEM device-encryption
        # volume still holding a clear key. -enable removes the clear key. If it
        # refuses, the protectors are half-suspended; a clean -disable first puts
        # them back into a state -enable accepts. -disable never decrypts.
        Write-Step "Turning protection on for $MountPoint..."
        $on = Invoke-ManageBde '-protectors', '-enable', $MountPoint
        if (-not $on) {
            Write-Step 'Re-suspending cleanly and retrying...'
            $on = (Invoke-ManageBde '-protectors', '-disable', $MountPoint) -and
                  (Invoke-ManageBde '-protectors', '-enable', $MountPoint)
        }
        if (-not $on) {
            Write-Fail "Protection is still off on $MountPoint."
            return
        }
    }

    if (-not $isOsDrive) {
        Write-Step "Enabling auto-unlock so $MountPoint opens at sign-in..."
        if (-not (Invoke-ManageBde '-autounlock', '-enable', $MountPoint)) {
            Write-Warn 'Auto-unlock was not enabled (the OS drive must be encrypted first). The drive will ask for its recovery password after a reboot.'
        }
    }

    Write-Ok "BitLocker is enabled on $MountPoint."
    if (-not $WhatIf) {
        Write-Warn 'Save the recovery key below before closing this window.'
        Show-RecoveryKey -MountPoint $MountPoint
    }
}

function Disable-DriveEncryption {
    param([Parameter(Mandatory)][string]$MountPoint)
    Write-Section "DISABLE BITLOCKER ON $MountPoint"
    if (Invoke-ManageBde '-off', $MountPoint) {
        Write-Ok "Decryption started on $MountPoint. It runs in the background."
    }
}

function Suspend-DriveProtection {
    param([Parameter(Mandatory)][string]$MountPoint)
    Write-Section "SUSPEND BITLOCKER ON $MountPoint"
    Write-Info 'Use before BIOS / firmware updates. Protection resumes after one reboot.'
    if (Invoke-ManageBde '-protectors', '-disable', $MountPoint, '-RebootCount', '1') {
        Write-Ok "BitLocker suspended on $MountPoint until the next reboot."
    }
}

function Resume-DriveProtection {
    param([Parameter(Mandatory)][string]$MountPoint)
    Write-Section "RESUME BITLOCKER ON $MountPoint"
    if (Invoke-ManageBde '-protectors', '-enable', $MountPoint) {
        Write-Ok "BitLocker protection resumed on $MountPoint."
    }
}

function Backup-RecoveryKey {
    param(
        [Parameter(Mandatory)][string]$MountPoint,
        [Parameter(Mandatory)][ValidateSet('AD','EntraID')][string]$Target
    )
    Write-Section "BACK UP RECOVERY KEY FOR $MountPoint TO $Target"
    $vol = Get-TargetVolume -MountPoint $MountPoint
    if (-not $vol) { return }

    $keys = @($vol.Protectors | Where-Object { $_.Type -eq 'RecoveryPassword' })
    if (-not $keys) {
        Write-Fail "No recovery password on $MountPoint. Enable BitLocker first."
        return
    }
    $verb = if ($Target -eq 'AD') { '-adbackup' } else { '-aadbackup' }
    foreach ($key in $keys) {
        if (Invoke-ManageBde '-protectors', $verb, $MountPoint, '-id', $key.Id) {
            Write-Ok "Recovery key $($key.Id) backed up to $Target."
        } else {
            Write-Warn "Backup failed. Check the machine is joined to $Target and can reach it."
        }
    }
}

function Export-EncryptionReport {
    Write-Section 'EXPORT ENCRYPTION REPORT'
    Write-Warn 'The report contains the recovery passwords. Store it somewhere secure, not on the encrypted drive.'

    if ($WhatIf) {
        Write-Host "  [~] Would write an HTML report with drive status and recovery keys." -ForegroundColor Cyan
        return
    }

    $volumes = @(Get-CipherVolume)

    $reportDir = if ($OutputPath) { $OutputPath } else { Resolve-LogDirectory -FallbackPath $PSScriptRoot }
    if (-not (Test-Path $reportDir)) { New-Item -ItemType Directory -Path $reportDir -Force | Out-Null }
    $reportPath = Join-Path $reportDir "CIPHER_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

    $cfg      = Get-TKConfig
    $subtitle = if ($cfg.OrgName) { "$($cfg.OrgName) — $env:COMPUTERNAME" } else { $env:COMPUTERNAME }

    $statusRows = foreach ($vol in $volumes) {
        $badge = if ($vol.Protection -eq 'On') { "<span class='tk-badge-ok'>On</span>" } else { "<span class='tk-badge-warn'>$(EscHtml $vol.Protection)</span>" }
        $types = if ($vol.Protectors.Count) { ($vol.Protectors.Type) -join ', ' } else { 'None' }
        "<tr><td class='tk-mono'>$(EscHtml $vol.MountPoint)</td><td>$(EscHtml $vol.Status)</td><td>$badge</td><td>$($vol.Percent)%</td><td>$(EscHtml $types)</td></tr>"
    }
    $keyRows = foreach ($vol in $volumes) {
        foreach ($key in ($vol.Protectors | Where-Object { $_.Type -eq 'RecoveryPassword' })) {
            "<tr><td class='tk-mono'>$(EscHtml $vol.MountPoint)</td><td class='tk-mono'>$(EscHtml $key.Id)</td><td class='tk-mono'>$(EscHtml $key.RecoveryPassword)</td></tr>"
        }
    }
    if (-not $keyRows) { $keyRows = "<tr><td colspan='3'>No recovery passwords on any drive.</td></tr>" }

    $html  = Get-TKHtmlHead -Title 'BitLocker Encryption Report' -ScriptName 'C.I.P.H.E.R.' `
                 -Subtitle $subtitle `
                 -MetaItems ([ordered]@{
                     'Generated' = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                     'Run As'    = "$env:USERDOMAIN\$env:USERNAME"
                     'Volumes'   = $volumes.Count
                 }) `
                 -NavItems @('Drive Status', 'Recovery Keys')
    $html += @"
<div class="tk-section" id="s01">
  <div class="tk-section-title"><span class="tk-section-num">01</span> Drive Status</div>
  <div class="tk-info-box"><span class="tk-info-label">SENSITIVE</span> This document contains BitLocker recovery passwords. Anyone who reads it can unlock the listed drives.</div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>Drive</th><th>Volume Status</th><th>Protection</th><th>Encrypted</th><th>Key Protectors</th></tr></thead>
      <tbody>$($statusRows -join "`n")</tbody>
    </table>
  </div>
</div>
<div class="tk-section" id="s02">
  <div class="tk-section-title"><span class="tk-section-num">02</span> Recovery Keys</div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>Drive</th><th>Key Protector ID</th><th>Recovery Password</th></tr></thead>
      <tbody>$($keyRows -join "`n")</tbody>
    </table>
  </div>
</div>
"@
    $html += Get-TKHtmlFoot -ScriptName 'C.I.P.H.E.R. v5.1'

    $html | Out-File -FilePath $reportPath -Encoding UTF8
    Show-TKReportResult -Path $reportPath -Unattended:$Unattended
}

# Runs one action against one drive. Shared by the unattended path and the menu.
function Invoke-CipherAction {
    param([Parameter(Mandatory)][string]$Name, [string]$MountPoint)
    switch ($Name) {
        'Status'        { Show-DriveStatus }
        'Enable'        { Enable-DriveEncryption  -MountPoint $MountPoint }
        'Disable'       { Disable-DriveEncryption -MountPoint $MountPoint }
        'Suspend'       { Suspend-DriveProtection -MountPoint $MountPoint }
        'Resume'        { Resume-DriveProtection  -MountPoint $MountPoint }
        'ShowKeys'      { Show-RecoveryKey        -MountPoint $MountPoint }
        'BackupAD'      { Backup-RecoveryKey      -MountPoint $MountPoint -Target AD }
        'BackupEntraID' { Backup-RecoveryKey      -MountPoint $MountPoint -Target EntraID }
        'Export'        { Export-EncryptionReport }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

$DefaultMount = "$($Drive.ToUpper().TrimEnd(':')):"

try {
    if ($Unattended) {
        if ($WhatIf) { Write-Host "  [~] DRY RUN - no changes will be made." -ForegroundColor Cyan }
        Invoke-CipherAction -Name $Action -MountPoint $DefaultMount
    } else {
        $menu = [ordered]@{
            '1' = @('Enable',        'Enable BitLocker')
            '2' = @('Disable',       'Disable BitLocker (decrypt)')
            '3' = @('Suspend',       'Suspend protection for one reboot')
            '4' = @('Resume',        'Resume protection')
            '5' = @('ShowKeys',      'Show recovery keys')
            '6' = @('BackupAD',      'Back up recovery key to Active Directory')
            '7' = @('BackupEntraID', 'Back up recovery key to Entra ID')
            '8' = @('Export',        'Export report (status + recovery keys)')
        }
        $destructive = @('Disable', 'Suspend')

        while ($true) {
            Clear-Host
            Write-Host ""
            Write-Host "  C.I.P.H.E.R. — Configures & Implements Policy-based Hardware Encryption & Recovery" -ForegroundColor $ColorSchema.Header
            if ($WhatIf) { Write-Host "  [~] DRY RUN - no changes will be made." -ForegroundColor Cyan }
            Show-DriveStatus

            foreach ($key in $menu.Keys) { Write-Host "  [$key] $($menu[$key][1])" -ForegroundColor $ColorSchema.Info }
            Write-Host "  [Q] Quit" -ForegroundColor $ColorSchema.Info
            Write-Host ""
            $choice = (Read-Host '  Selection').Trim().ToUpper()
            if ($choice -eq 'Q') { break }
            if (-not $menu.Contains($choice)) { continue }

            $name  = $menu[$choice][0]
            $mount = $DefaultMount
            if ($name -ne 'Export') {
                $letter = (Read-Host "  Drive letter [$($DefaultMount.TrimEnd(':'))]").Trim().TrimEnd(':').ToUpper()
                if ($letter -match '^[A-Z]$') { $mount = "${letter}:" }
            }
            $confirmed = $WhatIf -or ($name -notin $destructive) -or
                         ((Read-Host "  $name BitLocker on $mount? (Y/N)").Trim().ToUpper() -eq 'Y')
            if ($confirmed) { Invoke-CipherAction -Name $name -MountPoint $mount }

            Write-Host ""
            Read-Host '  Press Enter to return to the menu' | Out-Null
        }
    }
} catch {
    Write-Fail "BitLocker could not be read: $($_.Exception.Message)"
    Write-Info 'BitLocker needs a Pro, Enterprise or Education edition of Windows.'
    Write-TKError -ScriptName 'cipher' -Message $_.Exception.Message -Category 'BitLocker'
}

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
