<#
.SYNOPSIS
    T.O.T.E.M. — Trusted Observer of Transparent Execution Modules
    TPM Health Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Audits the state of the Trusted Platform Module -- presence,
    specification version, manufacturer and firmware, ownership, clear
    state, attestation readiness -- and cross-references against
    BitLocker to confirm which volumes actually depend on the TPM's
    protector chain. Produces a dark-themed HTML report with a red /
    yellow / green readiness verdict suitable for a Windows 11 /
    Autopilot / BitLocker readiness gate.

.USAGE
    PS C:\> .\totem.ps1                    # Interactive run
    PS C:\> .\totem.ps1 -Unattended        # Silent: export HTML and exit

.NOTES
    Version : 3.6

#>

param(
    [switch]$Unattended,
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
Invoke-AdminElevation -ScriptFile $PSCommandPath

if ($PSScriptRoot) {
    $ScriptPath = $PSScriptRoot
} elseif ($PSCommandPath) {
    $ScriptPath = Split-Path -Parent $PSCommandPath
} else {
    $ScriptPath = (Get-Location).Path
}

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $ScriptPath) }

$C = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
}

function Show-TotemBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  T.O.T.E.M. — Trusted Observer of Transparent Execution Modules" -ForegroundColor Cyan
    Write-Host "  TPM Health Audit Tool  v3.6" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Collectors, verdict, and report builders appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — TPM STATUS
# ─────────────────────────────────────────────────────────────────────────────

function Get-TpmStatus {
    # Get-Tpm is the canonical way to read TPM state on Windows 10/11. It's in
    # the TrustedPlatformModule module which is present in-box.
    $tpm = $null
    try {
        $tpm = Get-Tpm -ErrorAction Stop
    } catch {
        return [PSCustomObject]@{
            Present          = $false
            Ready            = $false
            Enabled          = $false
            Activated        = $false
            Owned            = $false
            ManufacturerId   = $null
            ManufacturerName = $null
            ManufacturerVersion = $null
            SpecVersion      = $null
            PhysicalPresence = $null
            AutoProvisioning = $null
            RestartPending   = $false
            CollectorError   = $_.Exception.Message
        }
    }

    return [PSCustomObject]@{
        Present             = [bool]$tpm.TpmPresent
        Ready               = [bool]$tpm.TpmReady
        Enabled             = [bool]$tpm.TpmEnabled
        Activated           = [bool]$tpm.TpmActivated
        Owned               = [bool]$tpm.TpmOwned
        ManufacturerId      = $tpm.ManufacturerId
        ManufacturerName    = $tpm.ManufacturerIdTxt
        ManufacturerVersion = $tpm.ManufacturerVersion
        SpecVersion         = $tpm.SpecVersion
        PhysicalPresence    = $tpm.PhysicalPresenceVersionInfo
        AutoProvisioning    = if ($null -ne $tpm.AutoProvisioning) { $tpm.AutoProvisioning.ToString() } else { 'Unknown' }
        RestartPending      = [bool]$tpm.RestartPending
        CollectorError      = $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — SPEC VERSION + MANUFACTURER PARSING
# ─────────────────────────────────────────────────────────────────────────────

function Get-TpmVersionSummary {
    param($Tpm)

    # SpecVersion is a comma-separated string like "2.0, 0, 1.38". First token
    # is the TPM spec. Second is the family level (rarely useful). Third is
    # the rev of the TPM library spec.
    $spec = 'Unknown'
    if ($Tpm.SpecVersion) {
        $first = ($Tpm.SpecVersion -split ',')[0].Trim()
        if ($first) { $spec = $first }
    }

    $label = switch ($spec) {
        '2.0' { 'TPM 2.0' }
        '1.2' { 'TPM 1.2' }
        default { "TPM $spec" }
    }

    return [PSCustomObject]@{
        SpecVersion   = $spec
        Label         = $label
        IsWin11Ready  = ($spec -eq '2.0')
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — BITLOCKER DEPENDENCY
# ─────────────────────────────────────────────────────────────────────────────

# Answers "if this TPM goes bad, which drives can't unlock?" -- any volume with
# a Tpm, TpmPin, TpmPinStartupKey, or RecoveryPassword protector chain hangs
# on TPM measurements.
function Get-BitLockerTpmDependency {
    $result = [System.Collections.Generic.List[object]]::new()

    try {
        $vols = Get-BitLockerVolume -ErrorAction Stop
    } catch {
        return @($result)
    }

    foreach ($v in $vols) {
        $tpmProtectors = @($v.KeyProtector | Where-Object {
            $_.KeyProtectorType -in @('Tpm','TpmPin','TpmPinStartupKey','TpmStartupKey')
        })
        $result.Add([PSCustomObject]@{
            MountPoint        = $v.MountPoint
            VolumeType        = $v.VolumeType
            ProtectionStatus  = $v.ProtectionStatus
            EncryptionMethod  = $v.EncryptionMethod
            Protectors        = (@($v.KeyProtector | Select-Object -ExpandProperty KeyProtectorType) -join ', ')
            TpmProtectorCount = $tpmProtectors.Count
            DependsOnTpm      = ($tpmProtectors.Count -gt 0)
        })
    }

    return @($result)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — ATTESTATION / EK INFO
# ─────────────────────────────────────────────────────────────────────────────

function Get-TpmAttestationInfo {
    $result = [PSCustomObject]@{
        EkPresent      = $null
        EkAlgorithms   = $null
        CollectorError = $null
    }

    try {
        $ek = Get-TpmEndorsementKeyInfo -ErrorAction Stop
        if ($ek) {
            $result.EkPresent    = $true
            # EK has manufacturer certificates plus RSA / ECC algorithm slots;
            # caller cares about presence, not the cert chain itself.
            $algos = @()
            if ($ek.IsPresent)      { $algos += 'EK present' }
            if ($ek.ManufacturerCertificates -and $ek.ManufacturerCertificates.Count -gt 0) {
                $algos += "$($ek.ManufacturerCertificates.Count) manufacturer cert(s)"
            }
            $result.EkAlgorithms = ($algos -join '; ')
        } else {
            $result.EkPresent = $false
        }
    } catch {
        $result.EkPresent      = $null
        $result.CollectorError = $_.Exception.Message
    }

    return $result
}

# ─────────────────────────────────────────────────────────────────────────────
# READINESS VERDICT
# ─────────────────────────────────────────────────────────────────────────────

function Get-Verdict {
    param($Tpm, $Version, [array]$BitLockerDependencies, $Attestation)

    $issues = [System.Collections.Generic.List[string]]::new()
    $warns  = [System.Collections.Generic.List[string]]::new()

    if (-not $Tpm.Present) {
        $issues.Add('No TPM is present on this machine — Windows 11 upgrade and BitLocker key protection are blocked.')
    } else {
        if (-not $Tpm.Enabled)   { $issues.Add('TPM is present but DISABLED in firmware — enable in BIOS/UEFI Security settings.') }
        if (-not $Tpm.Activated) { $issues.Add('TPM is DEACTIVATED — activate in BIOS/UEFI and take ownership.') }
        if (-not $Tpm.Ready)     { $warns.Add('TPM reports NOT READY — may need Initialize-Tpm to provision, or a clear-and-reprovision cycle.') }
        if (-not $Tpm.Owned)     { $warns.Add('TPM has no owner — BitLocker provisioning will handle this on first encrypt, but document the state before proceeding.') }
        if ($Tpm.RestartPending) { $warns.Add('TPM reports a pending restart — commit outstanding state changes before audit-time conclusions.') }
        if (-not $Version.IsWin11Ready) {
            $warns.Add("This machine has $($Version.Label), not TPM 2.0 — Windows 11 upgrade is blocked until firmware offers a 2.0-capable chip (vendor BIOS update or dTPM->fTPM toggle may help).")
        }
    }

    # BitLocker cross-check
    $dependent = @($BitLockerDependencies | Where-Object { $_.DependsOnTpm })
    $anyBitLocker = @($BitLockerDependencies | Where-Object { $_.ProtectionStatus -eq 'On' }).Count
    if ($Tpm.Present -and $Tpm.Ready -and $dependent.Count -eq 0 -and $anyBitLocker -gt 0) {
        $warns.Add('BitLocker is enabled but none of the protected volumes use a TPM-based key protector — consider adding TPM for seamless unlock.')
    }

    if ($null -eq $Attestation.EkPresent -and $Tpm.Present) {
        $warns.Add('Endorsement Key info could not be read — Autopilot device pre-registration and attested boot will fail until the EK is retrievable.')
    }

    $verdict = if ($issues.Count -gt 0) { 'NOT READY' }
               elseif ($warns.Count -gt 0) { 'READY WITH WARNINGS' }
               else { 'READY' }
    $class   = if ($issues.Count -gt 0) { 'err' }
               elseif ($warns.Count -gt 0) { 'warn' }
               else { 'ok' }

    return [PSCustomObject]@{
        Verdict = $verdict
        Class   = $class
        Issues  = @($issues)
        Warns   = @($warns)
        BitLockerVolumesDependingOnTpm = $dependent.Count
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param($Tpm, $Version, [array]$BitLocker, $Attestation, $Verdict)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    # TPM status badges
    function _yn { param($b) if ($b) { "<span class='tk-badge-ok'>Yes</span>" } else { "<span class='tk-badge-err'>No</span>" } }

    $specClass = if ($Version.IsWin11Ready) { 'tk-badge-ok' } else { 'tk-badge-warn' }

    $blRows = [System.Text.StringBuilder]::new()
    if ($BitLocker.Count -eq 0) {
        [void]$blRows.Append("<tr><td colspan='5' class='tk-badge-info' style='text-align:center;'>No BitLocker-managed volumes found.</td></tr>")
    } else {
        foreach ($v in $BitLocker) {
            $protState = switch ($v.ProtectionStatus) {
                'On'  { "<span class='tk-badge-ok'>On</span>" }
                'Off' { "<span class='tk-badge-warn'>Off</span>" }
                default { "<span class='tk-badge-info'>$(EscHtml $v.ProtectionStatus)</span>" }
            }
            $tpmDep = if ($v.DependsOnTpm) { "<span class='tk-badge-ok'>Yes ($($v.TpmProtectorCount))</span>" } else { "<span class='tk-badge-warn'>No</span>" }
            [void]$blRows.Append("<tr><td>$(EscHtml $v.MountPoint)</td><td>$(EscHtml $v.VolumeType)</td><td>$protState</td><td>$(EscHtml $v.Protectors)</td><td>$tpmDep</td></tr>`n")
        }
    }

    $verdictBlock = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$verdictBlock.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>`n") }
    foreach ($w in $Verdict.Warns)  { [void]$verdictBlock.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>`n") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$verdictBlock.Append("<li class='tk-badge-ok'>TPM posture is clean for Windows 11, BitLocker, and Autopilot.</li>")
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'T.O.T.E.M. TPM Health Report' `
        -ScriptName 'T.O.T.E.M.' `
        -Subtitle   "${orgPrefix}TPM Health Audit -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'     = $machine
            'Generated'   = $reportDate
            'TPM Spec'    = $Version.Label
            'Verdict'     = $Verdict.Verdict
        }) `
        -NavItems   @('Verdict', 'TPM Status', 'BitLocker Dependency', 'Attestation')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'T.O.T.E.M. v3.6'

    $collectorErrCell = if ($Tpm.CollectorError) {
        "<tr><th>Collector Error</th><td><span class='tk-badge-err'>$(EscHtml $Tpm.CollectorError)</span></td></tr>"
    } else { '' }

    $attErrCell = if ($Attestation.CollectorError) {
        "<tr><th>EK Read Error</th><td><span class='tk-badge-warn'>$(EscHtml $Attestation.CollectorError)</span></td></tr>"
    } else { '' }

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">TPM Readiness</div></div>
    <div class="tk-summary-card $(if ($Tpm.Present) { 'ok' } else { 'err' })"><div class="tk-summary-num">$(if ($Tpm.Present) { 'Present' } else { 'Absent' })</div><div class="tk-summary-lbl">TPM Hardware</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(EscHtml $Version.Label)</div><div class="tk-summary-lbl">Specification</div></div>
    <div class="tk-summary-card $(if ($Tpm.Ready) { 'ok' } else { 'warn' })"><div class="tk-summary-num">$(if ($Tpm.Ready) { 'Yes' } else { 'No' })</div><div class="tk-summary-lbl">TPM Ready</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Verdict.BitLockerVolumesDependingOnTpm)</div><div class="tk-summary-lbl">BitLocker Volumes Using TPM</div></div>
    <div class="tk-summary-card $(if ($Attestation.EkPresent) { 'ok' } elseif ($null -eq $Attestation.EkPresent) { 'warn' } else { 'err' })"><div class="tk-summary-num">$(if ($Attestation.EkPresent) { 'Yes' } elseif ($null -eq $Attestation.EkPresent) { 'Unknown' } else { 'No' })</div><div class="tk-summary-lbl">Endorsement Key</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Verdict &amp; Findings</span><span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span></div>
    <div class="tk-card"><ul class="tk-info-box" style="list-style:none;padding-left:0;">$($verdictBlock.ToString())</ul></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">TPM Status</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <tbody>
          <tr><th>Present</th><td>$(_yn $Tpm.Present)</td></tr>
          <tr><th>Enabled (firmware)</th><td>$(_yn $Tpm.Enabled)</td></tr>
          <tr><th>Activated</th><td>$(_yn $Tpm.Activated)</td></tr>
          <tr><th>Ready for Use</th><td>$(_yn $Tpm.Ready)</td></tr>
          <tr><th>Owned</th><td>$(_yn $Tpm.Owned)</td></tr>
          <tr><th>Specification</th><td><span class='$specClass'>$(EscHtml $Version.Label)</span></td></tr>
          <tr><th>Manufacturer (raw ID)</th><td><code>$(EscHtml $Tpm.ManufacturerId)</code></td></tr>
          <tr><th>Manufacturer (text)</th><td>$(EscHtml $Tpm.ManufacturerName)</td></tr>
          <tr><th>Manufacturer Version</th><td><code>$(EscHtml $Tpm.ManufacturerVersion)</code></td></tr>
          <tr><th>Physical Presence Version</th><td>$(EscHtml $Tpm.PhysicalPresence)</td></tr>
          <tr><th>Auto-Provisioning</th><td>$(EscHtml $Tpm.AutoProvisioning)</td></tr>
          <tr><th>Restart Pending</th><td>$(_yn $Tpm.RestartPending)</td></tr>
          $collectorErrCell
        </tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">BitLocker Dependency on TPM</span><span class="tk-section-num">$($BitLocker.Count) volume(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Mount</th><th>Volume Type</th><th>Protection</th><th>Protectors</th><th>Uses TPM?</th></tr></thead>
        <tbody>$($blRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Attestation / Endorsement Key</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <tbody>
          <tr><th>EK Present</th><td>$(if ($Attestation.EkPresent) { "<span class='tk-badge-ok'>Yes</span>" } elseif ($null -eq $Attestation.EkPresent) { "<span class='tk-badge-warn'>Unknown</span>" } else { "<span class='tk-badge-err'>No</span>" })</td></tr>
          <tr><th>EK Details</th><td>$(EscHtml $Attestation.EkAlgorithms)</td></tr>
          $attErrCell
        </tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-TotemBanner

Write-Section "TPM STATUS"
$tpm = Get-TpmStatus
if ($tpm.CollectorError) {
    Write-Fail "Get-Tpm failed: $($tpm.CollectorError)"
} else {
    $pColor = if ($tpm.Present) { $C.Success } else { $C.Error }
    Write-Host ("  Present   : {0}" -f $tpm.Present) -ForegroundColor $pColor
    if ($tpm.Present) {
        Write-Host ("  Enabled   : {0}" -f $tpm.Enabled) -ForegroundColor $(if ($tpm.Enabled) { $C.Success } else { $C.Error })
        Write-Host ("  Activated : {0}" -f $tpm.Activated) -ForegroundColor $(if ($tpm.Activated) { $C.Success } else { $C.Error })
        Write-Host ("  Ready     : {0}" -f $tpm.Ready) -ForegroundColor $(if ($tpm.Ready) { $C.Success } else { $C.Warning })
        Write-Host ("  Owned     : {0}" -f $tpm.Owned) -ForegroundColor $(if ($tpm.Owned) { $C.Success } else { $C.Warning })
        Write-Host ("  Mfr       : {0}" -f $tpm.ManufacturerName) -ForegroundColor $C.Info
    }
}
Write-Host ""

$version = Get-TpmVersionSummary -Tpm $tpm
Write-Section "SPECIFICATION"
$vColor = if ($version.IsWin11Ready) { $C.Success } else { $C.Warning }
Write-Host ("  {0}  (raw: {1})" -f $version.Label, $version.SpecVersion) -ForegroundColor $vColor
Write-Host ""

Write-Section "BITLOCKER DEPENDENCY"
$bitLocker = Get-BitLockerTpmDependency
if ($bitLocker.Count -eq 0) {
    Write-Warn "No BitLocker volumes found (or BitLocker module unavailable)."
} else {
    foreach ($v in $bitLocker) {
        $depLabel = if ($v.DependsOnTpm) { "uses TPM ($($v.TpmProtectorCount))" } else { 'no TPM protector' }
        $color = if ($v.ProtectionStatus -eq 'On' -and $v.DependsOnTpm) { $C.Success }
                 elseif ($v.ProtectionStatus -eq 'On') { $C.Warning }
                 else { $C.Info }
        Write-Host ("  {0,-4} {1,-16} Protection: {2,-3}  Protectors: {3}  ({4})" -f $v.MountPoint, $v.VolumeType, $v.ProtectionStatus, $v.Protectors, $depLabel) -ForegroundColor $color
    }
}
Write-Host ""

Write-Section "ATTESTATION / ENDORSEMENT KEY"
$attestation = Get-TpmAttestationInfo
if ($attestation.EkPresent) {
    Write-Ok "Endorsement Key info readable: $($attestation.EkAlgorithms)"
} elseif ($null -eq $attestation.EkPresent) {
    Write-Warn "Could not read EK info: $($attestation.CollectorError)"
} else {
    Write-Fail "No Endorsement Key data available from TPM."
}
Write-Host ""

$verdict = Get-Verdict -Tpm $tpm -Version $version -BitLockerDependencies $bitLocker -Attestation $attestation

Write-Section "TPM READINESS VERDICT"
$vColorOut = switch ($verdict.Class) { 'ok' { $C.Success } 'warn' { $C.Warning } default { $C.Error } }
Write-Host "  $($verdict.Verdict)" -ForegroundColor $vColorOut
foreach ($i in $verdict.Issues) { Write-Host "    [!!] $i" -ForegroundColor $C.Error }
foreach ($w in $verdict.Warns)  { Write-Host "    [~ ] $w" -ForegroundColor $C.Warning }
if ($verdict.Issues.Count -eq 0 -and $verdict.Warns.Count -eq 0) {
    Write-Host "    [+ ] All checks passed." -ForegroundColor $C.Success
}
Write-Host ""

Write-Step "Generating HTML report..."
$html      = Build-HtmlReport -Tpm $tpm -Version $version -BitLocker $bitLocker -Attestation $attestation -Verdict $verdict
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "TOTEM_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
