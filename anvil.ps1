<#
.SYNOPSIS
    A.N.V.I.L. — Audits & Notates Vendor Inventory & Lifecycle
    BIOS / UEFI / Firmware Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Inventories system identity, BIOS / UEFI state, Secure Boot posture,
    and vendor-specific firmware-update channel availability (Dell Command,
    HP CMSL, Lenovo Vantage / System Update, Microsoft Surface). Scans
    Windows Update for pending firmware and driver updates. Produces a
    dark-themed HTML report with a readiness verdict so a technician can
    confirm a machine's firmware is supportable before shipping it.

.USAGE
    PS C:\> .\anvil.ps1                    # Interactive run
    PS C:\> .\anvil.ps1 -Unattended        # Silent: export HTML and exit

.NOTES
    Version : 3.5.1

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

function Show-AnvilBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  A.N.V.I.L. — Audits & Notates Vendor Inventory & Lifecycle" -ForegroundColor Cyan
    Write-Host "  BIOS / UEFI / Firmware Audit Tool  v3.5.1" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Data collection and report builders appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — SYSTEM IDENTITY + BIOS
# ─────────────────────────────────────────────────────────────────────────────

function Get-SystemInfo {
    $cs   = Get-CimInstance -ClassName Win32_ComputerSystem       -ErrorAction SilentlyContinue
    $bios = Get-CimInstance -ClassName Win32_BIOS                 -ErrorAction SilentlyContinue
    $sys  = Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction SilentlyContinue

    $releaseDate = $null
    if ($bios -and $bios.ReleaseDate) {
        try {
            $releaseDate = [Management.ManagementDateTimeConverter]::ToDateTime($bios.ReleaseDate)
        } catch {
            # Date parse failures happen on some virtualised BIOSes; leave $null.
        }
    }

    $biosAgeDays = if ($releaseDate) { [math]::Round(((Get-Date) - $releaseDate).TotalDays, 0) } else { $null }

    # Manufacturer is useful upstream for vendor-channel detection.
    $mfr = if ($cs -and $cs.Manufacturer) { $cs.Manufacturer } else { '' }
    $vendor = switch -Wildcard ($mfr.ToLower()) {
        '*dell*'    { 'Dell' }
        '*hp*'      { 'HP' }
        '*hewlett*' { 'HP' }
        '*lenovo*'  { 'Lenovo' }
        '*microsoft*' { 'Microsoft' }
        default     { 'Other' }
    }

    return [PSCustomObject]@{
        Manufacturer = $mfr
        Vendor       = $vendor
        Model        = if ($cs) { $cs.Model } else { '' }
        SystemSKU    = if ($sys) { $sys.IdentifyingNumber } else { '' }
        UUID         = if ($sys) { $sys.UUID } else { '' }
        SerialNumber = if ($bios) { $bios.SerialNumber } else { '' }
        BIOSVendor   = if ($bios) { $bios.Manufacturer } else { '' }
        BIOSVersion  = if ($bios) { $bios.SMBIOSBIOSVersion } else { '' }
        BIOSReleaseDate = $releaseDate
        BIOSAgeDays  = $biosAgeDays
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — UEFI / SECURE BOOT POSTURE
# ─────────────────────────────────────────────────────────────────────────────

function Get-UefiPosture {
    $bootMode = 'Unknown'
    try {
        # PartitionStyle GPT on the system disk is a decent proxy for UEFI boot.
        $sys = Get-Disk -Number ([int](Get-CimInstance -ClassName Win32_OperatingSystem).SystemDevice.Replace('\\.\PHYSICALDRIVE','')) -ErrorAction SilentlyContinue
        if (-not $sys) { $sys = Get-Disk | Where-Object { $_.IsSystem } | Select-Object -First 1 }
        if ($sys) {
            $bootMode = if ($sys.PartitionStyle -eq 'GPT') { 'UEFI' } else { 'Legacy/BIOS' }
        }
    } catch {
        # Fall through to $bootMode = 'Unknown'.
    }

    # Firmware type from the registry is the authoritative check but may be absent.
    $firmwareType = 'Unknown'
    try {
        $env = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control' -Name 'PEFirmwareType' -ErrorAction SilentlyContinue
        if ($env -and $env.PEFirmwareType) {
            $firmwareType = switch ([int]$env.PEFirmwareType) {
                1 { 'BIOS (legacy)' }
                2 { 'UEFI' }
                default { "Unknown ($($env.PEFirmwareType))" }
            }
        }
    } catch {
        # PEFirmwareType may not exist on older builds.
    }

    $secureBoot = $null
    try {
        $secureBoot = Confirm-SecureBootUEFI -ErrorAction SilentlyContinue
    } catch {
        # Confirm-SecureBootUEFI requires UEFI firmware; on legacy BIOS it throws.
        $secureBoot = $null
    }

    return [PSCustomObject]@{
        BootMode         = $bootMode
        FirmwareType     = $firmwareType
        SecureBootState  = if ($null -eq $secureBoot) { 'Unsupported' } elseif ($secureBoot) { 'Enabled' } else { 'Disabled' }
        SecureBoot       = $secureBoot
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — VENDOR-SPECIFIC UPDATE CHANNELS
# ─────────────────────────────────────────────────────────────────────────────

# Known install locations for per-vendor firmware-update tooling. We test for
# the tool binary on disk — presence is what matters to a technician ("can I
# check firmware updates on this machine?"). Running the tool is out of scope
# for an audit.
$script:VendorChannels = @(
    @{ Vendor = 'Dell';      Name = 'Dell Command | Update (CLI)';       Paths = @('C:\Program Files\Dell\CommandUpdate\dcu-cli.exe', 'C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe') }
    @{ Vendor = 'Dell';      Name = 'Dell Command | Update (GUI)';       Paths = @('C:\Program Files\Dell\CommandUpdate\DellCommandUpdate.exe', 'C:\Program Files (x86)\Dell\CommandUpdate\DellCommandUpdate.exe') }
    @{ Vendor = 'HP';        Name = 'HP Image Assistant (HPIA)';         Paths = @('C:\Program Files\HP\HPIA\HPImageAssistant.exe', 'C:\SWSetup\HPIA\HPImageAssistant.exe') }
    @{ Vendor = 'HP';        Name = 'HP Support Assistant';              Paths = @('C:\Program Files (x86)\Hewlett-Packard\HP Support Solutions\HPSF.exe') }
    @{ Vendor = 'Lenovo';    Name = 'Lenovo System Update (Tvsu)';       Paths = @('C:\Program Files (x86)\Lenovo\System Update\Tvsu.exe') }
    @{ Vendor = 'Lenovo';    Name = 'Lenovo Vantage';                    Paths = @("$env:LOCALAPPDATA\Packages\E046963F.LenovoSettingsforEnterprise_k1h2ywk1493x8") }
    @{ Vendor = 'Microsoft'; Name = 'Surface UEFI Configurator (SEMM)';  Paths = @('C:\Program Files (x86)\Microsoft\Surface\UefiConfigurator\UEFIConfigurator.exe') }
)

function Get-VendorChannelStatus {
    param([string]$Vendor)

    $applicable = $script:VendorChannels | Where-Object { $_.Vendor -eq $Vendor -or $Vendor -eq 'Other' }
    if (-not $applicable) {
        return @([PSCustomObject]@{
            Name      = "(no known firmware-update channel for '$Vendor')"
            Vendor    = $Vendor
            Found     = $false
            Path      = ''
        })
    }

    $rows = foreach ($ch in $applicable) {
        $foundPath = $null
        foreach ($p in $ch.Paths) {
            if (Test-Path $p) { $foundPath = $p; break }
        }
        [PSCustomObject]@{
            Name   = $ch.Name
            Vendor = $ch.Vendor
            Found  = [bool]$foundPath
            Path   = if ($foundPath) { $foundPath } else { '(not installed)' }
        }
    }

    return @($rows)
}

# Also check PSWindowsUpdate availability since the next section uses it.
function Test-PSWindowsUpdate {
    return [bool](Get-Module -ListAvailable -Name PSWindowsUpdate)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — WINDOWS UPDATE PENDING FIRMWARE / DRIVER SCAN
# ─────────────────────────────────────────────────────────────────────────────

function Get-PendingFirmwareUpdates {
    Write-Step "Scanning Windows Update for pending driver/firmware updates..."

    if (-not (Test-PSWindowsUpdate)) {
        Write-Warn "PSWindowsUpdate module not installed — skipping WU scan. Install with:"
        Write-Info "  Install-Module PSWindowsUpdate -Scope CurrentUser -Force"
        return @()
    }

    try {
        Import-Module PSWindowsUpdate -ErrorAction Stop
        # -Category 'Drivers' covers firmware as well on most manufacturer distributions.
        $updates = Get-WindowsUpdate -Category 'Drivers' -ErrorAction Stop
    } catch {
        Write-Fail "Windows Update scan failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'anvil' -Message "Get-WindowsUpdate driver scan failed: $($_.Exception.Message)" -Category 'Windows Update'
        return @()
    }

    if (-not $updates -or $updates.Count -eq 0) {
        Write-Ok "No pending driver/firmware updates found."
        return @()
    }

    $rows = foreach ($u in $updates) {
        [PSCustomObject]@{
            Title    = $u.Title
            KB       = if ($u.KB) { $u.KB } else { '—' }
            Size     = if ($u.Size) { Format-Bytes $u.Size } else { '—' }
            Severity = if ($u.MsrcSeverity) { $u.MsrcSeverity } else { 'Optional' }
        }
    }

    Write-Warn "$($rows.Count) pending driver/firmware update(s)."
    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# READINESS VERDICT
# ─────────────────────────────────────────────────────────────────────────────

function Get-Verdict {
    param($System, $Uefi, [array]$Channels, [array]$Pending)

    $issues = [System.Collections.Generic.List[string]]::new()
    $warns  = [System.Collections.Generic.List[string]]::new()

    if ($Uefi.FirmwareType -like 'BIOS*') {
        $issues.Add('Machine is booting in legacy BIOS mode — Secure Boot and modern OS requirements (Windows 11, BitLocker PCR-7) need UEFI.')
    }
    if ($Uefi.SecureBootState -eq 'Disabled') {
        $issues.Add('Secure Boot is disabled — required for Windows 11 and many security baselines.')
    } elseif ($Uefi.SecureBootState -eq 'Unsupported') {
        $warns.Add('Secure Boot unsupported — the firmware does not expose UEFI variables; likely legacy BIOS.')
    }

    if ($System.BIOSAgeDays -and $System.BIOSAgeDays -ge 730) {
        $warns.Add("BIOS is $([math]::Round($System.BIOSAgeDays / 365, 1)) years old ($($System.BIOSReleaseDate.ToString('yyyy-MM-dd'))) — review the vendor release channel for a newer version.")
    }

    $installedChannels = @($Channels | Where-Object { $_.Found })
    if ($System.Vendor -ne 'Other' -and $installedChannels.Count -eq 0) {
        $warns.Add("No $($System.Vendor) firmware-update tooling detected — deploying DCU / HPIA / Lenovo System Update / Vantage makes firmware maintenance trackable.")
    }

    if ($Pending.Count -gt 0) {
        $warns.Add("$($Pending.Count) pending driver/firmware update(s) are available via Windows Update.")
    }

    $verdict = if ($issues.Count -gt 0) { 'ATTENTION REQUIRED' }
               elseif ($warns.Count -gt 0) { 'REVIEW RECOMMENDED' }
               else { 'READY' }
    $class   = if ($issues.Count -gt 0) { 'err' }
               elseif ($warns.Count -gt 0) { 'warn' }
               else { 'ok' }

    return [PSCustomObject]@{
        Verdict = $verdict
        Class   = $class
        Issues  = @($issues)
        Warns   = @($warns)
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param($System, $Uefi, [array]$Channels, [array]$Pending, $Verdict)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $biosAgeStr = if ($System.BIOSAgeDays) { "$($System.BIOSAgeDays) days" } else { 'Unknown' }
    $biosDateStr = if ($System.BIOSReleaseDate) { $System.BIOSReleaseDate.ToString('yyyy-MM-dd') } else { 'Unknown' }

    $sbClass = switch ($Uefi.SecureBootState) {
        'Enabled'     { 'tk-badge-ok' }
        'Disabled'    { 'tk-badge-err' }
        'Unsupported' { 'tk-badge-warn' }
        default       { 'tk-badge-info' }
    }
    $fwClass = if ($Uefi.FirmwareType -eq 'UEFI') { 'tk-badge-ok' } elseif ($Uefi.FirmwareType -like 'BIOS*') { 'tk-badge-err' } else { 'tk-badge-warn' }

    $chRows = [System.Text.StringBuilder]::new()
    foreach ($c in $Channels) {
        $badge = if ($c.Found) { "<span class='tk-badge-ok'>Installed</span>" } else { "<span class='tk-badge-warn'>Not installed</span>" }
        [void]$chRows.Append("<tr><td>$(EscHtml $c.Vendor)</td><td>$(EscHtml $c.Name)</td><td>$badge</td><td><code>$(EscHtml $c.Path)</code></td></tr>`n")
    }

    $pendingRows = [System.Text.StringBuilder]::new()
    if ($Pending.Count -eq 0) {
        [void]$pendingRows.Append("<tr><td colspan='4' class='tk-badge-ok' style='text-align:center;'>No pending driver/firmware updates.</td></tr>")
    } else {
        foreach ($p in $Pending) {
            [void]$pendingRows.Append("<tr><td>$(EscHtml $p.Title)</td><td>$(EscHtml $p.KB)</td><td>$(EscHtml $p.Size)</td><td>$(EscHtml $p.Severity)</td></tr>`n")
        }
    }

    $verdictBlock = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$verdictBlock.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>`n") }
    foreach ($w in $Verdict.Warns)  { [void]$verdictBlock.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>`n") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$verdictBlock.Append("<li class='tk-badge-ok'>Firmware posture is clean.</li>")
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'A.N.V.I.L. Firmware / BIOS Audit' `
        -ScriptName 'A.N.V.I.L.' `
        -Subtitle   "${orgPrefix}BIOS / UEFI / Firmware Posture -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'      = $machine
            'Manufacturer' = $System.Manufacturer
            'Model'        = $System.Model
            'Serial'       = $System.SerialNumber
            'Generated'    = $reportDate
            'Verdict'      = $Verdict.Verdict
        }) `
        -NavItems   @('Verdict', 'System Identity', 'UEFI / Secure Boot', 'Vendor Channels', 'Pending Updates')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'A.N.V.I.L. v3.5'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Firmware Readiness</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(EscHtml $System.Vendor)</div><div class="tk-summary-lbl">Vendor</div></div>
    <div class="tk-summary-card $fwClass"><div class="tk-summary-num">$(EscHtml $Uefi.FirmwareType)</div><div class="tk-summary-lbl">Firmware Type</div></div>
    <div class="tk-summary-card $sbClass"><div class="tk-summary-num">$(EscHtml $Uefi.SecureBootState)</div><div class="tk-summary-lbl">Secure Boot</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$biosAgeStr</div><div class="tk-summary-lbl">BIOS Age</div></div>
    <div class="tk-summary-card $(if ($Pending.Count -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Pending.Count)</div><div class="tk-summary-lbl">Pending WU Updates</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Verdict &amp; Findings</span><span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span></div>
    <div class="tk-card"><ul class="tk-info-box" style="list-style:none;padding-left:0;">$($verdictBlock.ToString())</ul></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">System Identity</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <tbody>
          <tr><th>Manufacturer</th><td>$(EscHtml $System.Manufacturer)</td></tr>
          <tr><th>Model</th><td>$(EscHtml $System.Model)</td></tr>
          <tr><th>System SKU / Service Tag</th><td><code>$(EscHtml $System.SystemSKU)</code></td></tr>
          <tr><th>Serial Number</th><td><code>$(EscHtml $System.SerialNumber)</code></td></tr>
          <tr><th>UUID</th><td><code>$(EscHtml $System.UUID)</code></td></tr>
          <tr><th>BIOS Vendor</th><td>$(EscHtml $System.BIOSVendor)</td></tr>
          <tr><th>BIOS Version</th><td><code>$(EscHtml $System.BIOSVersion)</code></td></tr>
          <tr><th>BIOS Release Date</th><td>$biosDateStr</td></tr>
          <tr><th>BIOS Age</th><td>$biosAgeStr</td></tr>
        </tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">UEFI / Secure Boot</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <tbody>
          <tr><th>Firmware Type</th><td><span class='$fwClass'>$(EscHtml $Uefi.FirmwareType)</span></td></tr>
          <tr><th>Boot Mode (from system disk)</th><td>$(EscHtml $Uefi.BootMode)</td></tr>
          <tr><th>Secure Boot State</th><td><span class='$sbClass'>$(EscHtml $Uefi.SecureBootState)</span></td></tr>
        </tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Vendor Firmware-Update Channels</span><span class="tk-section-num">$($Channels.Count) channel(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Vendor</th><th>Tool</th><th>Status</th><th>Path</th></tr></thead>
        <tbody>$($chRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Pending Driver / Firmware Updates (Windows Update)</span><span class="tk-section-num">$($Pending.Count)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Title</th><th>KB</th><th>Size</th><th>Severity</th></tr></thead>
        <tbody>$($pendingRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-AnvilBanner

Write-Section "SYSTEM IDENTITY"
$system = Get-SystemInfo
Write-Ok ("Manufacturer: {0} ({1})" -f $system.Manufacturer, $system.Vendor)
Write-Ok ("Model       : {0}" -f $system.Model)
Write-Ok ("Serial      : {0}" -f $system.SerialNumber)
Write-Ok ("BIOS        : {0} v{1}" -f $system.BIOSVendor, $system.BIOSVersion)
if ($system.BIOSReleaseDate) {
    $age = [math]::Round($system.BIOSAgeDays / 365, 1)
    $color = if ($system.BIOSAgeDays -ge 730) { $C.Warning } else { $C.Success }
    Write-Host ("  BIOS released {0} ({1} yr old)" -f $system.BIOSReleaseDate.ToString('yyyy-MM-dd'), $age) -ForegroundColor $color
}
Write-Host ""

Write-Section "UEFI / SECURE BOOT"
$uefi = Get-UefiPosture
$fwColor = if ($uefi.FirmwareType -eq 'UEFI') { $C.Success } elseif ($uefi.FirmwareType -like 'BIOS*') { $C.Error } else { $C.Warning }
Write-Host ("  Firmware Type : {0}" -f $uefi.FirmwareType) -ForegroundColor $fwColor
Write-Host ("  Boot Mode     : {0}" -f $uefi.BootMode) -ForegroundColor $C.Info
$sbColor = if ($uefi.SecureBootState -eq 'Enabled') { $C.Success } elseif ($uefi.SecureBootState -eq 'Disabled') { $C.Error } else { $C.Warning }
Write-Host ("  Secure Boot   : {0}" -f $uefi.SecureBootState) -ForegroundColor $sbColor
Write-Host ""

Write-Section "VENDOR FIRMWARE CHANNELS"
$channels = Get-VendorChannelStatus -Vendor $system.Vendor
foreach ($ch in $channels) {
    $color = if ($ch.Found) { $C.Success } else { $C.Info }
    Write-Host ("  [{0,-10}] {1,-40} {2}" -f $ch.Vendor, $ch.Name, $(if ($ch.Found) { 'Installed' } else { 'Not installed' })) -ForegroundColor $color
}
Write-Host ""

Write-Section "WINDOWS UPDATE FIRMWARE / DRIVER SCAN"
$pending = Get-PendingFirmwareUpdates
if ($pending.Count -gt 0) {
    foreach ($p in $pending) {
        Write-Host ("  [{0}] {1}" -f $p.Severity, $p.Title) -ForegroundColor $C.Warning
    }
}
Write-Host ""

Write-Section "FIRMWARE READINESS VERDICT"
$verdict = Get-Verdict -System $system -Uefi $uefi -Channels $channels -Pending $pending
$vColor = switch ($verdict.Class) { 'ok' { $C.Success } 'warn' { $C.Warning } default { $C.Error } }
Write-Host "  $($verdict.Verdict)" -ForegroundColor $vColor
foreach ($i in $verdict.Issues) { Write-Host "    [!!] $i" -ForegroundColor $C.Error }
foreach ($w in $verdict.Warns)  { Write-Host "    [~ ] $w" -ForegroundColor $C.Warning }
if ($verdict.Issues.Count -eq 0 -and $verdict.Warns.Count -eq 0) {
    Write-Host "    [+ ] All checks passed." -ForegroundColor $C.Success
}
Write-Host ""

Write-Step "Generating HTML report..."
$html      = Build-HtmlReport -System $system -Uefi $uefi -Channels $channels -Pending $pending -Verdict $verdict
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "ANVIL_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Write-Ok "Report saved: $outPath"
    if (-not $Unattended) {
        Write-Step "Opening in default browser..."
        Start-Process $outPath
    }
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath -and -not (Test-Path (Join-Path $PSScriptRoot '.git'))) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
