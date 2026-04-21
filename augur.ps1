<#
.SYNOPSIS
    A.U.G.U.R. — Analyzes, Uncovers & Gauges Unit Reliability
    Disk Health & SMART Assessment Tool for PowerShell 5.1+

.DESCRIPTION
    Inspects every physical disk in the system: health status, operational
    state, SMART failure prediction, volume integrity, and bus/media type.
    Flags disks that report degraded health or predicted failures and exports
    a dark-themed HTML report to the script directory.

.USAGE
    PS C:\> .\augur.ps1                    # Must be run as Administrator
    PS C:\> .\augur.ps1 -Unattended        # Silent mode — no prompts, no banner

.NOTES
    Version : 1.1

    Tools Available
    ─────────────────────────────────────────────────────────────────
    G.R.I.M.O.I.R.E.       — Technician Toolkit hub and central launcher
    O.R.A.C.L.E.           — System diagnostics & HTML report generation
    A.R.C.H.I.V.E.         — Pre-reimaging profile backup
    C.L.E.A.N.S.E.         — Disk cleanup — temp, update cache, browser caches
    A.U.G.U.R.             — Disk health assessment, SMART & reliability analysis

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success messages
    Yellow   Warnings and cautions
    Red      Critical errors
    Gray     Information and details
#>

param([switch]$Unattended)

# ─────────────────────────────────────────────────────────────────────────────
# INITIALIZATION
# ─────────────────────────────────────────────────────────────────────────────

Import-Module "$PSScriptRoot\TechnicianToolkit.psm1" -Force
Invoke-AdminElevation -ScriptFile $PSCommandPath

$ScriptPath = $PSScriptRoot

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

function Show-AugurBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

   █████╗ ██╗   ██╗ ██████╗ ██╗   ██╗██████╗
  ██╔══██╗██║   ██║██╔════╝ ██║   ██║██╔══██╗
  ███████║██║   ██║██║  ███╗██║   ██║██████╔╝
  ██╔══██║██║   ██║██║   ██║██║   ██║██╔══██╗
  ██║  ██║╚██████╔╝╚██████╔╝╚██████╔╝██║  ██║
  ╚═╝  ╚═╝ ╚═════╝  ╚═════╝  ╚═════╝ ╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    A.U.G.U.R. — Analyzes, Uncovers & Gauges Unit Reliability" -ForegroundColor Cyan
    Write-Host "    Disk Health & SMART Assessment Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function HtmlEncode([string]$s) {
    $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if (-not $Unattended) { Show-AugurBanner }

$reportTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$reportFilename  = "AUGUR_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$reportPath      = Join-Path $ScriptPath $reportFilename

Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  A.U.G.U.R. DISK HEALTH ASSESSMENT" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  Machine   : $env:COMPUTERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  Timestamp : $reportTimestamp" -ForegroundColor $ColorSchema.Info
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1: PHYSICAL DISKS
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "  [1/3] Reading physical disk health..." -ForegroundColor $ColorSchema.Progress

$diskReport = @()

try {
    $physDisks = Get-PhysicalDisk -ErrorAction Stop

    # SMART failure prediction via WMI
    $smartData = @{}
    try {
        $smartRaw = Get-WmiObject -Namespace root\wmi -Class MSStorageDriver_FailurePredictStatus -ErrorAction SilentlyContinue
        foreach ($s in $smartRaw) {
            # InstanceName typically ends in _0, _1, etc. — use it as a loose key
            $key = ($s.InstanceName -split '\\' | Select-Object -Last 1) -replace '_\d+$',''
            $smartData[$key] = $s
        }
    } catch {}

    # Win32_DiskDrive for serial/firmware
    $wmiDisks = @{}
    try {
        Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue | ForEach-Object {
            $wmiDisks[$_.Index] = $_
        }
    } catch {}

    foreach ($pd in $physDisks) {
        $sizeGB = if ($pd.Size -gt 0) { [math]::Round($pd.Size / 1GB, 1) } else { 0 }

        # Match SMART entry by device index or friendly name fragment
        $smartEntry  = $smartData.Values | Where-Object { $_.InstanceName -match [regex]::Escape($pd.DeviceId) } | Select-Object -First 1
        $smartFail   = if ($smartEntry) { $smartEntry.PredictFailure } else { $null }
        $smartReason = if ($smartEntry -and $smartEntry.Reason) { "0x{0:X8}" -f $smartEntry.Reason } else { 'N/A' }

        $wmiDisk   = $wmiDisks[$pd.DeviceId]
        $serial    = if ($wmiDisk -and $wmiDisk.SerialNumber) { $wmiDisk.SerialNumber.Trim() } else { 'N/A' }
        $firmware  = if ($wmiDisk -and $wmiDisk.FirmwareRevision) { $wmiDisk.FirmwareRevision.Trim() } else { 'N/A' }

        $healthColor = switch ($pd.HealthStatus) {
            'Healthy' { $ColorSchema.Success }
            'Warning' { $ColorSchema.Warning }
            default   { $ColorSchema.Error   }
        }

        $smartLabel = if ($null -eq $smartFail)   { 'N/A' }
                      elseif ($smartFail -eq $false) { 'OK' }
                      else                           { 'FAILING' }

        $displayLine = "  [{0}] {1} | {2} {3} | {4} GB | Health: {5} | SMART: {6}" -f `
            $pd.DeviceId, $pd.FriendlyName, $pd.MediaType, $pd.BusType, $sizeGB, $pd.HealthStatus, $smartLabel

        $lineColor = if ($smartFail -eq $true -or $pd.HealthStatus -ne 'Healthy') {
            $ColorSchema.Error
        } elseif ($pd.HealthStatus -eq 'Warning') {
            $ColorSchema.Warning
        } else {
            $ColorSchema.Success
        }

        Write-Host $displayLine -ForegroundColor $lineColor

        $diskReport += [PSCustomObject]@{
            ID                = $pd.DeviceId
            Name              = $pd.FriendlyName
            Serial            = $serial
            Firmware          = $firmware
            MediaType         = $pd.MediaType
            BusType           = $pd.BusType
            SizeGB            = $sizeGB
            HealthStatus      = $pd.HealthStatus
            OperationalStatus = $pd.OperationalStatus
            SMARTPrediction   = $smartLabel
            SMARTReason       = $smartReason
        }
    }

    Write-Host ""
    Write-Host "  [+] $($diskReport.Count) physical disk(s) assessed" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "  [-] Error reading physical disks: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2: VOLUME HEALTH
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "  [2/3] Reading volume health..." -ForegroundColor $ColorSchema.Progress

$volumeReport = @()

try {
    $volumes = Get-Volume -ErrorAction Stop | Where-Object { $_.DriveLetter -or $_.FileSystemLabel }

    foreach ($vol in $volumes) {
        $totalGB = if ($vol.Size -gt 0) { [math]::Round($vol.Size / 1GB, 1) } else { 0 }
        $freeGB  = if ($vol.SizeRemaining -gt 0) { [math]::Round($vol.SizeRemaining / 1GB, 1) } else { 0 }
        $pct     = if ($vol.Size -gt 0) { [math]::Round(($vol.Size - $vol.SizeRemaining) / $vol.Size * 100, 1) } else { 0 }

        $healthColor = switch ($vol.HealthStatus) {
            'Healthy' { $ColorSchema.Success }
            'Warning' { $ColorSchema.Warning }
            default   { $ColorSchema.Error   }
        }

        $label  = if ($vol.DriveLetter)      { "$($vol.DriveLetter):" } else { "(no letter)" }
        $fsLabel = if ($vol.FileSystemLabel) { $vol.FileSystemLabel }   else { "" }

        Write-Host ("  {0,-6} {1,-20} {2,7} GB total  {3,7} GB free  {4,5}%  [{5}]" -f `
            $label, $fsLabel, $totalGB, $freeGB, $pct, $vol.HealthStatus) -ForegroundColor $healthColor

        $volumeReport += [PSCustomObject]@{
            Drive       = $label
            Label       = $fsLabel
            FileSystem  = $vol.FileSystem
            TotalGB     = $totalGB
            FreeGB      = $freeGB
            PctUsed     = $pct
            Health      = $vol.HealthStatus
            DriveType   = $vol.DriveType
        }
    }

    Write-Host ""
    Write-Host "  [+] $($volumeReport.Count) volume(s) assessed" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "  [-] Error reading volumes: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3: GENERATE HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "  [3/3] Generating HTML report..." -ForegroundColor $ColorSchema.Progress

$healthyCount   = ($diskReport | Where-Object { $_.HealthStatus -eq 'Healthy' -and $_.SMARTPrediction -ne 'FAILING' }).Count
$warningCount   = ($diskReport | Where-Object { $_.HealthStatus -eq 'Warning' }).Count
$criticalCount  = ($diskReport | Where-Object { $_.HealthStatus -ne 'Healthy' -and $_.HealthStatus -ne 'Warning' }).Count
$smartFailCount = ($diskReport | Where-Object { $_.SMARTPrediction -eq 'FAILING' }).Count

# Physical disk rows
$diskRows = ""
foreach ($d in $diskReport) {
    $hBadge = switch ($d.HealthStatus) {
        'Healthy' { 'tk-badge-ok' } 'Warning' { 'tk-badge-warn' } default { 'tk-badge-err' }
    }
    $sBadge = if ($d.SMARTPrediction -eq 'FAILING') { 'tk-badge-err' } elseif ($d.SMARTPrediction -eq 'OK') { 'tk-badge-ok' } else { 'tk-badge-info' }
    $diskRows += @"
    <tr>
      <td>$(HtmlEncode $d.ID)</td>
      <td>$(HtmlEncode $d.Name)</td>
      <td>$(HtmlEncode $d.Serial)</td>
      <td>$(HtmlEncode $d.MediaType)</td>
      <td>$(HtmlEncode $d.BusType)</td>
      <td>$($d.SizeGB) GB</td>
      <td><span class="$hBadge">$(HtmlEncode $d.HealthStatus)</span></td>
      <td>$(HtmlEncode $d.OperationalStatus)</td>
      <td><span class="$sBadge">$(HtmlEncode $d.SMARTPrediction)</span></td>
      <td><code>$(HtmlEncode $d.SMARTReason)</code></td>
      <td>$(HtmlEncode $d.Firmware)</td>
    </tr>
"@
}

# Volume rows
$volumeRows = ""
foreach ($v in $volumeReport) {
    $barBadge  = if ($v.PctUsed -ge 90) { 'tk-badge-err' } elseif ($v.PctUsed -ge 75) { 'tk-badge-warn' } else { 'tk-badge-ok' }
    $vhBadge   = switch ($v.Health) { 'Healthy' { 'tk-badge-ok' } 'Warning' { 'tk-badge-warn' } default { 'tk-badge-err' } }
    $volumeRows += @"
    <tr>
      <td><code>$(HtmlEncode $v.Drive)</code></td>
      <td>$(HtmlEncode $v.Label)</td>
      <td>$(HtmlEncode $v.FileSystem)</td>
      <td>$($v.TotalGB) GB</td>
      <td>$($v.FreeGB) GB</td>
      <td><span class="$barBadge">$($v.PctUsed)%</span></td>
      <td><span class="$vhBadge">$(HtmlEncode $v.Health)</span></td>
    </tr>
"@
}

$overallBadge = if ($criticalCount -gt 0 -or $smartFailCount -gt 0) {
    "<span class='tk-badge-err'>$($criticalCount + $smartFailCount) critical issue(s)</span>"
} elseif ($warningCount -gt 0) {
    "<span class='tk-badge-warn'>$warningCount warning(s)</span>"
} else {
    "<span class='tk-badge-ok'>All Healthy</span>"
}

$htmlHead = Get-TKHtmlHead `
    -Title     'Disk Health Assessment' `
    -ScriptName 'A.U.G.U.R.' `
    -Subtitle   $env:COMPUTERNAME `
    -MetaItems  ([ordered]@{
        'Machine'   = $env:COMPUTERNAME
        'Run As'    = "$env:USERDOMAIN\$env:USERNAME"
        'Generated' = $reportTimestamp
        'Overall'   = $overallBadge
    }) `
    -NavItems   @('Physical Disks', 'Volumes')

$htmlFoot = Get-TKHtmlFoot -ScriptName 'A.U.G.U.R. v1.1'

$htmlReport = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card ok"><div class="tk-summary-num">$healthyCount</div><div class="tk-summary-lbl">Healthy</div></div>
    <div class="tk-summary-card warn"><div class="tk-summary-num">$warningCount</div><div class="tk-summary-lbl">Warning</div></div>
    <div class="tk-summary-card err"><div class="tk-summary-num">$criticalCount</div><div class="tk-summary-lbl">Critical</div></div>
    <div class="tk-summary-card err"><div class="tk-summary-num">$smartFailCount</div><div class="tk-summary-lbl">SMART Fail</div></div>
  </div>

  <!-- PHYSICAL DISKS -->
  <div class="tk-card">
    <div class="tk-card-header">
      <span class="tk-card-label">Physical Disks ($($diskReport.Count))</span>
    </div>
    $(if ($diskRows) {
      "<table class='tk-table'><thead><tr><th>ID</th><th>Name</th><th>Serial</th><th>Type</th><th>Bus</th><th>Size</th><th>Health</th><th>Status</th><th>SMART</th><th>SMART Reason</th><th>Firmware</th></tr></thead><tbody>$diskRows</tbody></table>"
    } else {
      "<p class='tk-info-box'>No physical disk data available.</p>"
    })
  </div>

  <div class="tk-divider"></div>

  <!-- VOLUMES -->
  <div class="tk-card">
    <div class="tk-card-header">
      <span class="tk-card-label">Volumes ($($volumeReport.Count))</span>
    </div>
    $(if ($volumeRows) {
      "<table class='tk-table'><thead><tr><th>Drive</th><th>Label</th><th>File System</th><th>Total</th><th>Free</th><th>Usage</th><th>Health</th></tr></thead><tbody>$volumeRows</tbody></table>"
    } else {
      "<p class='tk-info-box'>No volume data available.</p>"
    })
  </div>

"@ + $htmlFoot

try {
    [System.IO.File]::WriteAllText($reportPath, $htmlReport, [System.Text.Encoding]::UTF8)
    Write-Host "  [+] Report saved: $reportPath" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "  [-] Could not save report: $_" -ForegroundColor $ColorSchema.Error
}

# ─────────────────────────────────────────────────────────────────────────────
# SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  A.U.G.U.R. ASSESSMENT COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host "  Physical Disks : $($diskReport.Count) assessed" -ForegroundColor $ColorSchema.Info

if ($criticalCount -gt 0) {
    Write-Host "  Critical       : $criticalCount disk(s) in critical state!" -ForegroundColor $ColorSchema.Error
} else {
    Write-Host "  Critical       : None" -ForegroundColor $ColorSchema.Success
}

if ($smartFailCount -gt 0) {
    Write-Host "  SMART Failure  : $smartFailCount disk(s) predicting failure!" -ForegroundColor $ColorSchema.Error
} else {
    Write-Host "  SMART Failure  : None predicted" -ForegroundColor $ColorSchema.Success
}

if ($warningCount -gt 0) {
    Write-Host "  Warnings       : $warningCount disk(s) in warning state" -ForegroundColor $ColorSchema.Warning
} else {
    Write-Host "  Healthy        : All disks healthy" -ForegroundColor $ColorSchema.Success
}

Write-Host ""
Write-Host "  Report         : $reportPath" -ForegroundColor $ColorSchema.Accent
Write-Host ""

if (-not $Unattended) {
    $open = Read-Host "  Open the HTML report now? (Y/N)"
    if ($open -eq 'Y' -or $open -eq 'y') {
        try {
            Start-Process $reportPath
            Write-Host "  [+] Opening report..." -ForegroundColor $ColorSchema.Success
        }
        catch {
            Write-Host "  [-] Could not open report. Navigate to: $reportPath" -ForegroundColor $ColorSchema.Warning
        }
    }
}

Write-Host ""
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
