<#
.SYNOPSIS
    D.W.A.R.F. — Detects Wear, Audits Reliability & Forecasts Failures
    Disk Health & SMART Assessment Tool for PowerShell 5.1+

.DESCRIPTION
    Inspects every physical disk in the system: health status, operational
    state, SMART failure prediction, volume integrity, and bus/media type.
    Flags disks that report degraded health or predicted failures and exports
    a dark-themed HTML report to the script directory.

.USAGE
    PS C:\> .\dwarf.ps1                    # Must be run as Administrator
    PS C:\> .\dwarf.ps1 -Unattended        # Silent mode — no prompts, no banner

.NOTES
    Version : 1.0

    Tools Available
    ─────────────────────────────────────────────────────────────────
    G.R.I.M.O.I.R.E.       — Technician Toolkit hub and central launcher
    O.R.A.C.L.E.           — System diagnostics & HTML report generation
    A.R.C.H.I.V.E.         — Pre-reimaging profile backup
    P.U.R.G.E.             — Disk cleanup — temp, update cache, browser caches
    D.W.A.R.F.             — Disk wear & health assessment, SMART status

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

function Show-DwarfBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██████╗ ██╗    ██╗ █████╗ ██████╗ ███████╗
  ██╔══██╗██║    ██║██╔══██╗██╔══██╗██╔════╝
  ██║  ██║██║ █╗ ██║███████║██████╔╝█████╗
  ██║  ██║██║███╗██║██╔══██║██╔══██╗██╔══╝
  ██████╔╝╚███╔███╔╝██║  ██║██║  ██║██║
  ╚═════╝  ╚══╝╚══╝ ╚═╝  ╚═╝╚═╝  ╚═╝╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    D.W.A.R.F. — Detects Wear, Audits Reliability & Forecasts Failures" -ForegroundColor Cyan
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

if (-not $Unattended) { Show-DwarfBanner }

$reportTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$reportFilename  = "DWARF_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$reportPath      = Join-Path $ScriptPath $reportFilename

Write-Host ("  " + ("─" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  D.W.A.R.F. DISK HEALTH ASSESSMENT" -ForegroundColor $ColorSchema.Header
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
    $hColor = switch ($d.HealthStatus) {
        'Healthy' { '#2ecc71' } 'Warning' { '#f39c12' } default { '#e74c3c' }
    }
    $sColor = if ($d.SMARTPrediction -eq 'FAILING') { '#e74c3c' } elseif ($d.SMARTPrediction -eq 'OK') { '#2ecc71' } else { '#888' }
    $diskRows += @"
    <tr>
      <td>$(HtmlEncode $d.ID)</td>
      <td>$(HtmlEncode $d.Name)</td>
      <td>$(HtmlEncode $d.Serial)</td>
      <td>$(HtmlEncode $d.MediaType)</td>
      <td>$(HtmlEncode $d.BusType)</td>
      <td>$($d.SizeGB) GB</td>
      <td style="color:$hColor;font-weight:600;">$(HtmlEncode $d.HealthStatus)</td>
      <td>$(HtmlEncode $d.OperationalStatus)</td>
      <td style="color:$sColor;font-weight:600;">$(HtmlEncode $d.SMARTPrediction)</td>
      <td>$(HtmlEncode $d.SMARTReason)</td>
      <td>$(HtmlEncode $d.Firmware)</td>
    </tr>
"@
}

# Volume rows
$volumeRows = ""
foreach ($v in $volumeReport) {
    $barColor = if ($v.PctUsed -ge 90) { "#e74c3c" } elseif ($v.PctUsed -ge 75) { "#f39c12" } else { "#2ecc71" }
    $vhColor  = switch ($v.Health) { 'Healthy' { '#2ecc71' } 'Warning' { '#f39c12' } default { '#e74c3c' } }
    $volumeRows += @"
    <tr>
      <td>$(HtmlEncode $v.Drive)</td>
      <td>$(HtmlEncode $v.Label)</td>
      <td>$(HtmlEncode $v.FileSystem)</td>
      <td>$($v.TotalGB) GB</td>
      <td>$($v.FreeGB) GB</td>
      <td>
        <div style="background:#333;border-radius:4px;height:12px;width:120px;display:inline-block;">
          <div style="background:$barColor;width:$($v.PctUsed)%;height:12px;border-radius:4px;"></div>
        </div>
        $($v.PctUsed)%
      </td>
      <td style="color:$vhColor;font-weight:600;">$(HtmlEncode $v.Health)</td>
    </tr>
"@
}

$overallBadge = if ($criticalCount -gt 0 -or $smartFailCount -gt 0) {
    "<span class='badge badge-err'>$($criticalCount + $smartFailCount) critical issue(s)</span>"
} elseif ($warningCount -gt 0) {
    "<span class='badge badge-warn'>$warningCount warning(s)</span>"
} else {
    "<span class='badge badge-ok'>All Healthy</span>"
}

$htmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="dark">
<title>D.W.A.R.F. — $env:COMPUTERNAME — $reportTimestamp</title>
<style>
  :root { color-scheme: dark; }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', sans-serif; background: #1a1a2e; color: #e0e0e0; font-size: 14px; }
  header { background: linear-gradient(135deg, #0f3460, #16213e); padding: 28px 40px; border-bottom: 3px solid #00d4ff; }
  header h1 { color: #00d4ff; font-size: 2em; letter-spacing: 4px; font-weight: 700; }
  header p  { color: #aaa; margin-top: 6px; font-size: 0.9em; }
  header .meta { display: flex; gap: 30px; margin-top: 14px; flex-wrap: wrap; }
  header .meta span { color: #ccc; font-size: 0.85em; }
  header .meta strong { color: #00d4ff; }
  main { padding: 30px 40px; max-width: 1400px; margin: 0 auto; }
  .summary-cards { display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 28px; }
  .card { background: #16213e; border: 1px solid #0f3460; border-radius: 8px; padding: 16px 24px; min-width: 120px; text-align: center; }
  .card .val { font-size: 32px; font-weight: bold; color: #00d4ff; }
  .card .lbl { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 1px; margin-top: 4px; }
  .card.ok   .val { color: #2ecc71; }
  .card.warn .val { color: #f39c12; }
  .card.crit .val { color: #e74c3c; }
  section { background: #16213e; border-radius: 8px; margin-bottom: 24px; overflow: hidden; border: 1px solid #0f3460; }
  section h2 { background: #0f3460; color: #00d4ff; padding: 14px 20px; font-size: 1em;
               letter-spacing: 2px; text-transform: uppercase; display: flex; align-items: center; gap: 10px; }
  section .content { padding: 20px; overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 0.85em; }
  th { background: #0f3460; color: #00d4ff; padding: 10px 12px; text-align: left;
       font-weight: 600; letter-spacing: 1px; text-transform: uppercase; font-size: 0.78em; white-space: nowrap; }
  td { padding: 8px 12px; border-bottom: 1px solid #1e3a5f; color: #ccc; }
  tr:hover td { background: #1e3a5f; }
  .badge { display: inline-block; padding: 3px 10px; border-radius: 20px; font-size: 0.78em;
           font-weight: 700; letter-spacing: 1px; text-transform: uppercase; }
  .badge-ok   { background: #1a4a2e; color: #2ecc71; border: 1px solid #2ecc71; }
  .badge-warn { background: #4a3000; color: #f39c12; border: 1px solid #f39c12; }
  .badge-err  { background: #4a0000; color: #e74c3c; border: 1px solid #e74c3c; }
  footer { text-align: center; padding: 20px; color: #444; font-size: 0.8em; border-top: 1px solid #0f3460; }
</style>
</head>
<body>
<header>
  <h1>D.W.A.R.F.</h1>
  <p>Detects Wear, Audits Reliability &amp; Forecasts Failures</p>
  <div class="meta">
    <span><strong>Machine:</strong> $env:COMPUTERNAME</span>
    <span><strong>Run As:</strong> $env:USERDOMAIN\$env:USERNAME</span>
    <span><strong>Generated:</strong> $reportTimestamp</span>
    <span><strong>Overall:</strong> $overallBadge</span>
  </div>
</header>
<main>

  <div class="summary-cards">
    <div class="card ok"><div class="val">$healthyCount</div><div class="lbl">Healthy</div></div>
    <div class="card warn"><div class="val">$warningCount</div><div class="lbl">Warning</div></div>
    <div class="card crit"><div class="val">$criticalCount</div><div class="lbl">Critical</div></div>
    <div class="card crit"><div class="val">$smartFailCount</div><div class="lbl">SMART Fail</div></div>
  </div>

  <!-- PHYSICAL DISKS -->
  <section>
    <h2>Physical Disks ($($diskReport.Count))</h2>
    <div class="content">
      $(if ($diskRows) {
        "<table><thead><tr><th>ID</th><th>Name</th><th>Serial</th><th>Type</th><th>Bus</th><th>Size</th><th>Health</th><th>Status</th><th>SMART</th><th>SMART Reason</th><th>Firmware</th></tr></thead><tbody>$diskRows</tbody></table>"
      } else {
        "<p style='color:#666;font-style:italic;'>No physical disk data available.</p>"
      })
    </div>
  </section>

  <!-- VOLUMES -->
  <section>
    <h2>Volumes ($($volumeReport.Count))</h2>
    <div class="content">
      $(if ($volumeRows) {
        "<table><thead><tr><th>Drive</th><th>Label</th><th>File System</th><th>Total</th><th>Free</th><th>Usage</th><th>Health</th></tr></thead><tbody>$volumeRows</tbody></table>"
      } else {
        "<p style='color:#666;font-style:italic;'>No volume data available.</p>"
      })
    </div>
  </section>

</main>
<footer>
  Generated by D.W.A.R.F. — Part of the Technician Toolkit &nbsp;|&nbsp; $reportTimestamp
</footer>
</body>
</html>
"@

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
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  D.W.A.R.F. ASSESSMENT COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
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
Write-Host ("  " + ("═" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
