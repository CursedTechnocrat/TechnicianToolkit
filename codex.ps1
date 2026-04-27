<#
.SYNOPSIS
    C.O.D.E.X. — Compiles Output Documents into an EXhibit
    Toolkit Report Index Builder for PowerShell 5.1+

.DESCRIPTION
    Walks the configured log directory, finds every TechnicianToolkit-
    generated HTML report (filename ending in _YYYYMMDD_HHMMSS.html),
    groups them by tool prefix, and emits a single dark-themed HTML
    rollup with relative links to each child report. Where R.I.T.U.A.L.
    composes a fresh recipe run, CODEX answers "what reports already
    exist on disk?" -- useful when a technician has run AUSPEX, PYRE,
    AUGUR, etc. ad-hoc throughout the week and wants one bound volume
    to attach to a ticket.

    Reports are sorted newest-first within each tool group. The summary
    row tallies report count, distinct tool count, total disk size, and
    the date range covered. Files that don't match the timestamp naming
    convention are skipped so random HTMLs in the log directory don't
    pollute the index.

.USAGE
    PS C:\> .\codex.ps1                         # Index every report in the log dir
    PS C:\> .\codex.ps1 -DaysBack 7             # Only the last week of reports
    PS C:\> .\codex.ps1 -LogDir 'D:\Reports'    # Override the scan target
    PS C:\> .\codex.ps1 -Unattended             # Silent: write rollup and exit

.NOTES
    Version : 1.0

#>

param(
    [switch]$Unattended,
    [switch]$Transcript,
    [int]$DaysBack = 0,
    [string]$LogDir
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
    Accent   = 'Blue'
}

function Show-CodexBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  C.O.D.E.X. — Compiles Output Documents into an EXhibit" -ForegroundColor Cyan
    Write-Host "  Toolkit Report Index Builder  v1.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# REPORT DISCOVERY
# ─────────────────────────────────────────────────────────────────────────────

# Walks the log directory for HTML files whose names end in the toolkit's
# canonical _YYYYMMDD_HHMMSS timestamp. Files outside that pattern are
# skipped on purpose so that browser-saved pages or hand-renamed copies
# don't pollute the rollup. CODEX's own output is excluded so successive
# runs don't index each other recursively.
function Get-ToolkitReportFiles {
    param(
        [Parameter(Mandatory)] [string]$LogDir,
        [int]$DaysBack = 0
    )

    if (-not (Test-Path $LogDir)) { return @() }

    # Trailing-slash normalization keeps the Substring math below correct
    # whether the caller passed `C:\Logs` or `C:\Logs\`.
    $logRoot = $LogDir.TrimEnd('\','/')
    $cutoff  = if ($DaysBack -gt 0) { (Get-Date).AddDays(-$DaysBack) } else { $null }
    $files   = @(Get-ChildItem -Path $logRoot -Filter '*.html' -File -Recurse -ErrorAction SilentlyContinue)

    $rows = foreach ($f in $files) {
        if ($f.Name -match '^CODEX_') { continue }

        # The toolkit-wide convention is `<label>_YYYYMMDD_HHMMSS.html`.
        # The label is the tool acronym, optionally followed by a variant
        # tag (PYRE_battery_report, CITADEL_StaleAccounts, etc.). Capture
        # all three groups up front -- $matches gets clobbered by any
        # later -match call below, so we can't rely on it surviving.
        if ($f.BaseName -notmatch '^(?<label>.+)_(?<date>\d{8})_(?<time>\d{6})$') { continue }
        $label   = $matches['label']
        $dateStr = $matches['date']
        $timeStr = $matches['time']

        $tool    = ($label -split '_', 2)[0].ToUpper()
        $variant = if ($label -match '^[^_]+_(?<v>.+)$') { $matches['v'] } else { '' }

        $ts = $null
        try {
            $ts = [datetime]::ParseExact("${dateStr}_${timeStr}", 'yyyyMMdd_HHmmss', $null)
        } catch {
            $ts = $f.LastWriteTime
        }

        if ($cutoff -and $ts -lt $cutoff) { continue }

        # Relative path from the log directory keeps the rollup portable
        # when the whole folder is zipped or moved to a ticket attachment.
        $rel = if ($f.FullName.StartsWith($logRoot, [StringComparison]::OrdinalIgnoreCase)) {
            $f.FullName.Substring($logRoot.Length).TrimStart('\','/')
        } else {
            $f.FullName
        }

        [PSCustomObject]@{
            FullPath  = $f.FullName
            RelPath   = $rel
            Name      = $f.Name
            Tool      = $tool
            Variant   = $variant
            Timestamp = $ts
            SizeBytes = $f.Length
        }
    }

    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-CodexHtml {
    param(
        [Parameter(Mandatory)] [array]$Reports,
        [Parameter(Mandatory)] [string]$LogDir,
        [int]$DaysBack = 0
    )

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $totalCount = $Reports.Count
    $totalSize  = ($Reports | Measure-Object -Property SizeBytes -Sum).Sum
    if ($null -eq $totalSize) { $totalSize = 0 }

    $tools = @($Reports | Group-Object Tool | Sort-Object Name)
    $toolCount = $tools.Count

    $oldest = if ($totalCount -gt 0) { ($Reports | Measure-Object -Property Timestamp -Minimum).Minimum } else { $null }
    $newest = if ($totalCount -gt 0) { ($Reports | Measure-Object -Property Timestamp -Maximum).Maximum } else { $null }

    $weekCutoff = (Get-Date).AddDays(-7)
    $lastWeek   = @($Reports | Where-Object { $_.Timestamp -ge $weekCutoff }).Count

    $rangeText = if ($DaysBack -gt 0) { "Last $DaysBack day(s)" } else { 'All time' }

    # Per-tool sections
    $toolSections = [System.Text.StringBuilder]::new()
    if ($tools.Count -eq 0) {
        [void]$toolSections.Append(@"

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">No reports found</span></div>
    <div class="tk-card"><div class="tk-info-box">Log directory: <span class="tk-mono">$(EscHtml $LogDir)</span><br>Either no reports exist in that path, or none match the <span class="tk-mono">_YYYYMMDD_HHMMSS.html</span> naming convention.$(if ($DaysBack -gt 0) { " The <span class=""tk-mono"">-DaysBack $DaysBack</span> filter may also be excluding older reports." } else { '' })</div></div>
  </div>
"@)
    }

    $navItems = @('Overview')

    foreach ($g in $tools) {
        $navItems += $g.Name
        $entries  = @($g.Group | Sort-Object Timestamp -Descending)
        $newestEntry = ($entries | Select-Object -First 1).Timestamp
        $oldestEntry = ($entries | Select-Object -Last 1).Timestamp
        $groupSize   = ($entries | Measure-Object -Property SizeBytes -Sum).Sum
        if ($null -eq $groupSize) { $groupSize = 0 }

        $rowsSb = [System.Text.StringBuilder]::new()
        foreach ($r in $entries) {
            $variant = if ([string]::IsNullOrWhiteSpace($r.Variant)) { '<span class="tk-badge-info">main</span>' } else { "<span class=""tk-badge-blue"">$(EscHtml $r.Variant)</span>" }
            $href    = EscHtml ($r.RelPath -replace '\\','/')
            $size    = Format-Bytes $r.SizeBytes
            $when    = $r.Timestamp.ToString('yyyy-MM-dd HH:mm:ss')
            [void]$rowsSb.Append("<tr><td class=""tk-mono"">$when</td><td>$variant</td><td><a href=""$href"">$(EscHtml $r.Name)</a></td><td>$size</td></tr>`n")
        }

        $latestStr = $newestEntry.ToString('yyyy-MM-dd HH:mm')
        $earliestStr = $oldestEntry.ToString('yyyy-MM-dd HH:mm')

        [void]$toolSections.Append(@"

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">$(EscHtml $g.Name)</span><span class="tk-section-num">$($entries.Count) report(s)</span></div>
    <div class="tk-card">
      <div class="tk-info-box">
        <span class="tk-info-label">Latest</span> $latestStr&nbsp;&nbsp;
        <span class="tk-info-label">Earliest</span> $earliestStr&nbsp;&nbsp;
        <span class="tk-info-label">Total size</span> $(Format-Bytes $groupSize)
      </div>
      <table class="tk-table">
        <thead><tr><th>Generated</th><th>Variant</th><th>File</th><th>Size</th></tr></thead>
        <tbody>$($rowsSb.ToString())</tbody>
      </table>
    </div>
  </div>
"@)
    }

    $newestCard = if ($newest) { $newest.ToString('yyyy-MM-dd') } else { 'n/a' }
    $oldestCard = if ($oldest) { $oldest.ToString('yyyy-MM-dd') } else { 'n/a' }

    $htmlHead = Get-TKHtmlHead `
        -Title      'C.O.D.E.X. Toolkit Report Index' `
        -ScriptName 'C.O.D.E.X.' `
        -Subtitle   "${orgPrefix}Toolkit Report Index -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'    = $machine
            'Log Dir'    = $LogDir
            'Generated'  = $reportDate
            'Range'      = $rangeText
            'Reports'    = $totalCount
            'Tools'      = $toolCount
        }) `
        -NavItems   $navItems

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'C.O.D.E.X. v1.0'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card info"><div class="tk-summary-num">$totalCount</div><div class="tk-summary-lbl">Reports</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$toolCount</div><div class="tk-summary-lbl">Tools</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$lastWeek</div><div class="tk-summary-lbl">Last 7 Days</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(Format-Bytes $totalSize)</div><div class="tk-summary-lbl">Total Size</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$newestCard</div><div class="tk-summary-lbl">Newest</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$oldestCard</div><div class="tk-summary-lbl">Oldest</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Overview</span><span class="tk-section-num">$rangeText</span></div>
    <div class="tk-card">
      <div class="tk-info-box">
        <span class="tk-info-label">Log directory</span> <span class="tk-mono">$(EscHtml $LogDir)</span><br>
        <span class="tk-info-label">Filter</span> $(EscHtml $rangeText)$(if ($DaysBack -gt 0) { " (cutoff $((Get-Date).AddDays(-$DaysBack).ToString('yyyy-MM-dd HH:mm')))" } else { '' })<br>
        <span class="tk-info-label">Naming pattern</span> <span class="tk-mono">&lt;TOOL&gt;_YYYYMMDD_HHMMSS.html</span> (files outside this pattern are skipped)
      </div>
    </div>
  </div>
$($toolSections.ToString())
"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-CodexBanner

$resolvedLogDir = if (-not [string]::IsNullOrWhiteSpace($LogDir)) {
    $LogDir
} else {
    Resolve-LogDirectory -FallbackPath $ScriptPath
}

Write-Section "REPORT DISCOVERY"
Write-Step "Scanning: $resolvedLogDir"
if ($DaysBack -gt 0) { Write-Info "  Filter: last $DaysBack day(s)" }

if (-not (Test-Path $resolvedLogDir)) {
    Write-Fail "Log directory does not exist: $resolvedLogDir"
    if (-not $Unattended) { Read-Host "  Press Enter to exit" }
    if ($Transcript) { Stop-TKTranscript }
    exit 1
}

$reports = Get-ToolkitReportFiles -LogDir $resolvedLogDir -DaysBack $DaysBack

if ($reports.Count -eq 0) {
    Write-Warn "No matching reports found."
} else {
    $byTool = $reports | Group-Object Tool | Sort-Object Name
    Write-Ok ("Found {0} report(s) across {1} tool(s)." -f $reports.Count, $byTool.Count)
    foreach ($g in $byTool) {
        $latest = ($g.Group | Sort-Object Timestamp -Descending | Select-Object -First 1).Timestamp.ToString('yyyy-MM-dd HH:mm')
        Write-Host ("    {0,-22} {1,3} report(s)   newest {2}" -f $g.Name, $g.Count, $latest) -ForegroundColor $C.Info
    }
}
Write-Host ""

Write-Step "Generating rollup HTML..."
$html      = Build-CodexHtml -Reports $reports -LogDir $resolvedLogDir -DaysBack $DaysBack
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path $resolvedLogDir "CODEX_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Write-Ok "Index saved: $outPath"
    if (-not $Unattended) {
        Write-Step "Opening in default browser..."
        Start-Process $outPath
    }
} catch {
    Write-Fail "Could not save index: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
