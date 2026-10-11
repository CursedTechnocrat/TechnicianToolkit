# talon.ps1 - T.A.L.O.N. — Tracks Anomalies & Locates Otherwise-silent Nastiness
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
    T.A.L.O.N. — Tracks Anomalies & Locates Otherwise-silent Nastiness
    Persistence / Autoruns Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Sweeps the standard Windows persistence surfaces -- Run keys, Startup
    folders, Scheduled Tasks, Services, WMI event subscriptions, Image
    File Execution Options, and Winlogon hijack points -- and inventories
    every entry. Flags anything signed by a non-Microsoft publisher, any
    unsigned binary, and any target file that is missing from disk. Dark-
    themed HTML report with a summary by category and a full detail table
    per surface.

    Does NOT attempt to judge malice (every dev machine is full of
    legitimate persistence entries). The tool's job is visibility: give
    the technician one page that answers "what runs on this machine
    without me asking it to?"

.USAGE
    PS C:\> .\talon.ps1                    # Interactive run
    PS C:\> .\talon.ps1 -Unattended        # Silent: export HTML and exit

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [switch]$Transcript
)

# ===========================
# SHARED HELPERS
# ===========================
#region TK-SHARED
# Generated from shared/TechnicianToolkit.Shared.ps1 by build/Sync-TKShared.ps1. Do not edit
# this block: edit the source and re-run the script. The Pester suite fails on a stale block.

# Forcing UTF-8 keeps the box-drawing banners intact when a console is attached.
# When one is not — a GUI host, a WinRM runspace, a scheduled task — the setter
# throws "The handle is invalid", so this must never be fatal. Every tool runs
# this as the first line of its generated block.
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host ("  " + ("-" * 62)) -ForegroundColor Cyan
    Write-Host ""
}

function Write-Step { param([string]$Msg) Write-Host ("  [*] {0}" -f $Msg) -ForegroundColor Magenta }

function Write-Ok   { param([string]$Msg) Write-Host ("  [+] {0}" -f $Msg) -ForegroundColor Green   }

function Write-Warn { param([string]$Msg) Write-Host ("  [!] {0}" -f $Msg) -ForegroundColor Yellow  }

function Write-Fail { param([string]$Msg) Write-Host ("  [-] {0}" -f $Msg) -ForegroundColor Red     }

function Write-Info { param([string]$Msg) Write-Host ("      {0}" -f $Msg) -ForegroundColor Gray    }

function Show-TKReportResult {
    <#
    .SYNOPSIS
        Prints a consistent "report saved" line and offers to open the file.
    .DESCRIPTION
        Standardises the end-of-run experience for tools that emit a report.
        Prints the resolved path once, and — unless -Unattended is set — prompts
        to open it in the default handler. Safe to call when the file is missing:
        it prints a warning instead of prompting.
    .PARAMETER Path        Path to the generated report file.
    .PARAMETER Unattended  Suppress the interactive "open now?" prompt.
    .PARAMETER Label       Noun shown in the saved line (default 'Report').
    .EXAMPLE
        Show-TKReportResult -Path $reportPath -Unattended:$Unattended
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Unattended,
        [string]$Label = 'Report'
    )

    if (-not (Test-Path $Path)) {
        Write-Warn "$Label was not created: $Path"
        return
    }

    $full = (Resolve-Path $Path -ErrorAction SilentlyContinue).Path
    if (-not $full) { $full = $Path }

    Write-Ok "$Label saved:"
    Write-Info $full

    if ($Unattended) { return }

    Write-Host ""
    Write-Host -NoNewline ("  Open the {0} now? (Y/N) " -f $Label.ToLower()) -ForegroundColor Cyan
    $answer = Read-Host
    if ($answer -match '^(y|yes)$') {
        try   { Invoke-Item -Path $full -ErrorAction Stop }
        catch { Write-Warn "Could not open the file: $($_.Exception.Message)" }
    }
}

function EscHtml {
    param([string]$s)
    if (-not $s) { return '' }
    $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

function Get-TKHtmlCss {
    return @'
<style>
:root {
  --tk-bg:          #0a0e14;
  --tk-surface:     #111820;
  --tk-surface2:    #162030;
  --tk-border:      #1e2d3d;
  --tk-cyan:        #00e5cc;
  --tk-cyan-dim:    rgba(0,229,204,0.12);
  --tk-text:        #c8d4e0;
  --tk-text-dim:    #637587;
  --tk-green:       #3fb950;
  --tk-green-dim:   rgba(63,185,80,0.12);
  --tk-yellow:      #e3b341;
  --tk-yellow-dim:  rgba(227,179,65,0.12);
  --tk-red:         #f85149;
  --tk-red-dim:     rgba(248,81,73,0.12);
  --tk-blue:        #58a6ff;
  --tk-blue-dim:    rgba(88,166,255,0.12);
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--tk-bg);color:var(--tk-text);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',system-ui,sans-serif;font-size:14px;line-height:1.6}

/* nav */
.tk-nav{background:#0d1219;border-bottom:1px solid var(--tk-border);padding:0 32px;height:44px;display:flex;align-items:center;overflow-x:auto;white-space:nowrap;font-family:'Consolas','Courier New',monospace;font-size:11px;letter-spacing:.05em;color:var(--tk-text-dim);gap:0}
.tk-nav a{padding:0 16px;height:44px;display:inline-flex;align-items:center;text-decoration:none;color:var(--tk-text-dim);border-bottom:2px solid transparent;gap:6px}
.tk-nav a:hover{color:var(--tk-text)}
.tk-nav-num{color:var(--tk-cyan)}

/* page header */
.tk-page-header{background:linear-gradient(180deg,#0e1520 0%,var(--tk-bg) 100%);border-bottom:1px solid var(--tk-border);padding:36px 48px 32px}
.tk-report-label{font-family:'Consolas','Courier New',monospace;font-size:11px;letter-spacing:.15em;text-transform:uppercase;color:var(--tk-cyan);margin-bottom:10px}
.tk-page-title{font-size:28px;font-weight:600;color:#e8f0f8;line-height:1.2;margin-bottom:6px}
.tk-page-subtitle{font-size:13px;color:var(--tk-text-dim);margin-bottom:20px}
.tk-meta-bar{display:flex;gap:32px;flex-wrap:wrap;margin-top:18px}
.tk-meta-label{font-family:'Consolas','Courier New',monospace;font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--tk-text-dim);margin-bottom:3px}
.tk-meta-value{font-size:14px;font-weight:600;color:var(--tk-text)}

/* main */
.tk-main{padding:40px 48px;max-width:1280px}

/* section */
.tk-section{margin-bottom:48px}
.tk-section-tag{font-family:'Consolas','Courier New',monospace;font-size:10px;letter-spacing:.15em;text-transform:uppercase;color:var(--tk-cyan);margin-bottom:8px}
.tk-section-title{font-size:20px;font-weight:600;color:#e0eaf4;margin-bottom:4px;display:flex;align-items:baseline;gap:10px}
.tk-section-num{font-family:'Consolas','Courier New',monospace;font-size:12px;color:var(--tk-cyan)}
.tk-section-subtitle{font-size:13px;color:var(--tk-text-dim);margin-bottom:16px}
.tk-divider{border:none;border-top:1px solid var(--tk-border);margin:12px 0 20px}

/* card */
.tk-card{background:var(--tk-surface);border:1px solid var(--tk-border);border-radius:8px;padding:20px 24px;margin-bottom:16px}
.tk-card-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:16px}
.tk-card-label{font-family:'Consolas','Courier New',monospace;font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:var(--tk-cyan)}

/* summary row */
.tk-summary-row{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:32px}
.tk-summary-card{background:var(--tk-surface);border:1px solid var(--tk-border);border-radius:8px;padding:18px 24px;min-width:130px;flex:1}
.tk-summary-num{font-size:28px;font-weight:700;color:var(--tk-text);line-height:1;margin-bottom:6px}
.tk-summary-lbl{font-family:'Consolas','Courier New',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--tk-text-dim)}
.tk-summary-card.ok   .tk-summary-num{color:var(--tk-green)}
.tk-summary-card.warn .tk-summary-num{color:var(--tk-yellow)}
.tk-summary-card.err  .tk-summary-num{color:var(--tk-red)}
.tk-summary-card.info .tk-summary-num{color:var(--tk-cyan)}

/* table */
.tk-table-wrap{overflow-x:auto}
table.tk-table{width:100%;border-collapse:collapse;font-size:13px}
table.tk-table th{font-family:'Consolas','Courier New',monospace;font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--tk-cyan);text-align:left;padding:10px 12px;border-bottom:1px solid var(--tk-border);font-weight:normal;white-space:nowrap}
table.tk-table td{padding:11px 12px;border-bottom:1px solid #162030;color:var(--tk-text);vertical-align:middle}
table.tk-table tr:last-child td{border-bottom:none}
table.tk-table tr:hover td{background:rgba(255,255,255,.02)}

/* badges */
/* The pill shape belongs to every colour class too: most reports write class='tk-badge-ok' alone. */
.tk-badge,.tk-badge-ok,.tk-badge-warn,.tk-badge-err,.tk-badge-info,.tk-badge-blue{display:inline-block;padding:2px 10px;border-radius:20px;font-family:'Consolas','Courier New',monospace;font-size:11px;font-weight:600;letter-spacing:.03em;white-space:nowrap}
.tk-badge-ok   {background:var(--tk-green-dim); color:var(--tk-green); border:1px solid rgba(63,185,80,.25)}
.tk-badge-warn {background:var(--tk-yellow-dim);color:var(--tk-yellow);border:1px solid rgba(227,179,65,.25)}
.tk-badge-err  {background:var(--tk-red-dim);   color:var(--tk-red);   border:1px solid rgba(248,81,73,.25)}
.tk-badge-info {background:var(--tk-cyan-dim);  color:var(--tk-cyan);  border:1px solid rgba(0,229,204,.25)}
.tk-badge-blue {background:var(--tk-blue-dim);  color:var(--tk-blue);  border:1px solid rgba(88,166,255,.25)}

/* info box */
.tk-info-box{background:var(--tk-surface);border-left:3px solid var(--tk-cyan);border-radius:0 6px 6px 0;padding:14px 18px;margin-top:12px;font-size:13px}
.tk-info-label{font-family:'Consolas','Courier New',monospace;font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--tk-cyan);margin-bottom:4px}

/* progress */
.tk-progress-wrap{background:#162030;border-radius:4px;height:6px;overflow:hidden;width:120px;display:inline-block;vertical-align:middle;margin-left:8px}
.tk-progress-bar{height:100%;border-radius:4px}
.tk-progress-bar.ok  {background:var(--tk-green)}
.tk-progress-bar.warn{background:var(--tk-yellow)}
.tk-progress-bar.err {background:var(--tk-red)}

/* mono / code */
code,.tk-mono{font-family:'Consolas','Courier New',monospace;font-size:12px;background:var(--tk-surface2);padding:1px 5px;border-radius:3px}

/* footer */
.tk-footer{border-top:1px solid var(--tk-border);padding:20px 48px;font-family:'Consolas','Courier New',monospace;font-size:11px;color:var(--tk-text-dim);letter-spacing:.05em;display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px}

/* responsive */
img,svg,video{max-width:100%;height:auto}
.tk-main,.tk-card,.tk-info-box,.tk-summary-card{max-width:100%}
.tk-card{overflow-x:auto}
table.tk-table{max-width:100%}
table.tk-table td,table.tk-table th{overflow-wrap:anywhere;word-break:break-word}
.tk-info-box,.tk-meta-value,.tk-page-title,.tk-page-subtitle,.tk-section-title,.tk-section-subtitle{overflow-wrap:anywhere}

@media (max-width:900px){
  .tk-page-header{padding:24px 20px 20px}
  .tk-page-title{font-size:22px}
  .tk-meta-bar{gap:18px}
  .tk-nav{padding:0 16px}
  .tk-nav a{padding:0 12px}
  .tk-main{padding:24px 20px}
  .tk-section{margin-bottom:32px}
  .tk-card{padding:16px 18px}
  .tk-footer{padding:16px 20px}
  table.tk-table th,table.tk-table td{padding:9px 8px}
}
@media (max-width:560px){
  body{font-size:13px}
  .tk-page-header{padding:18px 14px 16px}
  .tk-page-title{font-size:19px}
  .tk-page-subtitle{font-size:12px}
  .tk-meta-bar{gap:14px}
  .tk-nav{padding:0 10px;font-size:10px}
  .tk-nav a{padding:0 10px}
  .tk-main{padding:18px 14px}
  .tk-section-title{font-size:17px}
  .tk-card{padding:14px}
  .tk-summary-row{gap:10px}
  .tk-summary-card{min-width:0;flex:1 1 calc(50% - 5px);padding:14px 16px}
  .tk-summary-num{font-size:22px}
  .tk-card-header{flex-direction:column;align-items:flex-start;gap:8px}
  .tk-footer{padding:14px;flex-direction:column;align-items:flex-start;gap:4px}
  table.tk-table{font-size:12px}
  table.tk-table th,table.tk-table td{padding:8px 6px}
  .tk-progress-wrap{width:80px}
}
</style>
'@
}

function Get-TKHtmlHead {
    <#
    .SYNOPSIS
        Returns the opening HTML, head, and page-header markup for a toolkit report.
    .PARAMETER Title      Report title shown as the page heading.
    .PARAMETER ScriptName Acronym label (e.g. 'A.U.S.P.E.X.') used in the report-label tag.
    .PARAMETER Subtitle   Optional subtitle / machine name line beneath the title.
    .PARAMETER MetaItems  Ordered hashtable of label->value pairs shown in the metadata bar.
    .PARAMETER NavItems   Array of section-label strings for the sticky nav bar.
    #>
    param(
        [string]   $Title      = 'Technician Toolkit Report',
        [string]   $ScriptName = 'T.K.',
        [string]   $Subtitle   = '',
        [hashtable]$MetaItems  = @{},
        [string[]] $NavItems   = @()
    )

    $css = Get-TKHtmlCss

    $metaHtml = ''
    if ($MetaItems.Count -gt 0) {
        $parts = foreach ($k in $MetaItems.Keys) {
            "<div class='tk-meta-item'><div class='tk-meta-label'>$(EscHtml $k)</div><div class='tk-meta-value'>$(EscHtml $MetaItems[$k])</div></div>"
        }
        $metaHtml = "<div class='tk-meta-bar'>$($parts -join '')</div>"
    }

    $subtitleHtml = if ($Subtitle) { "<div class='tk-page-subtitle'>$(EscHtml $Subtitle)</div>" } else { '' }

    $navHtml = ''
    if ($NavItems.Count -gt 0) {
        $links = for ($i = 0; $i -lt $NavItems.Count; $i++) {
            $n = '{0:D2}' -f ($i + 1)
            "<a href='#s$n'><span class='tk-nav-num'>$n</span> &middot; $(EscHtml $NavItems[$i])</a>"
        }
        $navHtml = "<nav class='tk-nav'>$($links -join '')</nav>"
    }

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>$Title</title>
$css
</head>
<body>
<div class="tk-page-header">
  <div class="tk-report-label">$(EscHtml $ScriptName) REPORT</div>
  <div class="tk-page-title">$(EscHtml $Title)</div>
  $subtitleHtml
  $metaHtml
</div>
$navHtml
<div class="tk-main">
"@
}

function Get-TKHtmlFoot {
    <#
    .SYNOPSIS
        Returns the closing HTML markup for a toolkit report.
    .PARAMETER ScriptName  Shown in footer right (e.g. 'A.U.S.P.E.X. v1.0').
    #>
    param([string]$ScriptName = 'TechnicianToolkit')
    $ts       = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $hostName = $env:COMPUTERNAME
    return @"
</div>
<div class="tk-footer">
  <span>Generated $ts on $hostName</span>
  <span>$(EscHtml $ScriptName)</span>
</div>
</body>
</html>
"@
}

function Get-TKConfig {
    <#
    .SYNOPSIS
        Returns the toolkit configuration as a PSCustomObject.
        Reads config.json from the tool's own directory; missing keys are filled
        with empty-string defaults so callers never receive null.
    #>
    $configPath = Join-Path $PSScriptRoot 'config.json'

    $defaults = [PSCustomObject]@{
        OrgName       = ''
        LogDirectory  = ''
        TeamsWebhook  = ''
        Archive       = [PSCustomObject]@{ DefaultDestination   = '' }
        Revenant      = [PSCustomObject]@{ DefaultDestination   = '' }
        Covenant      = [PSCustomObject]@{ DefaultTimezone = ''; DefaultLocalAdminUser = '' }
        # DirectDownloads is an array of { Name; Url; Args; Sha256 } — installers
        # CONJURE fetches by URL because they have no winget/Chocolatey package.
        Conjure       = [PSCustomObject]@{ DirectDownloads = @() }
    }

    if (-not (Test-Path $configPath)) { return $defaults }

    try {
        $raw = Get-Content $configPath -Raw -ErrorAction Stop | ConvertFrom-Json

        # Legacy migration: v2.x used a 'Phantom' section; v3.0 renamed it to 'Revenant'.
        # Copy any populated 'Phantom' value forward so upgrading users don't lose their default destination.
        if ($raw.PSObject.Properties['Phantom'] -and
            $raw.Phantom -is [PSCustomObject] -and
            -not [string]::IsNullOrWhiteSpace($raw.Phantom.DefaultDestination)) {
            if ($raw.PSObject.Properties['Revenant'] -isnot [PSCustomObject] -or
                [string]::IsNullOrWhiteSpace($raw.Revenant.DefaultDestination)) {
                if (-not $raw.PSObject.Properties['Revenant']) {
                    $raw | Add-Member -NotePropertyName 'Revenant' -NotePropertyValue ([PSCustomObject]@{ DefaultDestination = '' }) -Force
                }
                $raw.Revenant.DefaultDestination = $raw.Phantom.DefaultDestination
            }
        }

        # Ensure top-level keys exist
        foreach ($key in ($defaults | Get-Member -MemberType NoteProperty).Name) {
            if ($null -eq $raw.$key) {
                $raw | Add-Member -NotePropertyName $key -NotePropertyValue $defaults.$key -Force
            }
        }
        # Ensure nested keys exist
        foreach ($section in @('Archive','Revenant','Covenant','Conjure')) {
            if ($raw.$section -isnot [PSCustomObject]) {
                $raw | Add-Member -NotePropertyName $section -NotePropertyValue $defaults.$section -Force
            } else {
                foreach ($key in ($defaults.$section | Get-Member -MemberType NoteProperty).Name) {
                    if ($null -eq $raw.$section.$key) {
                        # Carry the default's own type: filling an absent array key
                        # with '' would hand the caller a string to iterate.
                        $fallback = if ($defaults.$section.$key -is [Array]) { @() } else { '' }
                        $raw.$section | Add-Member -NotePropertyName $key -NotePropertyValue $fallback -Force
                    }
                }
            }
        }
        return $raw
    }
    catch { return $defaults }
}

function Start-TKTranscript {
    <#
    .SYNOPSIS
        Starts a timestamped PowerShell transcript in the configured log directory.
        Call once near the top of a script when -Transcript is active.
    #>
    param([Parameter(Mandatory)][string]$LogRoot)
    $path = Join-Path $LogRoot "TK_Transcript_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    try {
        Start-Transcript -Path $path -Append -ErrorAction Stop | Out-Null
        Write-Host "  [*] Transcript: $path" -ForegroundColor Gray
    }
    catch {
        Write-Warn "Could not start transcript: $_"
    }
}

function Stop-TKTranscript {
    <#
    .SYNOPSIS
        Stops an active PowerShell transcript started by Start-TKTranscript.
    #>
    try { Stop-Transcript -ErrorAction Stop | Out-Null } catch {}
}

function Resolve-LogDirectory {
    <#
    .SYNOPSIS
        Returns the configured LogDirectory, or the supplied fallback path if
        LogDirectory is not set. Creates the directory if it does not exist.
    #>
    param([Parameter(Mandatory)][string]$FallbackPath)

    $cfg = Get-TKConfig
    if (-not [string]::IsNullOrWhiteSpace($cfg.LogDirectory)) {
        if (-not (Test-Path $cfg.LogDirectory)) {
            $null = New-Item -ItemType Directory -Path $cfg.LogDirectory -Force -ErrorAction SilentlyContinue
        }
        return $cfg.LogDirectory
    }
    return $FallbackPath
}

function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-AdminElevation {
    param([Parameter(Mandatory)][string]$ScriptFile)
    if (-not (Test-IsAdmin)) {
        Write-Host "  INFO: Restarting with administrator privileges..." -ForegroundColor Yellow
        $PSExe = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        Start-Process -FilePath $PSExe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptFile`"" -Verb RunAs
        exit
    }
}
#endregion TK-SHARED
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

function Show-TalonBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  T.A.L.O.N. — Tracks Anomalies & Locates Otherwise-silent Nastiness" -ForegroundColor Cyan
    Write-Host "  Persistence / Autoruns Audit Tool  v6.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Collectors, verdict, and report builders appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# ENTRY ENRICHMENT — signature, target existence
# ─────────────────────────────────────────────────────────────────────────────

# Each collector emits raw rows with Category / Source / Name / Command;
# this helper enriches every row with path existence and signature status so
# every surface is inspected uniformly.
function Get-EntryEnrichment {
    param([string]$CommandLine)

    $result = [PSCustomObject]@{
        TargetPath   = $null
        TargetExists = $false
        Signer       = $null
        SignatureStatus = 'Unknown'
        IsMicrosoft  = $false
    }

    if (-not $CommandLine) { return $result }

    # Pull the first token out of the command line. Handles both bare paths
    # ("C:\foo\bar.exe -arg") and quoted paths (""C:\foo\bar.exe" -arg").
    $path = $null
    $trim = $CommandLine.Trim()
    if ($trim.StartsWith('"')) {
        $close = $trim.IndexOf('"', 1)
        if ($close -gt 0) { $path = $trim.Substring(1, $close - 1) }
    } else {
        $path = ($trim -split '\s+', 2)[0]
    }

    if (-not $path) { return $result }

    # Resolve environment variables in the path (e.g. %SystemRoot%).
    try {
        $path = [Environment]::ExpandEnvironmentVariables($path)
    } catch {
        # Malformed environment references fall through and leave the path alone.
    }

    $result.TargetPath = $path

    if (Test-Path -LiteralPath $path -ErrorAction SilentlyContinue) {
        $result.TargetExists = $true
        try {
            $sig = Get-AuthenticodeSignature -FilePath $path -ErrorAction Stop
            $result.SignatureStatus = $sig.Status.ToString()
            if ($sig.SignerCertificate) {
                $result.Signer      = $sig.SignerCertificate.Subject
                $result.IsMicrosoft = $sig.SignerCertificate.Subject -match 'CN=Microsoft'
            }
        } catch {
            $result.SignatureStatus = 'Error'
        }
    }

    return $result
}

# Uniform row constructor so every collector emits the same shape.
function New-PersistenceRow {
    param(
        [string]$Category,
        [string]$Source,
        [string]$Name,
        [string]$Command,
        [hashtable]$Extra = @{}
    )

    $enr = Get-EntryEnrichment -CommandLine $Command

    $row = [PSCustomObject]@{
        Category        = $Category
        Source          = $Source
        Name            = $Name
        Command         = $Command
        TargetPath      = $enr.TargetPath
        TargetExists    = $enr.TargetExists
        SignatureStatus = $enr.SignatureStatus
        Signer          = $enr.Signer
        IsMicrosoft     = $enr.IsMicrosoft
    }

    foreach ($k in $Extra.Keys) {
        $row | Add-Member -NotePropertyName $k -NotePropertyValue $Extra[$k] -Force
    }

    return $row
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTOR 1 — RUN / RUNONCE KEYS
# ─────────────────────────────────────────────────────────────────────────────

function Get-RunKeyEntries {
    $keys = @(
        @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run';        Source = 'HKCU\Run' }
        @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce';    Source = 'HKCU\RunOnce' }
        @{ Path = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run';        Source = 'HKLM\Run' }
        @{ Path = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce';    Source = 'HKLM\RunOnce' }
        @{ Path = 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run';     Source = 'HKLM\Wow6432\Run' }
        @{ Path = 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\RunOnce'; Source = 'HKLM\Wow6432\RunOnce' }
    )

    $rows = foreach ($k in $keys) {
        if (-not (Test-Path $k.Path)) { continue }
        $props = Get-ItemProperty -Path $k.Path -ErrorAction SilentlyContinue
        if (-not $props) { continue }
        foreach ($n in $props.PSObject.Properties.Name) {
            if ($n -in @('PSPath','PSParentPath','PSChildName','PSProvider','PSDrive')) { continue }
            New-PersistenceRow -Category 'Run Key' -Source $k.Source -Name $n -Command ($props.$n)
        }
    }

    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTOR 2 — STARTUP FOLDERS (per-user + all-users)
# ─────────────────────────────────────────────────────────────────────────────

function Get-StartupFolderEntries {
    $folders = @(
        @{ Path = [Environment]::GetFolderPath('Startup');        Source = 'User Startup' }
        @{ Path = [Environment]::GetFolderPath('CommonStartup');  Source = 'All Users Startup' }
    )

    $rows = foreach ($f in $folders) {
        if (-not $f.Path -or -not (Test-Path $f.Path)) { continue }
        foreach ($item in (Get-ChildItem -LiteralPath $f.Path -File -Force -ErrorAction SilentlyContinue)) {
            # Shortcut files: resolve their target and inspect it instead of the .lnk itself.
            $target = $item.FullName
            if ($item.Extension -eq '.lnk') {
                try {
                    $shell = New-Object -ComObject WScript.Shell
                    $sc    = $shell.CreateShortcut($item.FullName)
                    if ($sc.TargetPath) { $target = $sc.TargetPath }
                } catch {
                    # COM failures leave $target as the .lnk path; the report still lists the shortcut.
                }
            }
            New-PersistenceRow -Category 'Startup Folder' -Source $f.Source -Name $item.Name -Command $target
        }
    }

    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTOR 3 — SERVICES (non-Microsoft, auto-start)
# ─────────────────────────────────────────────────────────────────────────────

function Get-ServiceEntries {
    $svcs = Get-CimInstance -ClassName Win32_Service -ErrorAction SilentlyContinue |
        Where-Object { $_.StartMode -match 'Auto' }

    $rows = foreach ($s in $svcs) {
        # Skip kernel / filesystem / plug-and-play drivers -- they show up in Win32_Service but
        # aren't "persistence" in the autoruns sense. Only include real services with a PathName.
        if (-not $s.PathName) { continue }

        $row = New-PersistenceRow -Category 'Service' -Source 'Win32_Service' -Name $s.Name -Command $s.PathName -Extra @{
            DisplayName = $s.DisplayName
            StartMode   = $s.StartMode
            State       = $s.State
            Account     = $s.StartName
        }

        # Microsoft-signed services are legitimate by definition at this layer; leave them but
        # flag non-Microsoft so the UI can colour them distinctly.
        $row
    }

    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTOR 4 — SCHEDULED TASKS (non-Microsoft, auto-trigger)
# ─────────────────────────────────────────────────────────────────────────────

function Get-ScheduledTaskEntries {
    try {
        $tasks = Get-ScheduledTask -ErrorAction Stop
    } catch {
        return @()
    }

    $rows = foreach ($t in $tasks) {
        # Skip the Microsoft-shipped \Microsoft\ subtree except for the root-level tasks that
        # are common LOLBins targets (e.g. \Microsoft\Windows\Diagnosis).
        if ($t.TaskPath -like '\Microsoft\*') { continue }

        foreach ($a in $t.Actions) {
            $cmd = if ($a.Execute) { $a.Execute } else { '' }
            if ($a.Arguments) { $cmd = "$cmd $($a.Arguments)" }
            $row = New-PersistenceRow -Category 'Scheduled Task' -Source $t.TaskPath -Name $t.TaskName -Command $cmd -Extra @{
                State     = $t.State
                TaskPath  = $t.TaskPath
                Triggers  = (@($t.Triggers | ForEach-Object { $_.TriggerType }) -join ', ')
                RunAs     = $t.Principal.UserId
            }
            $row
        }
    }

    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTOR 5 — WMI EVENT SUBSCRIPTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Get-WmiPersistenceEntries {
    $rows = [System.Collections.Generic.List[object]]::new()

    # Permanent WMI event subscriptions are a high-signal persistence surface -- they're
    # rarely used by legitimate software and are a known APT trick. Surface every one.
    try {
        $filters   = Get-CimInstance -Namespace root\subscription -ClassName __EventFilter -ErrorAction SilentlyContinue
        $consumers = Get-CimInstance -Namespace root\subscription -ClassName __EventConsumer -ErrorAction SilentlyContinue
        $bindings  = Get-CimInstance -Namespace root\subscription -ClassName __FilterToConsumerBinding -ErrorAction SilentlyContinue
    } catch {
        return @()
    }

    foreach ($b in $bindings) {
        $cmd = 'N/A'
        if ($b.Consumer -and $b.Consumer -match 'Name=') {
            $cname = ($b.Consumer -replace '.*Name="([^"]+)".*','$1')
            $match = $consumers | Where-Object { $_.Name -eq $cname } | Select-Object -First 1
            if ($match) {
                if ($match.CommandLineTemplate) { $cmd = $match.CommandLineTemplate }
                elseif ($match.ScriptText)      { $cmd = "(ActiveScript): " + ($match.ScriptText -split "`n" | Select-Object -First 1) }
                elseif ($match.ExecutablePath)  { $cmd = $match.ExecutablePath }
            }
        }
        $rows.Add((New-PersistenceRow -Category 'WMI Subscription' -Source 'root\subscription' -Name $b.Filter -Command $cmd))
    }

    # If there are no bindings, still surface bare filters/consumers so a reviewer knows the
    # subscription surface was inspected (no silent skips).
    if ($bindings.Count -eq 0 -and ($filters.Count -gt 0 -or $consumers.Count -gt 0)) {
        foreach ($f in $filters) {
            $rows.Add((New-PersistenceRow -Category 'WMI Subscription' -Source 'root\subscription' -Name "Filter: $($f.Name)" -Command $f.Query))
        }
    }

    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTOR 6 — IMAGE FILE EXECUTION OPTIONS (IFEO) DEBUGGER HIJACKS
# ─────────────────────────────────────────────────────────────────────────────

function Get-IfeoEntries {
    # IFEO "Debugger" entries are a classic process-hijack primitive. Auditing every one is
    # cheap and the list is short on a clean machine.
    $ifeoRoot = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options'
    if (-not (Test-Path $ifeoRoot)) { return @() }

    $rows = foreach ($sub in Get-ChildItem $ifeoRoot -ErrorAction SilentlyContinue) {
        $props = Get-ItemProperty -Path $sub.PSPath -ErrorAction SilentlyContinue
        if ($props -and $props.Debugger) {
            New-PersistenceRow -Category 'IFEO Debugger' -Source 'HKLM\IFEO' -Name $sub.PSChildName -Command $props.Debugger
        }
    }

    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTOR 7 — WINLOGON HIJACK POINTS (Shell + Userinit)
# ─────────────────────────────────────────────────────────────────────────────

function Get-WinlogonEntries {
    $rows = [System.Collections.Generic.List[object]]::new()

    try {
        $wl = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -ErrorAction Stop
        if ($wl.Shell)    { $rows.Add((New-PersistenceRow -Category 'Winlogon' -Source 'HKLM\Winlogon' -Name 'Shell'    -Command $wl.Shell)) }
        if ($wl.Userinit) { $rows.Add((New-PersistenceRow -Category 'Winlogon' -Source 'HKLM\Winlogon' -Name 'Userinit' -Command $wl.Userinit)) }
    } catch {
        # Winlogon key always exists on Windows; failure here means access-denied in a test
        # environment. Continue silently rather than emit a fake-missing row.
    }

    # AppInit_DLLs is legacy (>= Windows 8 requires signed, disabled by default) but still
    # auditable.
    try {
        $init = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows' -ErrorAction Stop
        if ($init.AppInit_DLLs)  { $rows.Add((New-PersistenceRow -Category 'Winlogon' -Source 'HKLM\Windows' -Name 'AppInit_DLLs'  -Command $init.AppInit_DLLs)) }
        if ($init.LoadAppInit_DLLs) { $rows.Add((New-PersistenceRow -Category 'Winlogon' -Source 'HKLM\Windows' -Name 'LoadAppInit_DLLs' -Command "$($init.LoadAppInit_DLLs)")) }
    } catch {
        # Same as above -- this key always exists under normal OS builds.
    }

    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param([array]$Entries)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $total      = $Entries.Count
    $missing    = @($Entries | Where-Object { -not $_.TargetExists -and $_.TargetPath }).Count
    $unsigned   = @($Entries | Where-Object { $_.TargetExists -and $_.SignatureStatus -ne 'Valid' }).Count
    $thirdParty = @($Entries | Where-Object { $_.TargetExists -and $_.SignatureStatus -eq 'Valid' -and -not $_.IsMicrosoft }).Count
    $microsoft  = @($Entries | Where-Object { $_.IsMicrosoft }).Count

    $categories = $Entries | Group-Object Category | Sort-Object Name

    $summaryTitle = if ($missing -gt 0 -or $unsigned -gt 0) { 'REVIEW FINDINGS' } else { 'CLEAN LOOK' }
    $summaryClass = if ($missing -gt 0) { 'err' }
                    elseif ($unsigned -gt 0) { 'warn' }
                    else { 'ok' }

    # Per-category detail tables
    $catBlocks = [System.Text.StringBuilder]::new()
    foreach ($g in $categories) {
        $catRows = [System.Text.StringBuilder]::new()
        foreach ($r in ($g.Group | Sort-Object Source, Name)) {
            $sigBadge = switch ($r.SignatureStatus) {
                'Valid'      { if ($r.IsMicrosoft) { "<span class='tk-badge-ok'>MS signed</span>" } else { "<span class='tk-badge-info'>3rd-party signed</span>" } }
                'NotSigned'  { "<span class='tk-badge-warn'>Unsigned</span>" }
                'HashMismatch' { "<span class='tk-badge-err'>Tampered</span>" }
                'Error'      { "<span class='tk-badge-warn'>Sig error</span>" }
                'Unknown'    { "<span class='tk-badge-info'>n/a</span>" }
                default      { "<span class='tk-badge-warn'>$(EscHtml $r.SignatureStatus)</span>" }
            }
            $existBadge = if (-not $r.TargetPath) { "<span class='tk-badge-info'>n/a</span>" }
                          elseif ($r.TargetExists) { "<span class='tk-badge-ok'>Present</span>" }
                          else { "<span class='tk-badge-err'>Missing</span>" }
            [void]$catRows.Append("<tr><td>$(EscHtml $r.Source)</td><td>$(EscHtml $r.Name)</td><td><code>$(EscHtml $r.Command)</code></td><td>$existBadge</td><td>$sigBadge</td></tr>`n")
        }
        [void]$catBlocks.Append(@"

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">$(EscHtml $g.Name)</span><span class="tk-section-num">$($g.Count) entry/entries</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Source</th><th>Name</th><th>Command</th><th>Target</th><th>Signature</th></tr></thead>
        <tbody>$($catRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@)
    }

    $navItems = @('Summary') + ($categories | ForEach-Object { $_.Name })

    $htmlHead = Get-TKHtmlHead `
        -Title      'T.A.L.O.N. Persistence Audit' `
        -ScriptName 'T.A.L.O.N.' `
        -Subtitle   "${orgPrefix}Autoruns / Persistence Surface -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Run As'    = "$env:USERDOMAIN\$env:USERNAME"
            'Generated' = $reportDate
            'Entries'   = $total
        }) `
        -NavItems   $navItems

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'T.A.L.O.N. v6.0'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $summaryClass"><div class="tk-summary-num">$(EscHtml $summaryTitle)</div><div class="tk-summary-lbl">Outcome</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$total</div><div class="tk-summary-lbl">Total Entries</div></div>
    <div class="tk-summary-card ok"><div class="tk-summary-num">$microsoft</div><div class="tk-summary-lbl">MS-signed</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$thirdParty</div><div class="tk-summary-lbl">3rd-party Signed</div></div>
    <div class="tk-summary-card $(if ($unsigned -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$unsigned</div><div class="tk-summary-lbl">Unsigned / Sig Error</div></div>
    <div class="tk-summary-card $(if ($missing -gt 0) { 'err' } else { 'ok' })"><div class="tk-summary-num">$missing</div><div class="tk-summary-lbl">Missing Targets</div></div>
  </div>

$($catBlocks.ToString())

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-TalonBanner

Write-Section "COLLECTING PERSISTENCE SURFACES"
$entries = [System.Collections.Generic.List[object]]::new()

Write-Step "Run / RunOnce keys..."
Get-RunKeyEntries          | ForEach-Object { [void]$entries.Add($_) }
Write-Step "Startup folders..."
Get-StartupFolderEntries   | ForEach-Object { [void]$entries.Add($_) }
Write-Step "Services..."
Get-ServiceEntries         | ForEach-Object { [void]$entries.Add($_) }
Write-Step "Scheduled tasks..."
Get-ScheduledTaskEntries   | ForEach-Object { [void]$entries.Add($_) }
Write-Step "WMI event subscriptions..."
Get-WmiPersistenceEntries  | ForEach-Object { [void]$entries.Add($_) }
Write-Step "IFEO debugger hooks..."
Get-IfeoEntries            | ForEach-Object { [void]$entries.Add($_) }
Write-Step "Winlogon hijack points..."
Get-WinlogonEntries        | ForEach-Object { [void]$entries.Add($_) }
Write-Host ""

$entriesArr = @($entries)
Write-Ok "Collected $($entriesArr.Count) persistence entry/entries."
Write-Host ""

Write-Section "SUMMARY"
$byCat = $entriesArr | Group-Object Category | Sort-Object Name
foreach ($g in $byCat) {
    Write-Host ("  {0,-22} {1,6}" -f $g.Name, $g.Count) -ForegroundColor $C.Info
}
Write-Host ""

$missing  = @($entriesArr | Where-Object { -not $_.TargetExists -and $_.TargetPath }).Count
$unsigned = @($entriesArr | Where-Object { $_.TargetExists -and $_.SignatureStatus -ne 'Valid' }).Count

if ($missing -gt 0) {
    Write-Fail "$missing entry/entries point to a target that is NOT on disk (broken or removed)."
}
if ($unsigned -gt 0) {
    Write-Warn "$unsigned entry/entries have an unsigned target or signature error."
}
if ($missing -eq 0 -and $unsigned -eq 0) {
    Write-Ok "Every entry is signed and its target is present."
}
Write-Host ""

Write-Step "Generating HTML report..."
$html      = Build-HtmlReport -Entries $entriesArr
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "TALON_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
