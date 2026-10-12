# ritual.ps1 - R.I.T.U.A.L. — Runs Integrated Tool Usage in Automation Loops
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
    R.I.T.U.A.L. — Runs Integrated Tool Usage in Automation Loops
    Workflow Orchestrator for the Technician Toolkit (PowerShell 5.1+)

.DESCRIPTION
    Runs an ordered sequence of toolkit scripts as a single named recipe.
    Ships with built-in recipes for common scenarios (Onboard, Retire,
    HealthCheck, TenantSweep) and accepts custom PSD1 recipe files. Each
    step is invoked with -Unattended so the whole recipe runs without
    prompts; step duration, exit code, and any new files written into the
    log directory are captured and rolled up into a single HTML summary.

    Built-in recipes:
      Onboard       -- New machine bring-up:
                       COVENANT -> BASILISK -> CONJURE -> WYRM -> AUSPEX -> PHOENIX
      Retire        -- Pre-reimage / pre-disposal:
                       PHYLACTERY -> EXHUME -> EMBALM -> CLEANSE
      HealthCheck   -- Quarterly machine review:
                       AUSPEX -> WARD -> FATHOM -> AUGUR -> VIGIL -> PHOENIX -> GRIFFIN
                       -> NECROPSY -> TORPOR -> SOLDER (audit only)
      SecuritySweep -- Endpoint security posture (read-only):
                       BASILISK -> TALON -> TOTEM -> GRIFFIN -> PHOENIX
      NetworkSweep  -- Endpoint network posture (read-only):
                       LEYLINE -> LANTERN -> WISP -> PORTAL -> LODESTAR
      TenantSweep   -- Cloud tenant posture (nine reports):
                       ZENITH -> ALMANAC -> ORBIT -> ECLIPSE -> ASTERISM -> CUMULUS
                       -> RAVEN -> HALO -> CARILLON

.USAGE
    PS C:\> .\ritual.ps1                                    # Interactive menu
    PS C:\> .\ritual.ps1 -Recipe HealthCheck                # Run a named recipe
    PS C:\> .\ritual.ps1 -RecipeFile .\custom.psd1          # Run a custom recipe file
    PS C:\> .\ritual.ps1 -Recipe Retire -ContinueOnError    # Ignore per-step failures

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [switch]$Transcript,
    [ValidateSet('Onboard','Retire','HealthCheck','SecuritySweep','NetworkSweep','TenantSweep')]
    [string]$Recipe = '',
    [ValidateScript({ [string]::IsNullOrWhiteSpace($_) -or (Test-Path -LiteralPath $_) })]
    [string]$RecipeFile = '',
    [switch]$ContinueOnError
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

function Write-TKError {
    <#
    .SYNOPSIS
        Appends a structured error to the toolkit's local error log,
        TK_Errors_<yyyyMM>.jsonl in the log directory. Nothing is sent off the
        machine: the toolkit keeps everything in the terminal or in files the
        technician can remove.
    .EXAMPLE
        Write-TKError -ScriptName 'basilisk' -Message $_.Exception.Message -Category 'Registry'
    #>
    param(
        [Parameter(Mandatory)][string]$ScriptName,
        [Parameter(Mandatory)][string]$Message,
        [string]$Category = 'General'
    )

    $entry = [PSCustomObject]@{
        Timestamp  = (Get-Date -Format 'o')
        Script     = $ScriptName
        Category   = $Category
        Message    = $Message
        Host       = $env:COMPUTERNAME
        User       = $env:USERNAME
    }

    # Append to monthly JSONL error log in the configured log directory
    try {
        $cfg     = Get-TKConfig
        $logRoot = if (-not [string]::IsNullOrWhiteSpace($cfg.LogDirectory) -and (Test-Path $cfg.LogDirectory)) {
            $cfg.LogDirectory
        } else {
            $PSScriptRoot
        }
        $logFile = Join-Path $logRoot "TK_Errors_$(Get-Date -Format 'yyyyMM').jsonl"
        $line    = $entry | ConvertTo-Json -Compress
        Add-Content -Path $logFile -Value $line -Encoding UTF8 -ErrorAction Stop
    }
    catch { <# never throw from a logging helper #> }
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
    <#
    .SYNOPSIS
        Stops the tool, with instructions, when the session is not elevated.
    .DESCRIPTION
        This used to relaunch the script elevated with the execution policy
        switched off. A script that re-launches itself as Administrator that
        way is exactly what malware does, and endpoint protection treats it
        so -- and the relaunch dropped every parameter the technician had
        passed. Elevation is now the technician's
        own, visible act: open PowerShell with "Run as administrator", then run
        the tool. The name is kept because the tools call it and the desktop
        app reads it as one of the two admin gates.
    #>
    param([Parameter(Mandatory)][string]$ScriptFile)
    if (-not (Test-IsAdmin)) {
        Write-Host "  $(Split-Path -Leaf $ScriptFile) needs Administrator rights." -ForegroundColor Yellow
        Write-Host "  Open PowerShell with 'Run as administrator' and run it again." -ForegroundColor Yellow
        exit 1
    }
}

# The one place the toolkit fetches its own code. GRIMOIRE and RITUAL call
# Save-TKVerifiedScript for a tool that is not beside them; it downloads the
# copy attached to the latest GitHub release and keeps it only if it carries a
# valid Authenticode signature from the toolkit's publisher. Anything else is
# deleted before it can run, so what runs is exactly what was released.
$script:TKReleaseBase = 'https://github.com/CursedTechnocrat/TechnicianToolkit/releases/latest/download'

$script:TKPublisher   = 'Open Source Developer John Bejarana'

function Test-TKScriptSignature {
    <#
    .SYNOPSIS
        $true only when a script carries a valid Authenticode signature from the
        expected publisher -- the toolkit's own unless -Publisher names another.
    .DESCRIPTION
        Valid means the file is unchanged since signing and the certificate
        chains to a trusted root; the publisher check means it is the expected
        certificate, matched by its common name so a renewed certificate (new
        thumbprint, same name) still passes. A quoted name with a comma in it
        ("Chocolatey Software, Inc.") is read whole. Any error -- including a
        platform without Get-AuthenticodeSignature -- is a no.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Publisher = $script:TKPublisher
    )

    try { $sig = Get-AuthenticodeSignature -FilePath $Path -ErrorAction Stop } catch { return $false }
    if ($sig.Status -ne 'Valid' -or -not $sig.SignerCertificate) { return $false }
    $cn = if ($sig.SignerCertificate.Subject -match '(?:^|,)\s*CN=(?:"(?<cn>[^"]+)"|(?<cn>[^,]+))') { $Matches['cn'].Trim() } else { '' }
    return ($cn -eq $Publisher)
}

function Save-TKVerifiedScript {
    <#
    .SYNOPSIS
        Downloads a toolkit script from the latest release into a folder, and
        keeps it only if its signature verifies. Returns the path, or $null.
    .DESCRIPTION
        Honours TK_DISABLE_DOWNLOAD. The file lands in a staging folder under its
        own name -- Windows picks the signature checker by extension, so it must
        stay a .ps1 -- and moves to its real place only after the check passes.
        A file that fails is deleted unrun.
    #>
    param(
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9_-]+\.ps1$')][string]$Name,
        [Parameter(Mandatory)][string]$Directory
    )

    if ($env:TK_DISABLE_DOWNLOAD -in @('1', 'true')) {
        Write-Fail "$Name is not beside this script, and TK_DISABLE_DOWNLOAD forbids fetching it."
        Write-Info "Deploy it next to this script from an approved release."
        return $null
    }

    $staging = Join-Path $Directory ('.tk-download-' + [guid]::NewGuid().ToString('N'))
    $file    = Join-Path $staging $Name
    Write-Step "$Name is not beside this script - downloading the signed copy from the latest release..."
    try {
        $null = New-Item -ItemType Directory -Path $staging -Force -ErrorAction Stop
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-RestMethod -Uri "$script:TKReleaseBase/$Name" -OutFile $file -ErrorAction Stop

        if (-not (Test-TKScriptSignature -Path $file)) {
            Write-Fail "$Name is not signed by $script:TKPublisher - discarded without running it."
            return $null
        }
        $target = Join-Path $Directory $Name
        Move-Item -LiteralPath $file -Destination $target -Force -ErrorAction Stop
        Write-Ok "Downloaded $Name (signature verified: $script:TKPublisher)."
        return $target
    }
    catch {
        Write-Fail "Could not download ${Name}: $($_.Exception.Message)"
        return $null
    }
    finally {
        Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
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

function Show-RitualBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  R.I.T.U.A.L. — Runs Integrated Tool Usage in Automation Loops" -ForegroundColor Cyan
    Write-Host "  Workflow Orchestrator for the Technician Toolkit  v6.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Recipes, execution, and rollup appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# BUILT-IN RECIPES
# ─────────────────────────────────────────────────────────────────────────────

# Each recipe is an ordered list of step hashtables. Fields:
#   Tool         -- file name of the script, resolved against $ScriptPath
#   Args         -- array of arguments forwarded to the script
#   StopOnError  -- override: $true means abort the recipe on step failure
#                   regardless of the run-wide -ContinueOnError flag
#   Label        -- pretty name shown in console / HTML
$script:BuiltInRecipes = @{
    'Onboard' = @{
        Name        = 'New Machine Onboard'
        Description = 'New-machine bring-up: identity, hardening, software, encryption, diagnostics, certificates.'
        Steps       = @(
            @{ Label = 'Machine onboarding';        Tool = 'covenant.ps1';   Args = @('-Unattended'); StopOnError = $true  }
            @{ Label = 'Security baseline';         Tool = 'basilisk.ps1';      Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Core software install';     Tool = 'conjure.ps1';    Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'BitLocker enable';          Tool = 'wyrm.ps1';     Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'System diagnostics';        Tool = 'auspex.ps1';     Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Certificate health';        Tool = 'phoenix.ps1';   Args = @('-Unattended'); StopOnError = $false }
        )
    }
    'Retire' = @{
        Name        = 'Pre-Reimage / Disposal'
        Description = 'Pre-reimage workflow: validate cloud data before local cleanup.'
        Steps       = @(
            @{ Label = 'OneDrive KFM validator';    Tool = 'phylactery.ps1';  Args = @('-Unattended'); StopOnError = $true  }
            @{ Label = 'Outlook PST discovery';     Tool = 'exhume.ps1';  Args = @('-Unattended'); StopOnError = $true  }
            @{ Label = 'Profile backup (ZIP)';      Tool = 'embalm.ps1'; Args = @('-Unattended'); StopOnError = $true  }
            @{ Label = 'Temp / cache cleanup';      Tool = 'cleanse.ps1'; Args = @('-Unattended'); StopOnError = $false }
        )
    }
    'HealthCheck' = @{
        Name        = 'Quarterly Machine Review'
        Description = 'Read-only sweep for a point-in-time machine health snapshot, including performance and Windows servicing health.'
        Steps       = @(
            @{ Label = 'System diagnostics';        Tool = 'auspex.ps1';    Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Account audit';             Tool = 'ward.ps1';      Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Disk space';                Tool = 'fathom.ps1'; Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Disk hardware (SMART)';     Tool = 'augur.ps1';     Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Services & tasks';          Tool = 'vigil.ps1';  Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Certificate health';        Tool = 'phoenix.ps1';  Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'AV / Defender health';      Tool = 'griffin.ps1';   Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Crash & reboot history';    Tool = 'necropsy.ps1';  Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Performance triage';        Tool = 'torpor.ps1';    Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Windows servicing health';  Tool = 'solder.ps1';    Args = @('-Unattended'); StopOnError = $false }
        )
    }
    'SecuritySweep' = @{
        Name        = 'Endpoint Security Posture'
        Description = 'Read-only security posture sweep: baseline, persistence, TPM, AV/Defender, certificates.'
        Steps       = @(
            @{ Label = 'Security baseline drift';   Tool = 'basilisk.ps1';     Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Persistence / autoruns';    Tool = 'talon.ps1';     Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'TPM health';                Tool = 'totem.ps1';     Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'AV / Defender health';      Tool = 'griffin.ps1';   Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Certificate health';        Tool = 'phoenix.ps1';  Args = @('-Unattended'); StopOnError = $false }
        )
    }
    'NetworkSweep' = @{
        Name        = 'Endpoint Network Posture'
        Description = 'Read-only network posture sweep: live diagnostics, LAN discovery, Wi-Fi profiles, VPN tunnels, domain trust.'
        Steps       = @(
            @{ Label = 'Network diagnostics';       Tool = 'leyline.ps1';   Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'LAN discovery';             Tool = 'lantern.ps1';   Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Wi-Fi profile audit';       Tool = 'wisp.ps1';    Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'VPN / Always-On audit';     Tool = 'portal.ps1';    Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Domain trust';              Tool = 'lodestar.ps1';      Args = @('-Unattended'); StopOnError = $false }
        )
    }
    'TenantSweep' = @{
        Name        = 'Cloud Tenant Posture'
        Description = 'Full tenant posture in one sign-in sequence: Azure, M365 licensing, Intune, Entra ID hygiene, Teams, SharePoint, Exchange Online, Conditional Access, Teams Phone call queues and auto attendants.'
        Steps       = @(
            @{ Label = 'Azure assessment';          Tool = 'zenith.ps1';  Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'M365 license audit';        Tool = 'almanac.ps1'; Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Intune / MDM compliance';   Tool = 'orbit.ps1';     Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Entra ID identity hygiene'; Tool = 'eclipse.ps1';    Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Microsoft Teams audit';     Tool = 'asterism.ps1';  Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'SharePoint Online audit';   Tool = 'cumulus.ps1';     Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Exchange Online security';  Tool = 'raven.ps1';     Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Conditional Access posture'; Tool = 'halo.ps1';  Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Teams Phone queues & attendants'; Tool = 'carillon.ps1'; Args = @('-Unattended'); StopOnError = $false }
        )
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# RECIPE RESOLVER
# ─────────────────────────────────────────────────────────────────────────────

function Resolve-Recipe {
    param(
        [string]$NamedRecipe,
        [string]$PathToFile
    )

    if ($PathToFile) {
        try {
            $data = Import-PowerShellDataFile -Path $PathToFile -ErrorAction Stop
        } catch {
            Write-Fail "Could not parse recipe file '$PathToFile': $($_.Exception.Message)"
            Write-TKError -ScriptName 'ritual' -Message "Recipe file parse failure: $($_.Exception.Message)" -Category 'Recipe Load'
            return $null
        }

        # A recipe file is just a hashtable with the same shape as the built-ins.
        if (-not $data.Steps -or $data.Steps.Count -eq 0) {
            Write-Fail "Recipe file '$PathToFile' has no Steps defined."
            return $null
        }

        if (-not $data.Name)        { $data.Name        = (Split-Path -Leaf $PathToFile) }
        if (-not $data.Description) { $data.Description = 'Custom recipe' }
        return $data
    }

    if ($NamedRecipe -and $script:BuiltInRecipes.ContainsKey($NamedRecipe)) {
        return $script:BuiltInRecipes[$NamedRecipe]
    }

    return $null
}

# ─────────────────────────────────────────────────────────────────────────────
# TOOL RESOLUTION
# ─────────────────────────────────────────────────────────────────────────────

# Return a usable path to the requested tool script. If it isn't co-located
# with RITUAL, fetch the signed copy from the latest release, so a lone
# ritual.ps1 can still run a full recipe -- but only a copy signed by the
# toolkit's publisher is kept (Save-TKVerifiedScript). Returns $null if the tool
# is neither present nor verifiably downloadable.
function Resolve-ToolFile {
    param([string]$ToolName)

    $local = Join-Path $ScriptPath $ToolName
    if (Test-Path $local) { return $local }
    return Save-TKVerifiedScript -Name $ToolName -Directory $ScriptPath
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

# Snapshot the log directory so we can diff it after a step runs and attribute
# any new files (HTML reports, CSV logs, etc.) to the step that produced them.
function Get-LogDirSnapshot {
    param([string]$LogRoot)
    if (-not (Test-Path $LogRoot)) { return @{} }
    $map = @{}
    Get-ChildItem -LiteralPath $LogRoot -File -ErrorAction SilentlyContinue |
        ForEach-Object { $map[$_.FullName] = $_.LastWriteTimeUtc.Ticks }
    return $map
}

function Get-NewLogFiles {
    param([hashtable]$Before, [string]$LogRoot)
    $after = Get-LogDirSnapshot -LogRoot $LogRoot
    $new = foreach ($path in $after.Keys) {
        if (-not $Before.ContainsKey($path) -or $Before[$path] -ne $after[$path]) {
            Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
        }
    }
    return @($new | Where-Object { $_ })
}

function Invoke-RecipeStep {
    param(
        [int]$StepNumber,
        [hashtable]$Step,
        [string]$LogRoot
    )

    $label = if ($Step.Label) { $Step.Label } else { $Step.Tool }

    Write-Section ("STEP {0} - {1}" -f $StepNumber, $label)
    Write-Info "Tool    : $($Step.Tool)"
    Write-Info "Args    : $($Step.Args -join ' ')"
    Write-Host ""

    $toolPath = Resolve-ToolFile -ToolName $Step.Tool
    if (-not $toolPath) {
        Write-Fail "Tool file not found and could not be downloaded: $($Step.Tool)"
        return [PSCustomObject]@{
            StepNumber = $StepNumber
            Label      = $label
            Tool       = $Step.Tool
            Status     = 'Missing'
            DurationSec = 0
            ExitCode   = $null
            NewFiles   = @()
            Error      = "Tool '$($Step.Tool)' not found in $ScriptPath and download from GitHub failed"
        }
    }

    $beforeSnap = Get-LogDirSnapshot -LogRoot $LogRoot
    $start      = Get-Date
    $errorMsg   = $null
    $status     = 'Succeeded'
    $exitCode   = 0

    try {
        # Use & with the script path so each step runs in its own variable scope.
        # The child script's -Unattended flag lives in $Step.Args; RITUAL itself is
        # always unattended toward children regardless of its own -Unattended state.
        & $toolPath @($Step.Args)
        if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) {
            $exitCode = $LASTEXITCODE
            $status   = 'Failed'
            $errorMsg = "Child script exited with code $exitCode"
        }
    } catch {
        $status   = 'Failed'
        $errorMsg = $_.Exception.Message
        Write-Fail "Step threw: $errorMsg"
        Write-TKError -ScriptName 'ritual' -Message "Step '$label' ($($Step.Tool)) failed: $errorMsg" -Category 'Recipe Step'
    }

    $duration = [math]::Round(((Get-Date) - $start).TotalSeconds, 1)
    $newFiles = Get-NewLogFiles -Before $beforeSnap -LogRoot $LogRoot

    if ($status -eq 'Succeeded') {
        Write-Ok ("Step finished in {0}s." -f $duration)
        if ($newFiles.Count -gt 0) {
            Write-Info ("Produced {0} log/report file(s) in the log directory." -f $newFiles.Count)
        }
    } else {
        Write-Fail ("Step failed after {0}s: {1}" -f $duration, $errorMsg)
    }

    Write-Host ""

    return [PSCustomObject]@{
        StepNumber  = $StepNumber
        Label       = $label
        Tool        = $Step.Tool
        Status      = $status
        DurationSec = $duration
        ExitCode    = $exitCode
        NewFiles    = @($newFiles)
        Error       = $errorMsg
        StopOnError = [bool]$Step.StopOnError
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML ROLLUP REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-RollupHtml {
    param(
        [hashtable]$RecipeData,
        [array]$Results,
        [datetime]$StartTime,
        [datetime]$EndTime
    )

    $tkCfg     = Get-TKConfig
    $orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $totalSec   = [math]::Round(($EndTime - $StartTime).TotalSeconds, 1)
    $succeeded  = @($Results | Where-Object { $_.Status -eq 'Succeeded' }).Count
    $failed     = @($Results | Where-Object { $_.Status -eq 'Failed' }).Count
    $skipped    = @($Results | Where-Object { $_.Status -eq 'Skipped' }).Count
    $missing    = @($Results | Where-Object { $_.Status -eq 'Missing' }).Count

    $overallClass = if ($failed -gt 0 -or $missing -gt 0) { 'err' }
                    elseif ($skipped -gt 0) { 'warn' }
                    else { 'ok' }
    $overallLabel = if ($failed -gt 0 -or $missing -gt 0) { 'RECIPE FAILED' }
                    elseif ($skipped -gt 0) { 'RECIPE PARTIAL' }
                    else { 'RECIPE SUCCEEDED' }

    $stepRows = [System.Text.StringBuilder]::new()
    foreach ($r in $Results) {
        $badge = switch ($r.Status) {
            'Succeeded' { "<span class='tk-badge-ok'>Succeeded</span>" }
            'Failed'    { "<span class='tk-badge-err'>Failed</span>" }
            'Skipped'   { "<span class='tk-badge-warn'>Skipped</span>" }
            'Missing'   { "<span class='tk-badge-err'>Missing</span>" }
            default     { "<span class='tk-badge-info'>$(EscHtml $r.Status)</span>" }
        }

        $artifactList = if ($r.NewFiles.Count -eq 0) {
            "<span class='tk-badge-info'>(none)</span>"
        } else {
            ($r.NewFiles | ForEach-Object {
                # Link out to the file with a relative file:// URI so the rollup
                # HTML can open each child report in the browser.
                $uri = [System.Uri]::new($_.FullName).AbsoluteUri
                "<a href='$uri'>$(EscHtml $_.Name)</a>"
            }) -join '<br>'
        }

        $errCell = if ($r.Error) { "<span class='tk-badge-err'>$(EscHtml $r.Error)</span>" } else { '' }

        [void]$stepRows.Append("<tr><td>$($r.StepNumber)</td><td>$(EscHtml $r.Label)</td><td><code>$(EscHtml $r.Tool)</code></td><td>$badge</td><td>$($r.DurationSec)s</td><td>$artifactList</td><td>$errCell</td></tr>`n")
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'R.I.T.U.A.L. Recipe Execution Report' `
        -ScriptName 'R.I.T.U.A.L.' `
        -Subtitle   "${orgPrefix}$(EscHtml $RecipeData.Name)" `
        -MetaItems  ([ordered]@{
            'Machine'     = $env:COMPUTERNAME
            'Run As'      = "$env:USERDOMAIN\$env:USERNAME"
            'Recipe'      = $RecipeData.Name
            'Started'     = $StartTime.ToString('yyyy-MM-dd HH:mm:ss')
            'Finished'    = $EndTime.ToString('yyyy-MM-dd HH:mm:ss')
            'Total'       = "${totalSec}s"
        }) `
        -NavItems   @('Overall', 'Steps')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'R.I.T.U.A.L. v6.0'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $overallClass"><div class="tk-summary-num">$overallLabel</div><div class="tk-summary-lbl">Outcome</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Results.Count)</div><div class="tk-summary-lbl">Total Steps</div></div>
    <div class="tk-summary-card ok"><div class="tk-summary-num">$succeeded</div><div class="tk-summary-lbl">Succeeded</div></div>
    <div class="tk-summary-card $(if ($failed -gt 0) { 'err' } else { 'ok' })"><div class="tk-summary-num">$failed</div><div class="tk-summary-lbl">Failed</div></div>
    <div class="tk-summary-card $(if ($skipped -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$skipped</div><div class="tk-summary-lbl">Skipped</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">${totalSec}s</div><div class="tk-summary-lbl">Total Duration</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Recipe -- $(EscHtml $RecipeData.Name)</span></div>
    <div class="tk-card"><div class="tk-info-box">$(EscHtml $RecipeData.Description)</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Step Results</span><span class="tk-section-num">$($Results.Count) step(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>#</th><th>Label</th><th>Tool</th><th>Status</th><th>Duration</th><th>Artifacts</th><th>Error</th></tr></thead>
        <tbody>$($stepRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

function Export-RollupHtml {
    param([hashtable]$RecipeData, [array]$Results, [datetime]$StartTime, [datetime]$EndTime)

    $html      = Build-RollupHtml -RecipeData $RecipeData -Results $Results -StartTime $StartTime -EndTime $EndTime
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "RITUAL_${timestamp}.html"

    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended -Label 'Rollup report'
    } catch {
        Write-Fail "Could not save rollup report: $($_.Exception.Message)"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# RECIPE DRIVER
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-Recipe {
    param([hashtable]$RecipeData)

    Show-RitualBanner
    Write-Section ("RECIPE: " + $RecipeData.Name)
    Write-Info $RecipeData.Description
    Write-Host ""

    $logRoot  = Resolve-LogDirectory -FallbackPath $ScriptPath
    $results  = [System.Collections.Generic.List[object]]::new()
    $abort    = $false
    $stepIdx  = 0
    $start    = Get-Date

    foreach ($step in $RecipeData.Steps) {
        $stepIdx++

        if ($abort) {
            Write-Section ("STEP {0} - {1}" -f $stepIdx, $step.Label)
            Write-Warn "Skipped (prior step aborted the recipe)."
            $results.Add([PSCustomObject]@{
                StepNumber  = $stepIdx
                Label       = $step.Label
                Tool        = $step.Tool
                Status      = 'Skipped'
                DurationSec = 0
                ExitCode    = $null
                NewFiles    = @()
                Error       = 'Prior step aborted the recipe'
                StopOnError = [bool]$step.StopOnError
            })
            continue
        }

        $result = Invoke-RecipeStep -StepNumber $stepIdx -Step $step -LogRoot $logRoot
        $results.Add($result) | Out-Null

        if ($result.Status -ne 'Succeeded') {
            $shouldAbort = if ($result.StopOnError) { $true } else { -not $ContinueOnError }
            if ($shouldAbort) {
                Write-Fail "Aborting remaining steps (step failed and continue-on-error is off)."
                Write-Host ""
                $abort = $true
            }
        }
    }

    $end = Get-Date

    # Final console summary
    $total   = $results.Count
    $ok      = @($results | Where-Object { $_.Status -eq 'Succeeded' }).Count
    $bad     = @($results | Where-Object { $_.Status -eq 'Failed' -or $_.Status -eq 'Missing' }).Count
    $skipped = @($results | Where-Object { $_.Status -eq 'Skipped' }).Count

    Write-Section "RECIPE SUMMARY"
    $color = if ($bad -gt 0) { $C.Error } elseif ($skipped -gt 0) { $C.Warning } else { $C.Success }
    Write-Host ("  Recipe   : {0}" -f $RecipeData.Name) -ForegroundColor $C.Info
    Write-Host ("  Total    : {0} steps" -f $total) -ForegroundColor $C.Info
    Write-Host ("  Succeeded: {0}" -f $ok)      -ForegroundColor $C.Success
    Write-Host ("  Failed   : {0}" -f $bad)     -ForegroundColor $(if ($bad -gt 0) { $C.Error } else { $C.Info })
    Write-Host ("  Skipped  : {0}" -f $skipped) -ForegroundColor $(if ($skipped -gt 0) { $C.Warning } else { $C.Info })
    Write-Host ("  Duration : {0}s" -f [math]::Round(($end - $start).TotalSeconds, 1)) -ForegroundColor $C.Info
    Write-Host ""

    Export-RollupHtml -RecipeData $RecipeData -Results @($results) -StartTime $start -EndTime $end

    return (@($results | Where-Object { $_.Status -eq 'Failed' -or $_.Status -eq 'Missing' }).Count -eq 0)
}

# ─────────────────────────────────────────────────────────────────────────────
# MENU
# ─────────────────────────────────────────────────────────────────────────────

function Show-Menu {
    Show-RitualBanner
    Write-Host "  Built-in recipes:" -ForegroundColor $C.Header
    Write-Host ""
    $i = 0
    $menuMap = @{}
    foreach ($name in $script:BuiltInRecipes.Keys) {
        $i++
        $menuMap["$i"] = $name
        $r = $script:BuiltInRecipes[$name]
        Write-Host ("  [{0}]  {1}" -f $i, $r.Name) -ForegroundColor $C.Info
        Write-Host ("        $($r.Description)") -ForegroundColor $C.Info
        Write-Host ("        Steps: " + (($r.Steps | ForEach-Object { $_.Tool -replace '\.ps1$','' }) -join ' -> ')) -ForegroundColor $C.Info
        Write-Host ""
    }
    Write-Host "  [F]  Load a custom recipe file (.psd1)" -ForegroundColor $C.Info
    Write-Host "  [Q]  Quit" -ForegroundColor $C.Info
    Write-Host ""
    return $menuMap
}

# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended -or $Recipe -or $RecipeFile) {
    $chosen = Resolve-Recipe -NamedRecipe $Recipe -PathToFile $RecipeFile
    if (-not $chosen) {
        if ($Unattended -and -not $Recipe -and -not $RecipeFile) {
            Write-Fail "Unattended mode requires -Recipe or -RecipeFile."
        } else {
            Write-Fail "Could not resolve the requested recipe."
        }
        if ($Transcript) { Stop-TKTranscript }
        exit 1
    }
    $ok = Invoke-Recipe -RecipeData $chosen
    if ($Transcript) { Stop-TKTranscript }
    exit $(if ($ok) { 0 } else { 1 })
}

do {
    $menuMap = Show-Menu
    $choice  = (Read-Host "  Select recipe").Trim().ToUpper()

    if ($choice -eq 'Q') { break }

    if ($choice -eq 'F') {
        $path = Read-Host "  Path to .psd1 recipe file"
        if (-not (Test-Path -LiteralPath $path)) {
            Write-Warn "File not found: $path"
            Start-Sleep -Seconds 1
            continue
        }
        $chosen = Resolve-Recipe -PathToFile $path
    } elseif ($menuMap.ContainsKey($choice)) {
        $chosen = Resolve-Recipe -NamedRecipe $menuMap[$choice]
    } else {
        Write-Warn "Invalid selection."
        Start-Sleep -Milliseconds 800
        continue
    }

    if ($chosen) {
        Invoke-Recipe -RecipeData $chosen | Out-Null
        Read-Host "  Press Enter to return to menu"
    }

} while ($true)

if ($Transcript) { Stop-TKTranscript }
