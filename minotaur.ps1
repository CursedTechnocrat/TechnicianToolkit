# minotaur.ps1 - M.I.N.O.T.A.U.R. — Maps Inheritance, NTFS Owners, Trustees & Access on UNC Roots
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
    M.I.N.O.T.A.U.R. — Maps Inheritance, NTFS Owners, Trustees & Access on UNC Roots
    File Share & NTFS Permissions Review Tool for PowerShell 5.1+

.DESCRIPTION
    Answers "who has access to this share?" on a file server, the way ARGUS
    answers it for the domain and WARD for the local machine:

      - Every non-administrative SMB share with its share permissions and the
        NTFS permissions on its root
      - A walk down each share (default two folder levels) recording every
        folder with explicit permissions or broken inheritance
      - Everyone / Authenticated Users / Users / Domain Users granted write
        where the share permissions let it through -- the combination that
        actually lets anyone in the organisation change or encrypt the data
      - Permissions granted straight to user accounts instead of groups,
        entries for deleted accounts (bare SIDs), and Deny entries

    Produces an HTML report and a CSV of every access-control entry scanned,
    ready for a customer access review. Read-only -- no permission is changed.

.USAGE
    PS C:\> .\minotaur.ps1                         # Interactive: all shares, two levels deep
    PS C:\> .\minotaur.ps1 -Unattended             # Silent: all shares + HTML + CSV
    PS C:\> .\minotaur.ps1 -Depth 4                # Walk four folder levels into each share
    PS C:\> .\minotaur.ps1 -Path 'D:\Data'         # Review one folder tree instead of the shares

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [ValidateScript({ [string]::IsNullOrWhiteSpace($_) -or (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]$Path,
    [ValidateRange(0, 10)]
    [int]$Depth = 2,
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

# Script-scoped note buffer. Lives for the duration of one tool run; the sync
# script copies it only into tools that use the note functions.
$script:TKNotes = New-Object System.Collections.ArrayList

function Add-TKNote {
    <#
    .SYNOPSIS
        Records a timestamped technician note for the current session.
    .DESCRIPTION
        Notes accumulate in script-scoped state for the life of one tool run.
        The tool saves them with Save-TKNoteReport at the end of the run.
    .PARAMETER Text        The note text.
    .PARAMETER Category    One of Info, Action, Warning, Issue, Resolution.
    .PARAMETER ScriptName  Optional originating tool name, shown in the report.
    .EXAMPLE
        Add-TKNote -Text 'Renamed computer to DESKTOP-01' -Category Action -ScriptName 'covenant'
    #>
    param(
        [Parameter(Mandatory)][string]$Text,
        [ValidateSet('Info','Action','Warning','Issue','Resolution')]
        [string]$Category   = 'Info',
        [string]$ScriptName = ''
    )
    [void]$script:TKNotes.Add([PSCustomObject]@{
        Timestamp = Get-Date
        Category  = $Category
        Script    = $ScriptName
        Text      = $Text
    })
}

function Export-TKNoteReport {
    <#
    .SYNOPSIS
        Writes the session's technician notes to a ticket-ready HTML report.
    .DESCRIPTION
        Renders the notes recorded with Add-TKNote into an HTML file built on
        the shared report template. The report includes a plain-text block
        sized for pasting straight into a ticketing system's comment field.
    .PARAMETER Path        Output .html file path.
    .PARAMETER Title       Report heading.
    .PARAMETER ScriptName  Acronym label for the report header/footer.
    .PARAMETER Ticket      Optional ticket / case reference shown in the metadata.
    .PARAMETER Technician  Name of the technician (defaults to the current user).
    .PARAMETER Summary     Optional free-text summary shown above the notes.
    .EXAMPLE
        Export-TKNoteReport -Path 'C:\Reports\notes.html' -Ticket 'INC012345'
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Title      = 'Technician Notes',
        [string]$ScriptName = 'T.K.',
        [string]$Ticket     = '',
        [string]$Technician = $env:USERNAME,
        [string]$Summary    = ''
    )

    $notes     = @($script:TKNotes)
    $generated = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $hostName  = $env:COMPUTERNAME

    $badgeClass = @{
        Info       = 'tk-badge-info'
        Action     = 'tk-badge-blue'
        Warning    = 'tk-badge-warn'
        Issue      = 'tk-badge-err'
        Resolution = 'tk-badge-ok'
    }

    $countAction     = @($notes | Where-Object { $_.Category -eq 'Action' }).Count
    $countIssue      = @($notes | Where-Object { $_.Category -eq 'Issue' }).Count
    $countResolution = @($notes | Where-Object { $_.Category -eq 'Resolution' }).Count

    $meta = [ordered]@{
        'Generated'  = $generated
        'Machine'    = $hostName
        'Technician' = $Technician
    }
    if (-not [string]::IsNullOrWhiteSpace($Ticket)) { $meta['Ticket'] = $Ticket }

    $html = Get-TKHtmlHead -Title $Title -ScriptName $ScriptName -Subtitle $hostName `
                -MetaItems $meta -NavItems @('Notes', 'Ticket Summary')

    $html += @"
<div class="tk-summary-row">
  <div class="tk-summary-card info"><div class="tk-summary-num">$($notes.Count)</div><div class="tk-summary-lbl">Total Notes</div></div>
  <div class="tk-summary-card"><div class="tk-summary-num">$countAction</div><div class="tk-summary-lbl">Actions</div></div>
  <div class="tk-summary-card err"><div class="tk-summary-num">$countIssue</div><div class="tk-summary-lbl">Issues</div></div>
  <div class="tk-summary-card ok"><div class="tk-summary-num">$countResolution</div><div class="tk-summary-lbl">Resolutions</div></div>
</div>
"@

    if (-not [string]::IsNullOrWhiteSpace($Summary)) {
        $html += "<div class=`"tk-info-box`"><div class=`"tk-info-label`">Summary</div>$(EscHtml $Summary)</div>"
    }

    $rows = ''
    if ($notes.Count -eq 0) {
        $rows = "<tr><td colspan='4'>No notes were recorded this session.</td></tr>"
    }
    else {
        foreach ($n in $notes) {
            $cls = $badgeClass[[string]$n.Category]
            if (-not $cls) { $cls = 'tk-badge-info' }
            $src = if ($n.Script) { EscHtml $n.Script } else { '&mdash;' }
            $rows += "<tr><td class='tk-mono'>$($n.Timestamp.ToString('HH:mm:ss'))</td>" +
                     "<td><span class='tk-badge $cls'>$(EscHtml ([string]$n.Category))</span></td>" +
                     "<td>$src</td><td>$(EscHtml $n.Text)</td></tr>"
        }
    }

    $html += @"
<div class="tk-section" id="s01">
  <div class="tk-section-title"><span class="tk-section-num">01</span> Notes</div>
  <div class="tk-card">
    <div class="tk-table-wrap">
    <table class="tk-table"><thead><tr><th>Time</th><th>Category</th><th>Source</th><th>Note</th></tr></thead>
    <tbody>$rows</tbody></table>
    </div>
  </div>
</div>
"@

    # Plain-text block — copy/paste straight into a ticket comment field.
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("TECHNICIAN NOTES - $Title")
    [void]$sb.AppendLine("Machine    : $hostName")
    [void]$sb.AppendLine("Technician : $Technician")
    [void]$sb.AppendLine("Generated  : $generated")
    if (-not [string]::IsNullOrWhiteSpace($Ticket)) { [void]$sb.AppendLine("Ticket     : $Ticket") }
    if (-not [string]::IsNullOrWhiteSpace($Summary)) {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("Summary: $Summary")
    }
    [void]$sb.AppendLine("")
    if ($notes.Count -eq 0) {
        [void]$sb.AppendLine("(no notes recorded)")
    }
    else {
        foreach ($n in $notes) {
            [void]$sb.AppendLine(("[{0}] [{1,-10}] {2}" -f `
                $n.Timestamp.ToString('HH:mm:ss'), ([string]$n.Category).ToUpper(), $n.Text))
        }
    }

    $html += @"
<div class="tk-section" id="s02">
  <div class="tk-section-title"><span class="tk-section-num">02</span> Ticket Summary</div>
  <div class="tk-section-subtitle">Copy the block below straight into the ticket comment field.</div>
  <div class="tk-card">
    <pre class="tk-mono" style="white-space:pre-wrap;display:block;padding:14px;line-height:1.5">$(EscHtml $sb.ToString())</pre>
  </div>
</div>
"@

    $html += Get-TKHtmlFoot -ScriptName $ScriptName

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path $dir)) {
        $null = New-Item -ItemType Directory -Path $dir -Force -ErrorAction SilentlyContinue
    }
    Set-Content -Path $Path -Value $html -Encoding UTF8
    return $Path
}

function Save-TKNoteReport {
    <#
    .SYNOPSIS
        Saves the run's technician notes beside the tool's reports, if it has any.
    .DESCRIPTION
        The note buffer lives only as long as the tool run, so every tool that
        records notes calls this once, at the end of the run, after its menu
        loop -- one file per run, holding every note the run recorded. Writes
        <Prefix>_Notes_<yyyyMMdd_HHmmss>.html, the name CODEX groups under the
        tool. Returns the path, or $null when there were no notes or the write
        failed; a failure is reported, never thrown, so it cannot end a run that
        has already done its work.
    .PARAMETER Prefix      Tool name used for the file label, e.g. 'CONDUIT'.
    .PARAMETER Directory   Folder to write into -- normally Resolve-LogDirectory.
    .PARAMETER ScriptName  Header label; defaults to the dotted form of Prefix.
    .EXAMPLE
        $null = Save-TKNoteReport -Prefix 'CONDUIT' -Directory (Resolve-LogDirectory -FallbackPath $ScriptPath)
    #>
    param(
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z]+$')][string]$Prefix,
        [Parameter(Mandatory)][string]$Directory,
        [string]$ScriptName = ''
    )

    if (@($script:TKNotes).Count -eq 0) { return $null }
    if (-not $ScriptName) { $ScriptName = ($Prefix.ToUpper().ToCharArray() -join '.') + '.' }

    $path = Join-Path $Directory ('{0}_Notes_{1}.html' -f $Prefix.ToUpper(), (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        $null = Export-TKNoteReport -Path $path -ScriptName $ScriptName -ErrorAction Stop
        Write-Ok "Technician notes saved:"
        Write-Info $path
        return $path
    }
    catch {
        Write-Warn "Could not save technician notes: $($_.Exception.Message)"
        return $null
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
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

# Principals that mean "effectively everyone". Domain Users is matched by its
# relative ID (-513) because the domain part of the SID differs per domain.
$BroadSids = [ordered]@{
    'S-1-1-0'      = 'Everyone'
    'S-1-5-7'      = 'Anonymous Logon'
    'S-1-5-11'     = 'Authenticated Users'
    'S-1-5-32-545' = 'BUILTIN\Users'
    'S-1-5-32-546' = 'BUILTIN\Guests'
}
$BroadDomainRids = @{ '513' = 'Domain Users'; '514' = 'Domain Guests' }

# Shares whose broad read access is by design.
$ExpectedBroadReadShares = @('NETLOGON', 'SYSVOL')

# A walk past this many folders stops, so one enormous share cannot run for hours.
$MaxFolders = 5000

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
# ─────────────────────────────────────────────────────────────────────────────

$MinotaurFindings = @{
    'BroadWriteAccess' = @{
        Severity = 'Error'
        Title    = 'Everyone-type group can write'
        Summary  = 'Both the share and NTFS permissions let Everyone, Authenticated Users, Users or Domain Users change files -- so any account in the organisation, including a compromised one running ransomware, can modify or encrypt them.'
        Remedy   = 'Grant write to the specific groups that need it and reduce the broad group to read (or remove it).'
    }
    'BroadReadAccess' = @{
        Severity = 'Warning'
        Title    = 'Everyone-type group can read'
        Summary  = 'Any account in the organisation can read this share. Fine for public material; not for HR, finance or client data.'
        Remedy   = 'Confirm the content is meant for everyone. If not, replace the broad group with the teams that need it.'
    }
    'DirectUserAce' = @{
        Severity = 'Warning'
        Title    = 'Permissions granted directly to user accounts'
        Summary  = 'Access given to named users instead of groups is invisible in group membership reviews and is left behind when people change roles.'
        Remedy   = 'Move each user into a group that holds the access, then remove the user entry.'
    }
    'OrphanedSid' = @{
        Severity = 'Warning'
        Title    = 'Permissions for deleted accounts'
        Summary  = 'These entries name a SID that no longer resolves -- the account or group was deleted but its access stayed.'
        Remedy   = 'Remove the orphaned entries (they grant nothing today, but hide what the ACL really says).'
    }
    'InheritanceBroken' = @{
        Severity = 'Info'
        Title    = 'Folders with inheritance disabled'
        Summary  = 'These folders do not take permissions from their parent, so a change at the share root will not reach them.'
        Remedy   = 'Expected for deliberately restricted folders. Review any you cannot explain.'
    }
    'DenyAce' = @{
        Severity = 'Info'
        Title    = 'Deny entries present'
        Summary  = 'Deny entries override allows and are a frequent cause of "I can''t open it" tickets.'
        Remedy   = 'Prefer removing the allow over adding a deny; keep denies documented.'
    }
    'FolderUnreadable' = @{
        Severity = 'Info'
        Title    = 'Folders M.I.N.O.T.A.U.R. could not read'
        Summary  = 'Access was denied even to the elevated session, so these folders were not reviewed.'
        Remedy   = 'Take ownership only if you are authorised to; otherwise review them with their owner.'
    }
    'ScanTruncated' = @{
        Severity = 'Warning'
        Title    = 'Folder walk stopped early'
        Summary  = 'The walk reached its folder limit ($MaxFolders), so deeper folders were not reviewed.'
        Remedy   = 'Rerun against the share with -Path and a smaller -Depth, or split the review by top-level folder.'
    }
    'NoShares' = @{
        Severity = 'Info'
        Title    = 'No file shares on this machine'
        Summary  = 'Only the administrative shares (C$, ADMIN$, IPC$) exist.'
        Remedy   = 'Run on the file server, or pass -Path to review a folder tree.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()

function Add-MinotaurFinding {
    param([Parameter(Mandatory)][string]$Code, [string]$Subject = '', [string]$Detail = '')
    $meta = $MinotaurFindings[$Code]
    if (-not $meta) { $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' } }
    [void]$Findings.Add([PSCustomObject]@{
        Code = $Code; Severity = $meta.Severity; Title = $meta.Title; Summary = $meta.Summary; Remedy = $meta.Remedy; Subject = $Subject; Detail = $Detail
    })
}

function Get-SeverityClass {
    param([string]$Severity)
    switch ($Severity) {
        'Error'   { return 'err'  }
        'Warning' { return 'warn' }
        default   { return 'info' }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-MinotaurBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ███╗   ███╗██╗███╗   ██╗ ██████╗ ████████╗ █████╗ ██╗   ██╗██████╗
  ████╗ ████║██║████╗  ██║██╔═══██╗╚══██╔══╝██╔══██╗██║   ██║██╔══██╗
  ██╔████╔██║██║██╔██╗ ██║██║   ██║   ██║   ███████║██║   ██║██████╔╝
  ██║╚██╔╝██║██║██║╚██╗██║██║   ██║   ██║   ██╔══██║██║   ██║██╔══██╗
  ██║ ╚═╝ ██║██║██║ ╚████║╚██████╔╝   ██║   ██║  ██║╚██████╔╝██║  ██║
  ╚═╝     ╚═╝╚═╝╚═╝  ╚═══╝ ╚═════╝    ╚═╝   ╚═╝  ╚═╝ ╚═════╝ ╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    M.I.N.O.T.A.U.R. — Maps Inheritance, NTFS Owners, Trustees & Access on UNC Roots" -ForegroundColor Cyan
    Write-Host "    File Share & NTFS Permissions Review Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PURE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function Get-NtfsRightsLevel {
    # Collapses a FileSystemRights mask to Full / Modify / Write / Read /
    # Special. Generic rights (GENERIC_ALL etc.) appear on inheritable ACEs
    # and are mapped to what they grant.
    param([long]$Value)
    $v = $Value
    if ($v -lt 0) { $v = $v + 4294967296 }
    if (($v -band 0x10000000) -ne 0) { return 'Full' }
    if (($v -band 0x1F01FF) -eq 0x1F01FF) { return 'Full' }
    if (($v -band 0x301BF) -eq 0x301BF) { return 'Modify' }
    if (($v -band 0x40000000) -ne 0 -or ($v -band 0x6) -ne 0) { return 'Write' }
    if (($v -band 0x80000000) -ne 0 -or ($v -band 0x20000000) -ne 0 -or ($v -band 0x1) -ne 0) { return 'Read' }
    return 'Special'
}

function Test-WriteLevel {
    param([string]$Level)
    return $Level -in @('Full', 'Modify', 'Write', 'Change')
}

function Get-SidCategory {
    # 'Broad' for everyone-type principals, 'WellKnown' for the other built-in
    # and service SIDs, 'Account' for anything a directory or SAM issued.
    param([string]$Sid)
    if ($BroadSids.Contains($Sid)) { return 'Broad' }
    if ($Sid -match '^S-1-5-21-\d+-\d+-\d+-(\d+)$' -and $BroadDomainRids.ContainsKey($Matches[1])) { return 'Broad' }
    if ($Sid -match '^S-1-5-21-') { return 'Account' }
    return 'WellKnown'
}

function Get-MinotaurVerdict {
    param([object[]]$FindingList)
    $sev = @($FindingList | ForEach-Object { $_.Severity })
    if ($sev -contains 'Error')   { return [PSCustomObject]@{ Verdict = 'Exposed'; Class = 'err'  } }
    if ($sev -contains 'Warning') { return [PSCustomObject]@{ Verdict = 'Review';  Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Tidy'; Class = 'ok' }
}

# ─────────────────────────────────────────────────────────────────────────────
# PRINCIPAL RESOLUTION (cached)
# ─────────────────────────────────────────────────────────────────────────────

$PrincipalCache = @{}

function Resolve-Principal {
    # Name and kind for a SID: Broad / WellKnown / User / Group / Computer /
    # Orphaned / Account (resolved, class unknown). Domain classes come from
    # an LDAP bind by SID; local ones from the SAM.
    param([string]$Sid)
    if ($PrincipalCache.ContainsKey($Sid)) { return $PrincipalCache[$Sid] }

    $category = Get-SidCategory -Sid $Sid
    $name = $Sid
    try { $name = ([System.Security.Principal.SecurityIdentifier]$Sid).Translate([System.Security.Principal.NTAccount]).Value } catch { $name = $null }

    $kind = $category
    if ($category -eq 'Broad' -and $BroadSids.Contains($Sid)) { $name = $BroadSids[$Sid] }
    if ($category -eq 'Account') {
        if (-not $name) {
            $kind = 'Orphaned'
        } else {
            $kind = 'Account'
            if (Get-LocalUser -SID $Sid -ErrorAction SilentlyContinue)  { $kind = 'User' }
            elseif (Get-LocalGroup -SID $Sid -ErrorAction SilentlyContinue) { $kind = 'Group' }
            else {
                try {
                    $entry   = [adsi]"LDAP://<SID=$Sid>"
                    $classes = @($entry.Properties['objectClass'] | ForEach-Object { "$_" })
                    if ($classes -contains 'computer')   { $kind = 'Computer' }
                    elseif ($classes -contains 'group')  { $kind = 'Group' }
                    elseif ($classes -contains 'user')   { $kind = 'User' }
                } catch {
                    $kind = 'Account'
                }
            }
        }
    }
    $result = [PSCustomObject]@{ Sid = $Sid; Name = if ($name) { $name } else { $Sid }; Kind = $kind }
    $PrincipalCache[$Sid] = $result
    return $result
}

function ConvertTo-SidString {
    param($Identity)
    try {
        if ($Identity -is [System.Security.Principal.SecurityIdentifier]) { return $Identity.Value }
        return ([System.Security.Principal.NTAccount]"$Identity").Translate([System.Security.Principal.SecurityIdentifier]).Value
    } catch {
        # A bare SID string is how Windows shows an entry it cannot resolve.
        if ("$Identity" -match '^S-1-[\d-]+$') { return "$Identity" }
        return $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-MinotaurTargets {
    if ($Path) {
        return @([PSCustomObject]@{ Name = $Path; Path = $Path; ShareAccess = $null; Description = 'Folder tree (-Path)' })
    }
    $shares = @()
    try {
        $shares = @(Get-SmbShare -ErrorAction Stop | Where-Object { -not $_.Special -and "$($_.ShareType)" -eq 'FileSystemDirectory' -and $_.Path })
    } catch {
        Write-Fail "Get-SmbShare failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'minotaur' -Message "Get-SmbShare failed: $($_.Exception.Message)" -Category 'Shares'
    }
    $targets = foreach ($s in $shares) {
        $access = @()
        try {
            $access = @(Get-SmbShareAccess -Name $s.Name -ErrorAction Stop | ForEach-Object {
                $sid = ConvertTo-SidString -Identity $_.AccountName
                [PSCustomObject]@{
                    Account = "$($_.AccountName)"
                    Sid     = $sid
                    Broad   = [bool]($sid -and (Get-SidCategory -Sid $sid) -eq 'Broad')
                    Right   = "$($_.AccessRight)"
                    Type    = "$($_.AccessControlType)"
                }
            })
        } catch {
            $access = @()
        }
        [PSCustomObject]@{ Name = $s.Name; Path = $s.Path; ShareAccess = $access; Description = "$($s.Description)" }
    }
    return @($targets)
}

function Get-FolderAces {
    # Returns the ACL's rules as rows, or $null when the folder cannot be read.
    param([string]$Folder)
    try {
        $acl = Get-Acl -LiteralPath $Folder -ErrorAction Stop
    } catch {
        return $null
    }
    $rows = foreach ($r in $acl.Access) {
        $sid = ConvertTo-SidString -Identity $r.IdentityReference
        if (-not $sid) { continue }
        $p = Resolve-Principal -Sid $sid
        [PSCustomObject]@{
            Folder    = $Folder
            Identity  = $p.Name
            Sid       = $sid
            Kind      = $p.Kind
            Rights    = "$($r.FileSystemRights)"
            Level     = Get-NtfsRightsLevel -Value ([long]$r.FileSystemRights.value__)
            Type      = "$($r.AccessControlType)"
            Inherited = [bool]$r.IsInherited
        }
    }
    return [PSCustomObject]@{ Protected = [bool]$acl.AreAccessRulesProtected; Owner = "$($acl.Owner)"; Aces = @($rows) }
}

function Invoke-MinotaurWalk {
    # Breadth-first to -Depth, recording the root and every folder that has
    # explicit entries or broken inheritance. Inherited-only folders repeat
    # their parent and are not stored.
    param([object]$Target, [ref]$Budget)

    $interesting = [System.Collections.Generic.List[object]]::new()
    $unreadable  = [System.Collections.Generic.List[string]]::new()
    $queue = [System.Collections.Generic.Queue[object]]::new()
    $queue.Enqueue(@($Target.Path, 0))
    $truncated = $false

    while ($queue.Count -gt 0) {
        $item   = $queue.Dequeue()
        $folder = $item[0]; $level = $item[1]
        if ($Budget.Value -le 0) { $truncated = $true; break }
        $Budget.Value = $Budget.Value - 1

        $acl = Get-FolderAces -Folder $folder
        if ($null -eq $acl) { [void]$unreadable.Add($folder); continue }

        $explicit = @($acl.Aces | Where-Object { -not $_.Inherited })
        if ($level -eq 0 -or $acl.Protected -or $explicit.Count -gt 0) {
            [void]$interesting.Add([PSCustomObject]@{
                Folder = $folder; Level = $level; Protected = $acl.Protected; Owner = $acl.Owner
                Aces = $acl.Aces; Explicit = $explicit
            })
        }

        if ($level -lt $Depth) {
            try {
                foreach ($child in @(Get-ChildItem -LiteralPath $folder -Directory -Force -ErrorAction Stop)) {
                    if (($child.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                    $queue.Enqueue(@($child.FullName, ($level + 1)))
                }
            } catch {
                [void]$unreadable.Add($folder)
            }
        }
    }
    return [PSCustomObject]@{ Folders = @($interesting); Unreadable = @($unreadable | Select-Object -Unique); Truncated = $truncated }
}

# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-MinotaurAudit {
    Write-Section "SHARES"
    $targets = Get-MinotaurTargets
    if ($targets.Count -eq 0) {
        Add-MinotaurFinding -Code 'NoShares' -Subject $env:COMPUTERNAME
        Write-Info "No file shares to review."
        return [PSCustomObject]@{ Shares = @(); AceRows = @() }
    }
    foreach ($t in $targets) { Write-Info ("{0,-24} {1}" -f $t.Name, $t.Path) }

    $budget  = $MaxFolders
    $results = [System.Collections.Generic.List[object]]::new()
    $aceRows = [System.Collections.Generic.List[object]]::new()

    foreach ($t in $targets) {
        Write-Section "SHARE: $($t.Name)"
        Write-Step "Walking $($t.Path) ($Depth level(s))..."
        $walk = Invoke-MinotaurWalk -Target $t -Budget ([ref]$budget)
        Write-Info ("{0} folder(s) with their own permissions, {1} unreadable" -f $walk.Folders.Count, $walk.Unreadable.Count)

        foreach ($f in $walk.Folders) {
            foreach ($a in $f.Aces) {
                [void]$aceRows.Add([PSCustomObject]@{
                    Share = $t.Name; Folder = $f.Folder; Identity = $a.Identity; Kind = $a.Kind; Level = $a.Level
                    Rights = $a.Rights; Type = $a.Type; Inherited = $a.Inherited; InheritanceDisabled = $f.Protected
                })
            }
        }

        # Share permissions gate NTFS: a broad write only lands if the share
        # also lets a broad principal change. With -Path there is no share.
        $shareBroadWrite = ($null -eq $t.ShareAccess) -or @($t.ShareAccess | Where-Object { $_.Broad -and $_.Type -eq 'Allow' -and (Test-WriteLevel $_.Right) }).Count -gt 0
        $shareBroadRead  = ($null -eq $t.ShareAccess) -or @($t.ShareAccess | Where-Object { $_.Broad -and $_.Type -eq 'Allow' }).Count -gt 0

        $root = $walk.Folders | Where-Object { $_.Level -eq 0 } | Select-Object -First 1
        $broadWrite = @($walk.Folders | ForEach-Object { $folder = $_.Folder; $_.Aces | Where-Object {
            $_.Kind -eq 'Broad' -and $_.Type -eq 'Allow' -and (Test-WriteLevel $_.Level) } | ForEach-Object { "$($_.Identity) $($_.Level) on $folder" } } | Select-Object -Unique)
        $broadRead = @()
        if ($root) {
            $broadRead = @($root.Aces | Where-Object { $_.Kind -eq 'Broad' -and $_.Type -eq 'Allow' -and -not (Test-WriteLevel $_.Level) } | ForEach-Object { "$($_.Identity) $($_.Level)" } | Select-Object -Unique)
        }

        if ($shareBroadWrite -and $broadWrite.Count -gt 0) {
            Add-MinotaurFinding -Code 'BroadWriteAccess' -Subject $t.Name -Detail (($broadWrite | Select-Object -First 10) -join '; ')
        }
        if ($shareBroadRead -and $broadRead.Count -gt 0 -and $t.Name -notin $ExpectedBroadReadShares) {
            Add-MinotaurFinding -Code 'BroadReadAccess' -Subject $t.Name -Detail ($broadRead -join '; ')
        }

        $direct = @($walk.Folders | ForEach-Object { $_.Explicit } | Where-Object { $_.Kind -eq 'User' -and $_.Type -eq 'Allow' })
        if ($direct.Count -gt 0) {
            Add-MinotaurFinding -Code 'DirectUserAce' -Subject $t.Name -Detail ("{0} entr(ies): {1}" -f $direct.Count, (@($direct | Select-Object -First 10 | ForEach-Object { "$($_.Identity) on $($_.Folder)" }) -join '; '))
        }
        $orphans = @($walk.Folders | ForEach-Object { $_.Aces } | Where-Object { $_.Kind -eq 'Orphaned' } | ForEach-Object { $_.Sid } | Select-Object -Unique)
        if ($orphans.Count -gt 0) { Add-MinotaurFinding -Code 'OrphanedSid' -Subject $t.Name -Detail ($orphans -join ', ') }
        $protected = @($walk.Folders | Where-Object { $_.Protected -and $_.Level -gt 0 })
        if ($protected.Count -gt 0) {
            Add-MinotaurFinding -Code 'InheritanceBroken' -Subject $t.Name -Detail ("{0} folder(s): {1}" -f $protected.Count, (@($protected | Select-Object -First 10 | ForEach-Object { $_.Folder }) -join '; '))
        }
        $denies = @($walk.Folders | ForEach-Object { $_.Explicit } | Where-Object { $_.Type -eq 'Deny' })
        if ($denies.Count -gt 0) {
            Add-MinotaurFinding -Code 'DenyAce' -Subject $t.Name -Detail (@($denies | Select-Object -First 10 | ForEach-Object { "$($_.Identity) on $($_.Folder)" }) -join '; ')
        }
        if ($walk.Unreadable.Count -gt 0) {
            Add-MinotaurFinding -Code 'FolderUnreadable' -Subject $t.Name -Detail ("{0} folder(s): {1}" -f $walk.Unreadable.Count, (@($walk.Unreadable | Select-Object -First 10) -join '; '))
        }
        if ($walk.Truncated) { Add-MinotaurFinding -Code 'ScanTruncated' -Subject $t.Name }

        [void]$results.Add([PSCustomObject]@{
            Name        = $t.Name
            Path        = $t.Path
            Description = $t.Description
            ShareAccess = $t.ShareAccess
            Root        = $root
            Folders     = $walk.Folders
            Unreadable  = $walk.Unreadable.Count
            BroadWrite  = ($shareBroadWrite -and $broadWrite.Count -gt 0)
            BroadRead   = ($shareBroadRead -and $broadRead.Count -gt 0)
        })
        if ($walk.Truncated) { break }
    }

    return [PSCustomObject]@{ Shares = @($results); AceRows = @($aceRows) }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Format-AceSummary {
    param([object[]]$Aces)
    return (@($Aces | Where-Object { $_.Type -eq 'Allow' } | Sort-Object Identity | ForEach-Object { "$($_.Identity): $($_.Level)" } | Select-Object -Unique) -join '; ')
}

function Build-MinotaurReport {
    param([object]$Audit, [object]$Verdict, [string]$CsvName)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $machine    = $env:COMPUTERNAME
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $order      = @{ 'Error' = 0; 'Warning' = 1; 'Info' = 2 }

    $fRows = [System.Text.StringBuilder]::new()
    if ($Findings.Count -eq 0) { [void]$fRows.Append("<tr><td colspan='4'>Nothing to report.</td></tr>") }
    foreach ($f in ($Findings | Sort-Object { $order[$_.Severity] })) {
        [void]$fRows.Append(
            "<tr><td><span class='tk-badge-$(Get-SeverityClass $f.Severity)'>$(EscHtml $f.Severity)</span></td>" +
            "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)</td>" +
            "<td>$(EscHtml $f.Subject)" + $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
            "<td>$(EscHtml $f.Remedy)</td></tr>")
    }

    $sRows = [System.Text.StringBuilder]::new()
    if ($Audit.Shares.Count -eq 0) { [void]$sRows.Append("<tr><td colspan='5'>No shares reviewed.</td></tr>") }
    foreach ($s in $Audit.Shares) {
        $share = if ($null -eq $s.ShareAccess) { '(folder tree, no share)' } else { (@($s.ShareAccess | ForEach-Object { "$($_.Account): $($_.Right)$(if ($_.Type -ne 'Allow') { " ($($_.Type))" })" }) -join '; ') }
        $ntfs  = if ($s.Root) { Format-AceSummary -Aces $s.Root.Aces } else { '(unreadable)' }
        $flag  = if ($s.BroadWrite) { "<span class='tk-badge-err'>Broad write</span>" } elseif ($s.BroadRead) { "<span class='tk-badge-warn'>Broad read</span>" } else { "<span class='tk-badge-ok'>Scoped</span>" }
        [void]$sRows.Append(
            "<tr><td><strong>$(EscHtml $s.Name)</strong><br/><span class='tk-mono'>$(EscHtml $s.Path)</span></td>" +
            "<td>$flag</td><td>$(EscHtml $share)</td><td>$(EscHtml $ntfs)</td><td>$(@($s.Folders).Count)</td></tr>")
    }

    $dRows = [System.Text.StringBuilder]::new()
    $explicitFolders = @($Audit.Shares | ForEach-Object { $share = $_.Name; $_.Folders | Where-Object { $_.Level -gt 0 } | ForEach-Object { $_ | Add-Member -NotePropertyName Share -NotePropertyValue $share -Force -PassThru } })
    if ($explicitFolders.Count -eq 0) { [void]$dRows.Append("<tr><td colspan='4'>Every folder below the share roots inherits its permissions.</td></tr>") }
    foreach ($f in ($explicitFolders | Select-Object -First 500)) {
        $explicit = Format-AceSummary -Aces $f.Explicit
        [void]$dRows.Append(
            "<tr><td>$(EscHtml $f.Share)</td><td class='tk-mono'>$(EscHtml $f.Folder)</td>" +
            "<td>$(if ($f.Protected) { "<span class='tk-badge-info'>Disabled</span>" } else { 'On' })</td><td>$(EscHtml $explicit)</td></tr>")
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'M.I.N.O.T.A.U.R. Share Permissions Report' `
        -ScriptName 'M.I.N.O.T.A.U.R.' `
        -Subtitle   "${orgPrefix}File Share & NTFS Permissions -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Generated' = $reportDate
            'Scope'     = $(if ($Path) { $Path } else { "$(@($Audit.Shares).Count) share(s)" })
            'Depth'     = "$Depth level(s)"
            'Verdict'   = $Verdict.Verdict
        }) `
        -NavItems   @('Findings', 'Shares', 'Folders With Their Own Permissions')

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Verdict</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(@($Audit.Shares).Count)</div><div class="tk-summary-lbl">Shares Reviewed</div></div>
    <div class="tk-summary-card $(if (@($Audit.Shares | Where-Object { $_.BroadWrite }).Count) { 'err' } else { 'ok' })"><div class="tk-summary-num">$(@($Audit.Shares | Where-Object { $_.BroadWrite }).Count)</div><div class="tk-summary-lbl">Broad Write</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($explicitFolders.Count)</div><div class="tk-summary-lbl">Folders With Own Permissions</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(@($Audit.AceRows).Count)</div><div class="tk-summary-lbl">Entries in CSV</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Severity</th><th>Finding</th><th>Where</th><th>Remedy</th></tr></thead>
      <tbody>$($fRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Shares</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Share</th><th>Exposure</th><th>Share permissions</th><th>NTFS at the root (allow)</th><th>Folders recorded</th></tr></thead>
      <tbody>$($sRows.ToString())</tbody></table>
      <div class="tk-info-box"><span class="tk-info-label">Effective access</span> is the more restrictive of the share and NTFS permissions. Broad write is reported only where both let an everyone-type group change files.</div></div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Folders With Their Own Permissions</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Share</th><th>Folder</th><th>Inheritance</th><th>Explicit allow entries</th></tr></thead>
      <tbody>$($dRows.ToString())</tbody></table>
      <div class="tk-info-box"><span class="tk-info-label">Full detail</span> Every access-control entry scanned is in <span class="tk-mono">$(EscHtml $CsvName)</span> next to this report.</div></div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'M.I.N.O.T.A.U.R. v6.0')
    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-MinotaurBanner

if ($Path -and -not (Test-Path -LiteralPath $Path -PathType Container)) {
    Write-Fail "-Path '$Path' is not a folder."
    exit 1
}

$audit = Invoke-MinotaurAudit

Write-Section "FINDINGS"
if ($Findings.Count -eq 0) { Write-Ok "No broad, direct or orphaned access found." }
foreach ($f in $Findings) {
    $line = "$($f.Title) -- $($f.Subject)"
    switch ($f.Severity) { 'Error' { Write-Fail $line } 'Warning' { Write-Warn $line } default { Write-Info $line } }
}

$verdict = Get-MinotaurVerdict -FindingList $Findings.ToArray()
Write-Section "VERDICT"
switch ($verdict.Class) { 'err' { Write-Fail $verdict.Verdict } 'warn' { Write-Warn $verdict.Verdict } default { Write-Ok $verdict.Verdict } }
Add-TKNote -Text ("MINOTAUR on {0}: verdict {1}; {2} share(s), {3} finding(s)." -f $env:COMPUTERNAME, $verdict.Verdict, @($audit.Shares).Count, $Findings.Count) -Category 'Info' -ScriptName 'minotaur'

$stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$logDir  = Resolve-LogDirectory -FallbackPath $ScriptPath
$csvName = "MINOTAUR_$stamp.csv"
try {
    $audit.AceRows | Export-Csv -LiteralPath (Join-Path $logDir $csvName) -NoTypeInformation -Encoding UTF8
    Write-Ok "Access-control entries exported: $csvName"
} catch {
    Write-Warn "Could not write the CSV: $($_.Exception.Message)"
}

Write-Step "Generating HTML report..."
$outPath = Join-Path $logDir "MINOTAUR_$stamp.html"
try {
    [System.IO.File]::WriteAllText($outPath, (Build-MinotaurReport -Audit $audit -Verdict $verdict -CsvName $csvName), [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
    Write-TKError -ScriptName 'minotaur' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
}

# The note buffer ends with this run, so save it now -- one file per run.
$null = Save-TKNoteReport -Prefix 'MINOTAUR' -Directory (Resolve-LogDirectory -FallbackPath $ScriptPath)

if (-not $Unattended) { Read-Host "  Press Enter to exit" | Out-Null }
if ($Transcript) { Stop-TKTranscript }
