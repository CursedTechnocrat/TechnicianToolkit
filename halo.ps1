# halo.ps1 - H.A.L.O. — Holistic Access-policy Logic Overview
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
    H.A.L.O. — Holistic Access-policy Logic Overview
    Entra ID Conditional Access Posture Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Scores an Entra ID tenant's Conditional Access policies against the
    baseline every tenant should have, then reviews the policies themselves:

      - Baseline coverage: MFA for all users, MFA for privileged roles,
        legacy authentication blocked, and -- where licensed -- sign-in and
        user risk policies and a device-compliance requirement. Security
        defaults are taken into account when no Conditional Access policy
        exists.
      - Hygiene: policies left in report-only or disabled, broad exclusion
        lists on enforcing policies, policies that reference deleted users
        or groups, and trusted named locations wide enough to exempt most
        of the internet.
      - Emergency access: the accounts or groups excluded from every
        enforcing all-users policy -- the break-glass path. None at all
        means a misconfigured policy can lock every administrator out.

    Read-only -- nothing in the tenant is changed. Requires only the
    Microsoft.Graph.Authentication module (offered for install if missing)
    and the Policy.Read.All and Directory.Read.All delegated scopes.

.USAGE
    PS C:\> .\halo.ps1                    # Interactive: sign in, audit, HTML report
    PS C:\> .\halo.ps1 -Unattended        # Silent: sign in, audit, export HTML

.NOTES
    Version : 6.2

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

$GraphScopes = @('Policy.Read.All', 'Directory.Read.All')

# Directory role template IDs are the same in every tenant. These are the roles
# that can take over the tenant or its security configuration, so each should
# sit behind an enforcing MFA policy.
$PrivilegedRoleTemplates = [ordered]@{
    '62e90394-69f5-4237-9190-012177145e10' = 'Global Administrator'
    'e8611ab8-c189-46e8-94e1-60213ab1f814' = 'Privileged Role Administrator'
    '7be44c8a-adaf-4e2a-84d6-ab2649e08a13' = 'Privileged Authentication Administrator'
    '194ae4cb-b126-40b2-bd5b-6091b380977d' = 'Security Administrator'
    'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9' = 'Conditional Access Administrator'
    '29232cdf-9323-42fd-ade2-1d097af3e4de' = 'Exchange Administrator'
    'f28a1f50-f6e7-4571-818b-6a12f2af6b6c' = 'SharePoint Administrator'
    'fe930be7-5e62-47db-91af-98c3a49a38b1' = 'User Administrator'
    'c4e39bd9-1100-46d3-8c65-fb160da0071f' = 'Authentication Administrator'
    '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3' = 'Application Administrator'
    '158c047a-c907-4556-b7ef-446551a6b5f7' = 'Cloud Application Administrator'
    '729827e3-9c14-49f7-bb1b-9608f156bbb8' = 'Helpdesk Administrator'
    '966707d0-3269-4727-9be2-8c3a10f19b9d' = 'Password Administrator'
    '3a2c62db-5318-420d-8d74-23affee5d9d5' = 'Intune Administrator'
}
$GlobalAdminTemplateId = '62e90394-69f5-4237-9190-012177145e10'

# An enforcing all-users policy with more direct exclusions than this is worth
# a second look -- each exclusion is a user the baseline does not reach.
$ExclusionThreshold = 5

# Include/exclude values that are keywords rather than object IDs.
$TargetKeywords = @('All', 'None', 'GuestsOrExternalUsers')

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
#
# Every condition H.A.L.O. can report, keyed by a stable code. The Pester
# suite extracts it by AST lookup and checks it against the codes the script
# raises.
# ─────────────────────────────────────────────────────────────────────────────

$HaloFindings = @{
    'NoBaseline' = @{
        Severity = 'Error'
        Title    = 'No Conditional Access policy and security defaults off'
        Summary  = 'Nothing requires MFA or blocks legacy authentication anywhere in the tenant.'
        Remedy   = 'Turn security defaults on today, or build the baseline policies: MFA for all users, MFA for admins, block legacy authentication.'
    }
    'SecurityDefaultsOnly' = @{
        Severity = 'Info'
        Title    = 'Protected by security defaults, not Conditional Access'
        Summary  = 'Security defaults give every tenant MFA and a legacy-auth block, but allow no exclusions, locations or device conditions.'
        Remedy   = 'Fine for small tenants. Move to Conditional Access (Entra ID P1) when you need exceptions or device-based rules.'
    }
    'NoMfaAllUsers' = @{
        Severity = 'Error'
        Title    = 'No enforced MFA policy for all users'
        Summary  = 'No enabled policy requires MFA (or an authentication strength) for all users on all cloud apps, so a stolen password alone signs in.'
        Remedy   = 'Create a policy: All users, All cloud apps, Grant: require MFA. Exclude only the emergency-access accounts.'
    }
    'NoMfaAdmins' = @{
        Severity = 'Error'
        Title    = 'Global Administrator is not behind enforced MFA'
        Summary  = 'No enabled policy requires MFA for the Global Administrator role on all cloud apps.'
        Remedy   = 'Create a policy targeting the privileged directory roles, All cloud apps, requiring MFA or a phishing-resistant authentication strength.'
    }
    'AdminRolesNotCovered' = @{
        Severity = 'Warning'
        Title    = 'Privileged roles outside enforced MFA'
        Summary  = 'These privileged directory roles are not covered by any enabled MFA policy for all cloud apps.'
        Remedy   = 'Add them to the admin MFA policy, or cover them with an all-users MFA policy that does not exclude them.'
    }
    'LegacyAuthNotBlocked' = @{
        Severity = 'Error'
        Title    = 'Legacy authentication is not blocked'
        Summary  = 'No enabled policy blocks Exchange ActiveSync and other legacy clients for all users. Legacy protocols cannot do MFA, so they bypass it.'
        Remedy   = 'Create a policy: All users, All cloud apps, Client apps: Exchange ActiveSync + Other clients, Grant: block.'
    }
    'NoRiskPolicies' = @{
        Severity = 'Info'
        Title    = 'No sign-in or user risk policy'
        Summary  = 'No enabled policy reacts to Entra ID Protection risk. Needs Entra ID P2.'
        Remedy   = 'If licensed: require MFA at medium sign-in risk and a secure password change at high user risk.'
    }
    'NoDevicePolicy' = @{
        Severity = 'Info'
        Title    = 'No device-based policy'
        Summary  = 'No enabled policy requires a compliant or hybrid-joined device, so any device with valid credentials and MFA is trusted.'
        Remedy   = 'Consider requiring a compliant device for admin portals or sensitive apps once Intune enrolment is in place.'
    }
    'ReportOnlyPolicy' = @{
        Severity = 'Warning'
        Title    = 'Policy left in report-only mode'
        Summary  = 'Report-only policies are evaluated and logged but never enforced.'
        Remedy   = 'Review its sign-in log impact, then switch it On -- or delete it if it is no longer wanted.'
    }
    'DisabledPolicy' = @{
        Severity = 'Info'
        Title    = 'Disabled policies'
        Summary  = 'Disabled policies do nothing. Old ones make the policy set harder to reason about.'
        Remedy   = 'Delete the ones that are no longer needed.'
    }
    'BroadExclusion' = @{
        Severity = 'Warning'
        Title    = 'Enforcing all-users policy with a long exclusion list'
        Summary  = 'Every excluded user or group skips the control. Long lists usually hold one-off exceptions nobody removed.'
        Remedy   = 'Trim the exclusions to the emergency-access accounts and documented service accounts.'
    }
    'NoEmergencyAccess' = @{
        Severity = 'Warning'
        Title    = 'No emergency-access exclusion'
        Summary  = 'No user or group is excluded from every enforcing all-users policy. A mistake in one policy, or an MFA outage, can lock every administrator out.'
        Remedy   = 'Create two cloud-only emergency-access Global Administrators with long random passwords (or FIDO2 keys), exclude them from all policies, and alert on their sign-ins.'
    }
    'EmergencyAccessPresent' = @{
        Severity = 'Info'
        Title    = 'Emergency-access exclusion in place'
        Summary  = 'These users or groups are excluded from every enforcing all-users policy -- the break-glass path.'
        Remedy   = 'Confirm they are the intended accounts, cloud-only, and that their sign-ins raise an alert.'
    }
    'StaleObjectReference' = @{
        Severity = 'Warning'
        Title    = 'Policy references deleted users or groups'
        Summary  = 'An include or exclude list names an object that no longer exists. A deleted include silently narrows the policy.'
        Remedy   = 'Edit the policy and remove or replace the missing objects.'
    }
    'BroadTrustedLocation' = @{
        Severity = 'Warning'
        Title    = 'Trusted named location is very broad'
        Summary  = 'A trusted IP range this wide covers far more than an office network. Policies that skip MFA from trusted locations skip it for everyone inside it.'
        Remedy   = 'Narrow the range to the organisation''s actual egress addresses.'
    }
    'GraphQueryFailed' = @{
        Severity = 'Warning'
        Title    = 'Part of the tenant could not be read'
        Summary  = 'A Microsoft Graph query failed, so the results below may be incomplete.'
        Remedy   = 'Sign in with an account holding Security Reader, Global Reader or Conditional Access Administrator, and consent to Policy.Read.All.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()

function Add-HaloFinding {
    param(
        [Parameter(Mandatory)][string]$Code,
        [string]$Subject = '',
        [string]$Detail = ''
    )
    $meta = $HaloFindings[$Code]
    if (-not $meta) { $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' } }
    [void]$Findings.Add([PSCustomObject]@{
        Code = $Code; Severity = $meta.Severity; Title = $meta.Title; Summary = $meta.Summary
        Remedy = $meta.Remedy; Subject = $Subject; Detail = $Detail
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

function Show-HaloBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██╗  ██╗ █████╗ ██╗      ██████╗
  ██║  ██║██╔══██╗██║     ██╔═══██╗
  ███████║███████║██║     ██║   ██║
  ██╔══██║██╔══██║██║     ██║   ██║
  ██║  ██║██║  ██║███████╗╚██████╔╝
  ╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝ ╚═════╝

"@ -ForegroundColor Cyan
    Write-Host "    H.A.L.O. — Holistic Access-policy Logic Overview" -ForegroundColor Cyan
    Write-Host "    Entra ID Conditional Access Posture Audit Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PURE HELPERS
#
# Policies arrive from Graph as nested hashtables; the tests pass hashtables or
# PSCustomObjects. Member access reads both the same way, so nothing below
# cares which. No I/O until the COLLECTORS section.
# ─────────────────────────────────────────────────────────────────────────────

function Get-CaList {
    # Graph returns missing collections as $null and single values unwrapped.
    param($Value)
    if ($null -eq $Value) { return @() }
    return @($Value | ForEach-Object { "$_" } | Where-Object { $_ })
}

function Test-CaPolicyEnforced {
    param($Policy)
    return "$($Policy.state)" -eq 'enabled'
}

function Test-CaTargetsAllUsers {
    param($Policy)
    return (Get-CaList $Policy.conditions.users.includeUsers) -contains 'All'
}

function Test-CaTargetsAllApps {
    param($Policy)
    return (Get-CaList $Policy.conditions.applications.includeApplications) -contains 'All'
}

function Test-CaRequiresMfa {
    # True when every way through the grant control includes MFA: an
    # authentication strength, or 'mfa' under AND, or an OR list offering
    # nothing but MFA. "MFA OR compliant device" lets a compliant device in
    # without MFA, so it does not count.
    param($Policy)
    $grant = $Policy.grantControls
    if ($null -eq $grant) { return $false }
    if ($null -ne $grant.authenticationStrength) { return $true }
    $controls = Get-CaList $grant.builtInControls
    if ($controls -notcontains 'mfa') { return $false }
    if ("$($grant.operator)" -eq 'OR' -and @($controls | Where-Object { $_ -ne 'mfa' }).Count -gt 0) { return $false }
    return $true
}

function Test-CaBlocks {
    param($Policy)
    return (Get-CaList $Policy.grantControls.builtInControls) -contains 'block'
}

function Test-CaBlocksLegacyAuth {
    # A block for all users whose client-app condition includes both legacy
    # client types. 'all' client apps would block everything, which is not a
    # legacy-auth policy and would lock the tenant, so it does not count.
    param($Policy)
    if (-not (Test-CaBlocks $Policy) -or -not (Test-CaTargetsAllUsers $Policy)) { return $false }
    $apps = Get-CaList $Policy.conditions.clientAppTypes
    return ($apps -contains 'exchangeActiveSync') -and ($apps -contains 'other')
}

function Test-CaCoversRole {
    # Does this policy put the given role behind its control on all apps?
    param($Policy, [string]$RoleTemplateId)
    if (-not (Test-CaTargetsAllApps $Policy)) { return $false }
    $users = $Policy.conditions.users
    if ((Get-CaList $users.excludeRoles) -contains $RoleTemplateId) { return $false }
    if ((Get-CaList $users.includeUsers) -contains 'All') { return $true }
    return (Get-CaList $users.includeRoles) -contains $RoleTemplateId
}

function Test-CaUsesRisk {
    param($Policy)
    return (@(Get-CaList $Policy.conditions.signInRiskLevels).Count + @(Get-CaList $Policy.conditions.userRiskLevels).Count) -gt 0
}

function Test-CaUsesDevice {
    param($Policy)
    $controls = Get-CaList $Policy.grantControls.builtInControls
    return ($controls -contains 'compliantDevice') -or ($controls -contains 'domainJoinedDevice')
}

function Get-CaEmergencyExclusion {
    # The users and groups excluded from EVERY enforcing all-users policy that
    # blocks or requires MFA -- anyone outside that set is caught by at least
    # one of them. Returns $null when there is no such policy to be excluded from.
    param([object[]]$Policies)
    $gates = @($Policies | Where-Object {
        (Test-CaPolicyEnforced $_) -and (Test-CaTargetsAllUsers $_) -and ((Test-CaBlocks $_) -or (Test-CaRequiresMfa $_))
    })
    if ($gates.Count -eq 0) { return $null }

    $users  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $groups = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($id in (Get-CaList $gates[0].conditions.users.excludeUsers))  { [void]$users.Add($id) }
    foreach ($id in (Get-CaList $gates[0].conditions.users.excludeGroups)) { [void]$groups.Add($id) }
    foreach ($p in ($gates | Select-Object -Skip 1)) {
        $users.IntersectWith([string[]]@(Get-CaList $p.conditions.users.excludeUsers))
        $groups.IntersectWith([string[]]@(Get-CaList $p.conditions.users.excludeGroups))
    }
    return [PSCustomObject]@{
        GateCount = $gates.Count
        Users     = @($users)
        Groups    = @($groups)
    }
}

function Test-BroadCidr {
    # IPv4 wider than /16, or IPv6 wider than /32, is more than any one
    # organisation's egress. An unparsable range is not judged.
    param([string]$Cidr)
    $parts = "$Cidr".Split('/')
    if ($parts.Count -ne 2) { return $false }
    $prefix = 0
    if (-not [int]::TryParse($parts[1], [ref]$prefix)) { return $false }
    $ip = $null
    if (-not [System.Net.IPAddress]::TryParse($parts[0], [ref]$ip)) { return $false }
    if ($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) { return $prefix -lt 32 }
    return $prefix -lt 16
}

function Get-HaloVerdict {
    param([object[]]$FindingList)
    $sev = @($FindingList | ForEach-Object { $_.Severity })
    if ($sev -contains 'Error')   { return [PSCustomObject]@{ Verdict = 'Exposed';  Class = 'err'  } }
    if ($sev -contains 'Warning') { return [PSCustomObject]@{ Verdict = 'Gaps';     Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Enforced'; Class = 'ok' }
}

# ─────────────────────────────────────────────────────────────────────────────
# MODULE + CONNECTION
# ─────────────────────────────────────────────────────────────────────────────

function Install-HaloModule {
    Write-Section "MODULE CHECK"
    $name = 'Microsoft.Graph.Authentication'
    if (Get-Module -ListAvailable -Name $name) {
        Write-Ok "$name — installed"
    } else {
        Write-Warn "$name — NOT found"
        $doInstall = $Unattended
        if (-not $Unattended) {
            $ans = Read-Host "  Install $name for current user? [Y/N]"
            $doInstall = $ans -match '^[Yy]'
        }
        if (-not $doInstall) { Write-Fail "$name is required."; return $false }
        try {
            Install-Module -Name $name -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            Write-Ok "$name installed."
        } catch {
            Write-Fail "Install failed: $($_.Exception.Message)"
            Write-TKError -ScriptName 'halo' -Message "$name install failed: $($_.Exception.Message)" -Category 'Module Install'
            return $false
        }
    }
    try { Import-Module $name -ErrorAction Stop; return $true }
    catch { Write-Fail "Could not import ${name}: $($_.Exception.Message)"; return $false }
}

function Connect-HaloGraph {
    Write-Section "CONNECT TO MICROSOFT GRAPH"
    try {
        $ctx = Get-MgContext -ErrorAction Stop
        if ($ctx -and $ctx.Account -and -not @($GraphScopes | Where-Object { $ctx.Scopes -notcontains $_ })) {
            Write-Ok "Reusing existing session: $($ctx.Account)"
            return $ctx.Account
        }
    } catch {
        Write-Verbose "No cached Graph context: $($_.Exception.Message)"
    }
    Write-Step "Requesting interactive sign-in..."
    Write-Info "Scopes: $($GraphScopes -join ', ')"
    try {
        Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext
        Write-Ok "Connected as: $($ctx.Account)"
        return $ctx.Account
    } catch {
        Write-Fail "Connect-MgGraph failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'halo' -Message "Connect-MgGraph failed: $($_.Exception.Message)" -Category 'Graph Auth'
        return $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-HaloGraph {
    # GET with paging. Returns every item of 'value', or the single object for
    # endpoints that are not collections.
    param([string]$Uri, [string]$Label)
    $items = [System.Collections.Generic.List[object]]::new()
    $next  = $Uri
    try {
        while ($next) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
            if ($resp.ContainsKey('value')) {
                foreach ($v in @($resp['value'])) { [void]$items.Add($v) }
                $next = $resp['@odata.nextLink']
            } else {
                return $resp
            }
        }
    } catch {
        Add-HaloFinding -Code 'GraphQueryFailed' -Subject $Label -Detail $_.Exception.Message
        Write-Warn "$Label query failed: $($_.Exception.Message)"
        return $null
    }
    return , $items.ToArray()
}

function Resolve-HaloObjects {
    # directoryObjects/getByIds resolves users and groups in one call per 1000
    # IDs. Anything asked for and not returned no longer exists.
    param([string[]]$Ids)
    $names = @{}
    $ids   = @($Ids | Where-Object { $_ -and $_ -notin $TargetKeywords } | Select-Object -Unique)
    for ($i = 0; $i -lt $ids.Count; $i += 1000) {
        $chunk = @($ids[$i..([math]::Min($i + 999, $ids.Count - 1))])
        try {
            $body = @{ ids = $chunk; types = @('user', 'group') } | ConvertTo-Json -Depth 3
            $resp = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/directoryObjects/getByIds' -Body $body -ContentType 'application/json' -ErrorAction Stop
            foreach ($o in @($resp['value'])) {
                $label = if ($o['userPrincipalName']) { $o['userPrincipalName'] } else { $o['displayName'] }
                $names["$($o['id'])"] = "$label"
            }
        } catch {
            Add-HaloFinding -Code 'GraphQueryFailed' -Subject 'directoryObjects/getByIds' -Detail $_.Exception.Message
            return $null
        }
    }
    return $names
}

# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-HaloAudit {
    Write-Section "CONDITIONAL ACCESS"
    Write-Step "Reading policies, named locations and security defaults..."
    $policies  = Invoke-HaloGraph -Uri 'v1.0/identity/conditionalAccess/policies' -Label 'Conditional Access policies'
    $locations = Invoke-HaloGraph -Uri 'v1.0/identity/conditionalAccess/namedLocations' -Label 'Named locations'
    $defaults  = Invoke-HaloGraph -Uri 'v1.0/policies/identitySecurityDefaultsEnforcementPolicy' -Label 'Security defaults'

    $policies  = @($policies | Where-Object { $_ })
    $locations = @($locations | Where-Object { $_ })
    $defaultsOn = [bool]($defaults -and $defaults['isEnabled'])
    $enforced  = @($policies | Where-Object { Test-CaPolicyEnforced $_ })

    Write-Info ("Policies: {0} ({1} enforced)  |  Named locations: {2}  |  Security defaults: {3}" -f `
        $policies.Count, $enforced.Count, $locations.Count, $(if ($defaultsOn) { 'on' } else { 'off' }))

    # Resolve every user / group ID the policies mention, once.
    $ids = foreach ($p in $policies) {
        $u = $p.conditions.users
        Get-CaList $u.includeUsers; Get-CaList $u.excludeUsers; Get-CaList $u.includeGroups; Get-CaList $u.excludeGroups
    }
    $names = Resolve-HaloObjects -Ids @($ids)
    function _name { param([string]$Id) if ($names -and $names.ContainsKey($Id)) { $names[$Id] } else { $Id } }

    # ── Baseline coverage ──
    $coverage = [System.Collections.Generic.List[object]]::new()
    function _cover { param([string]$Check, [object[]]$By, [string]$Missing)
        [void]$coverage.Add([PSCustomObject]@{ Check = $Check; Met = ($By.Count -gt 0); By = (@($By | ForEach-Object { $_.displayName }) -join '; '); Missing = $Missing })
    }

    if ($policies.Count -eq 0 -or $enforced.Count -eq 0) {
        if ($defaultsOn) { Add-HaloFinding -Code 'SecurityDefaultsOnly' -Subject 'Tenant' }
        else             { Add-HaloFinding -Code 'NoBaseline' -Subject 'Tenant' }
    }

    $mfaAll = @($enforced | Where-Object { (Test-CaTargetsAllUsers $_) -and (Test-CaTargetsAllApps $_) -and (Test-CaRequiresMfa $_) })
    $legacy = @($enforced | Where-Object { Test-CaBlocksLegacyAuth $_ })
    $risk   = @($enforced | Where-Object { Test-CaUsesRisk $_ })
    $device = @($enforced | Where-Object { Test-CaUsesDevice $_ })
    $mfaGa  = @($enforced | Where-Object { (Test-CaRequiresMfa $_) -and (Test-CaCoversRole -Policy $_ -RoleTemplateId $GlobalAdminTemplateId) })

    _cover 'MFA for all users, all apps'       $mfaAll 'NoMfaAllUsers'
    _cover 'MFA for Global Administrator'      $mfaGa  'NoMfaAdmins'
    _cover 'Legacy authentication blocked'     $legacy 'LegacyAuthNotBlocked'
    _cover 'Sign-in / user risk policy'        $risk   'NoRiskPolicies'
    _cover 'Device-based requirement'          $device 'NoDevicePolicy'

    # Security defaults already enforce MFA and block legacy auth, so with
    # them on (and no CA to replace them) those two baseline gaps are closed.
    if (-not $defaultsOn -or $enforced.Count -gt 0) {
        if ($mfaAll.Count -eq 0) { Add-HaloFinding -Code 'NoMfaAllUsers' -Subject 'Tenant' }
        if ($mfaGa.Count  -eq 0) { Add-HaloFinding -Code 'NoMfaAdmins'  -Subject 'Global Administrator' }
        if ($legacy.Count -eq 0) { Add-HaloFinding -Code 'LegacyAuthNotBlocked' -Subject 'Tenant' }
    }
    if ($enforced.Count -gt 0) {
        if ($risk.Count   -eq 0) { Add-HaloFinding -Code 'NoRiskPolicies' -Subject 'Tenant' }
        if ($device.Count -eq 0) { Add-HaloFinding -Code 'NoDevicePolicy' -Subject 'Tenant' }
    }

    if ($mfaGa.Count -gt 0) {
        $uncovered = @($PrivilegedRoleTemplates.Keys | Where-Object {
            $role = $_
            @($enforced | Where-Object { (Test-CaRequiresMfa $_) -and (Test-CaCoversRole -Policy $_ -RoleTemplateId $role) }).Count -eq 0
        } | ForEach-Object { $PrivilegedRoleTemplates[$_] })
        if ($uncovered.Count -gt 0) { Add-HaloFinding -Code 'AdminRolesNotCovered' -Subject "$($uncovered.Count) role(s)" -Detail ($uncovered -join ', ') }
    }

    # ── Hygiene ──
    foreach ($p in ($policies | Where-Object { "$($_.state)" -eq 'enabledForReportingButNotEnforced' })) {
        Add-HaloFinding -Code 'ReportOnlyPolicy' -Subject "$($p.displayName)"
    }
    $disabled = @($policies | Where-Object { "$($_.state)" -eq 'disabled' })
    if ($disabled.Count -gt 0) {
        Add-HaloFinding -Code 'DisabledPolicy' -Subject "$($disabled.Count) policy(ies)" -Detail (@($disabled | ForEach-Object { $_.displayName }) -join '; ')
    }
    foreach ($p in ($enforced | Where-Object { (Test-CaTargetsAllUsers $_) -and ((Test-CaBlocks $_) -or (Test-CaRequiresMfa $_)) })) {
        $ex = @(Get-CaList $p.conditions.users.excludeUsers) + @(Get-CaList $p.conditions.users.excludeGroups)
        if ($ex.Count -gt $ExclusionThreshold) {
            Add-HaloFinding -Code 'BroadExclusion' -Subject "$($p.displayName)" -Detail ("{0} exclusions: {1}" -f $ex.Count, (@($ex | Select-Object -First 15 | ForEach-Object { _name $_ }) -join ', '))
        }
    }

    if ($null -ne $names) {
        $stale = [System.Collections.Generic.List[string]]::new()
        foreach ($p in $policies) {
            $u = $p.conditions.users
            $missing = @(@(Get-CaList $u.includeUsers) + @(Get-CaList $u.excludeUsers) + @(Get-CaList $u.includeGroups) + @(Get-CaList $u.excludeGroups) |
                Where-Object { $_ -notin $TargetKeywords -and -not $names.ContainsKey($_) })
            if ($missing.Count -gt 0) { [void]$stale.Add("$($p.displayName): $($missing -join ', ')") }
        }
        if ($stale.Count -gt 0) { Add-HaloFinding -Code 'StaleObjectReference' -Subject "$($stale.Count) policy(ies)" -Detail ($stale -join ' | ') }
    }

    $emergency = Get-CaEmergencyExclusion -Policies $policies
    if ($null -ne $emergency) {
        $who = @(@($emergency.Users | ForEach-Object { "user $(_name $_)" }) + @($emergency.Groups | ForEach-Object { "group $(_name $_)" }))
        if ($who.Count -eq 0) { Add-HaloFinding -Code 'NoEmergencyAccess' -Subject "$($emergency.GateCount) enforcing all-users policy(ies)" }
        else                  { Add-HaloFinding -Code 'EmergencyAccessPresent' -Subject "$($who.Count) exclusion(s)" -Detail ($who -join ', ') }
    }

    $locationRows = foreach ($l in $locations) {
        $ranges = @($l.ipRanges | ForEach-Object { "$($_.cidrAddress)" } | Where-Object { $_ })
        $broad  = @($ranges | Where-Object { Test-BroadCidr $_ })
        if ([bool]$l.isTrusted -and $broad.Count -gt 0) {
            Add-HaloFinding -Code 'BroadTrustedLocation' -Subject "$($l.displayName)" -Detail ($broad -join ', ')
        }
        [PSCustomObject]@{
            Name    = "$($l.displayName)"
            Type    = if ($ranges.Count -gt 0) { 'IP ranges' } elseif ($l.countriesAndRegions) { 'Countries' } else { 'Other' }
            Trusted = [bool]$l.isTrusted
            Detail  = if ($ranges.Count -gt 0) { $ranges -join ', ' } else { (@($l.countriesAndRegions) -join ', ') }
            Broad   = $broad.Count -gt 0
        }
    }

    # ── Inventory rows ──
    $policyRows = foreach ($p in ($policies | Sort-Object { "$($_.state)" -ne 'enabled' }, { "$($_.displayName)" })) {
        $u = $p.conditions.users
        $inc = @(@(Get-CaList $u.includeUsers | ForEach-Object { if ($_ -in $TargetKeywords) { $_ } else { _name $_ } }) +
                 @(Get-CaList $u.includeGroups | ForEach-Object { "group $(_name $_)" }) +
                 @(Get-CaList $u.includeRoles  | ForEach-Object { if ($PrivilegedRoleTemplates.Contains($_)) { "role $($PrivilegedRoleTemplates[$_])" } else { "role $_" } }))
        $exc = @(@(Get-CaList $u.excludeUsers | ForEach-Object { _name $_ }) + @(Get-CaList $u.excludeGroups | ForEach-Object { "group $(_name $_)" }) + @(Get-CaList $u.excludeRoles | ForEach-Object { "role $_" }))
        $grant = @(Get-CaList $p.grantControls.builtInControls)
        if ($p.grantControls -and $p.grantControls.authenticationStrength) { $grant += "strength: $($p.grantControls.authenticationStrength.displayName)" }
        [PSCustomObject]@{
            Name       = "$($p.displayName)"
            State      = "$($p.state)"
            Users      = $inc -join ', '
            Excluded   = $exc -join ', '
            Apps       = (Get-CaList $p.conditions.applications.includeApplications) -join ', '
            ClientApps = (Get-CaList $p.conditions.clientAppTypes) -join ', '
            Grant      = if ($grant.Count) { "$(if ($p.grantControls.operator) { "$($p.grantControls.operator): " })$($grant -join ', ')" } else { '(session controls only)' }
        }
    }

    return [PSCustomObject]@{
        Policies     = @($policyRows)
        PolicyCount  = $policies.Count
        Enforced     = $enforced.Count
        DefaultsOn   = $defaultsOn
        Coverage     = @($coverage)
        Locations    = @($locationRows)
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# CONSOLE + HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Show-HaloFindings {
    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) { Write-Ok "Nothing to report."; return }
    $order = @{ 'Error' = 0; 'Warning' = 1; 'Info' = 2 }
    foreach ($f in ($Findings | Sort-Object { $order[$_.Severity] })) {
        $line = "$($f.Title)" + $(if ($f.Subject) { " -- $($f.Subject)" } else { '' })
        switch ($f.Severity) { 'Error' { Write-Fail $line } 'Warning' { Write-Warn $line } default { Write-Info $line } }
    }
}

function Build-HaloReport {
    param([object]$Audit, [object]$Verdict, [string]$ConnectedAs)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant     = if ($ConnectedAs -match '@(.+)$') { $Matches[1] } else { 'Tenant' }
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

    $cRows = [System.Text.StringBuilder]::new()
    foreach ($c in $Audit.Coverage) {
        $sev   = $HaloFindings[$c.Missing].Severity
        $badge = if ($c.Met) { "<span class='tk-badge-ok'>Met</span>" } else { "<span class='tk-badge-$(Get-SeverityClass $sev)'>Missing</span>" }
        [void]$cRows.Append("<tr><td>$(EscHtml $c.Check)</td><td>$badge</td><td>$(EscHtml $(if ($c.By) { $c.By } else { '-' }))</td></tr>")
    }
    if ($Audit.DefaultsOn) {
        [void]$cRows.Append("<tr><td>Security defaults</td><td><span class='tk-badge-info'>On</span></td><td>Tenant-wide MFA registration and legacy-auth block</td></tr>")
    }

    $pRows = [System.Text.StringBuilder]::new()
    if ($Audit.Policies.Count -eq 0) { [void]$pRows.Append("<tr><td colspan='6'>No Conditional Access policies.</td></tr>") }
    foreach ($p in $Audit.Policies) {
        $stateBadge = switch ($p.State) {
            'enabled'                          { "<span class='tk-badge-ok'>On</span>" }
            'enabledForReportingButNotEnforced' { "<span class='tk-badge-warn'>Report-only</span>" }
            default                            { "<span class='tk-badge-info'>Off</span>" }
        }
        [void]$pRows.Append(
            "<tr><td><strong>$(EscHtml $p.Name)</strong></td><td>$stateBadge</td>" +
            "<td>$(EscHtml $p.Users)" + $(if ($p.Excluded) { "<br/><span class='tk-mono'>excl: $(EscHtml $p.Excluded)</span>" } else { '' }) + "</td>" +
            "<td>$(EscHtml $p.Apps)</td><td>$(EscHtml $p.ClientApps)</td><td>$(EscHtml $p.Grant)</td></tr>")
    }

    $lRows = [System.Text.StringBuilder]::new()
    if ($Audit.Locations.Count -eq 0) { [void]$lRows.Append("<tr><td colspan='4'>No named locations.</td></tr>") }
    foreach ($l in $Audit.Locations) {
        $trust = if ($l.Trusted -and $l.Broad) { "<span class='tk-badge-warn'>Trusted (broad)</span>" } elseif ($l.Trusted) { "<span class='tk-badge-ok'>Trusted</span>" } else { 'No' }
        [void]$lRows.Append("<tr><td>$(EscHtml $l.Name)</td><td>$(EscHtml $l.Type)</td><td>$trust</td><td class='tk-mono'>$(EscHtml $l.Detail)</td></tr>")
    }

    $metCount = @($Audit.Coverage | Where-Object { $_.Met }).Count

    $htmlHead = Get-TKHtmlHead `
        -Title      'H.A.L.O. Conditional Access Posture Report' `
        -ScriptName 'H.A.L.O.' `
        -Subtitle   "${orgPrefix}Conditional Access Posture -- $tenant" `
        -MetaItems  ([ordered]@{
            'Tenant'       = $tenant
            'Signed in as' = $ConnectedAs
            'Generated'    = $reportDate
            'Verdict'      = $Verdict.Verdict
        }) `
        -NavItems   @('Findings', 'Baseline Coverage', 'Policies', 'Named Locations')

    $errCount  = @($Findings | Where-Object { $_.Severity -eq 'Error' }).Count
    $warnCount = @($Findings | Where-Object { $_.Severity -eq 'Warning' }).Count

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Verdict</div></div>
    <div class="tk-summary-card $(if ($errCount) { 'err' } else { 'ok' })"><div class="tk-summary-num">$errCount</div><div class="tk-summary-lbl">Baseline Gaps</div></div>
    <div class="tk-summary-card $(if ($warnCount) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$warnCount</div><div class="tk-summary-lbl">Warnings</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$metCount / $($Audit.Coverage.Count)</div><div class="tk-summary-lbl">Baseline Checks Met</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Audit.Enforced) / $($Audit.PolicyCount)</div><div class="tk-summary-lbl">Policies Enforced</div></div>
    <div class="tk-summary-card $(if ($Audit.DefaultsOn) { 'ok' } else { 'info' })"><div class="tk-summary-num">$(if ($Audit.DefaultsOn) { 'On' } else { 'Off' })</div><div class="tk-summary-lbl">Security Defaults</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Severity</th><th>Finding</th><th>Where</th><th>Remedy</th></tr></thead>
      <tbody>$($fRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Baseline Coverage</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Check</th><th>Status</th><th>Satisfied by</th></tr></thead>
      <tbody>$($cRows.ToString())</tbody></table>
      <div class="tk-info-box"><span class="tk-info-label">Counting rule</span> Only enabled policies count. "MFA OR compliant device" does not count as MFA, because a compliant device gets in without it.</div></div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Policies</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Policy</th><th>State</th><th>Users</th><th>Apps</th><th>Client apps</th><th>Grant</th></tr></thead>
      <tbody>$($pRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> Named Locations</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Name</th><th>Type</th><th>Trusted</th><th>Ranges / countries</th></tr></thead>
      <tbody>$($lRows.ToString())</tbody></table></div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'H.A.L.O. v6.2')

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-HaloRun {
    $Findings.Clear()

    if (-not (Install-HaloModule)) { if ($Unattended) { exit 1 }; return }
    $who = Connect-HaloGraph
    if (-not $who) { if ($Unattended) { exit 1 }; return }

    $audit = Invoke-HaloAudit
    Show-HaloFindings

    $verdict = Get-HaloVerdict -FindingList $Findings.ToArray()
    Write-Section "VERDICT"
    switch ($verdict.Class) { 'err' { Write-Fail $verdict.Verdict } 'warn' { Write-Warn $verdict.Verdict } default { Write-Ok $verdict.Verdict } }

    Add-TKNote -Text ("HALO Conditional Access audit: verdict {0}; {1} of {2} policies enforced; {3} finding(s)." -f $verdict.Verdict, $audit.Enforced, $audit.PolicyCount, $Findings.Count) -Category 'Info' -ScriptName 'halo'

    Write-Step "Generating HTML report..."
    $html    = Build-HaloReport -Audit $audit -Verdict $verdict -ConnectedAs "$who"
    $outPath = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) ("HALO_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'halo' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

Show-HaloBanner
Invoke-HaloRun

# The note buffer ends with this run, so save it now -- one file per run.
$null = Save-TKNoteReport -Prefix 'HALO' -Directory (Resolve-LogDirectory -FallbackPath $ScriptPath)

if (-not $Unattended) { Read-Host "  Press Enter to exit" | Out-Null }
if ($Transcript) { Stop-TKTranscript }
