# raven.ps1 - R.A.V.E.N. — Reviews Auto-forwarding, Vulnerable Exchange settings & Nefarious rules
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
    R.A.V.E.N. — Reviews Auto-forwarding, Vulnerable Exchange settings & Nefarious rules
    Exchange Online Mailbox Security Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Audits an Exchange Online tenant for the signs and preconditions of a
    compromised mailbox -- the business-email-compromise pattern where an
    attacker signs in, quietly forwards or hides mail, and waits:

      - Mailbox forwarding (ForwardingSmtpAddress / ForwardingAddress) to
        addresses outside the tenant's accepted domains
      - Inbox rules that forward externally, move mail into folders nobody
        reads and mark it read, delete messages about payments or security,
        or carry the throwaway names attackers give them
      - Outbound spam policies that allow automatic external forwarding, and
        transport rules that redirect or copy mail outside the organisation
      - SMTP AUTH left enabled org-wide or per mailbox, and mailbox auditing
        switched off
      - Full Access and Send As delegation, for access review
      - SPF, DKIM and DMARC for each domain

    The DNS checks need no sign-in: -DnsOnly -Domain contoso.com checks email
    authentication for any domain, which also works before a tenant is taken on.

    Read-only -- nothing in the tenant is changed. Requires the
    ExchangeOnlineManagement module (offered for install if missing) and an
    account that can read recipient and transport configuration (Global Reader
    or View-Only Organization Management is enough).

.USAGE
    PS C:\> .\raven.ps1                                   # Interactive menu
    PS C:\> .\raven.ps1 -Unattended                       # Full audit: sign in, audit, HTML report
    PS C:\> .\raven.ps1 -Unattended -SkipDelegation       # Skip the per-mailbox permission sweep (faster)
    PS C:\> .\raven.ps1 -Unattended -DnsOnly -Domain contoso.com, fabrikam.com   # SPF / DKIM / DMARC only, no sign-in

.NOTES
    Version : 6.2

#>

param(
    [switch]$Unattended,
    [string[]]$Domain,
    [switch]$DnsOnly,
    [switch]$SkipDelegation,
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
$C = $ColorSchema

# ─────────────────────────────────────────────────────────────────────────────
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

# Folders an inbox rule can file mail into where the mailbox owner will not see
# it. Moving mail here and marking it read is the classic way an attacker hides
# the replies to a fraudulent payment request.
$HiddenFolderPattern = '(?i)(^|\\)(RSS Feeds|RSS Subscriptions|Conversation History|Archive|Junk E-?mail|Deleted Items|Notes|Sync Issues|Outbox)$'

# Words that, in an inbox rule that deletes or hides mail, point at an attacker
# suppressing the evidence: payment conversations, or the warnings about them.
$SensitiveKeywordPattern = '(?i)\b(invoice|payment|wire|bank|transfer|remittance|ach|payroll|w-?2|password|security|verify|verification|phish|phishing|hack|hacked|compromise|compromised|suspicious|fraud|helpdesk|help desk|it support|mfa)\b'

# A rule name of nothing but punctuation, digits or a character or two -- '.',
# '..', 'a', '1' -- is what attackers name the rules they create.
$SuspiciousRuleNamePattern = '^[\W_\d]{0,3}$|^.$'

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
#
# Every condition R.A.V.E.N. can report, keyed by a stable code. Kept as one
# table so the Pester suite can extract it by AST lookup and assert that every
# code the tool raises exists here.
# ─────────────────────────────────────────────────────────────────────────────

$RavenFindings = @{
    'ExternalForwarding' = @{
        Severity = 'Error'
        Title    = 'Mailbox forwards to an external address'
        Summary  = 'Mailbox-level forwarding sends every message to an address outside the tenant. Set by an attacker, it silently copies all mail out.'
        Remedy   = 'Confirm with the mailbox owner. If unexpected, remove the forward, reset the password, revoke sessions and review sign-in logs.'
    }
    'InternalForwarding' = @{
        Severity = 'Info'
        Title    = 'Mailbox forwards internally'
        Summary  = 'Forwarding to another mailbox in the tenant -- usually a leaver''s mail routed to their manager.'
        Remedy   = 'Confirm the forward is still wanted; convert a leaver''s mailbox to shared instead where possible.'
    }
    'InboxRuleExternalForward' = @{
        Severity = 'Error'
        Title    = 'Inbox rule forwards or redirects externally'
        Summary  = 'A user-level inbox rule sends mail to an address outside the tenant. This is the most common persistence step after a mailbox compromise.'
        Remedy   = 'Confirm with the owner. If unexpected: disable the rule, reset the password, revoke sessions, and search the audit log for what was sent.'
    }
    'InboxRuleHidesMail' = @{
        Severity = 'Error'
        Title    = 'Inbox rule hides or deletes mail'
        Summary  = 'The rule moves mail to a folder nobody reads and marks it read, or deletes messages about payments or security -- how an attacker keeps the owner from seeing replies.'
        Remedy   = 'Treat the mailbox as compromised until shown otherwise: disable the rule, reset credentials, revoke sessions, review sign-ins.'
    }
    'InboxRuleSuspicious' = @{
        Severity = 'Warning'
        Title    = 'Inbox rule with suspicious traits'
        Summary  = 'A throwaway name (".", "..", a single character) or a move into a rarely opened folder. Not proof of compromise, but worth a look.'
        Remedy   = 'Ask the owner whether they created it.'
    }
    'AutoForwardAllowed' = @{
        Severity = 'Warning'
        Title    = 'Outbound spam policy allows external auto-forwarding'
        Summary  = 'AutoForwardingMode = On lets any mailbox forward automatically outside the tenant, which removes the safety net if one is compromised.'
        Remedy   = 'Set AutoForwardingMode to Off (or Automatic) and allow forwarding only for the mailboxes that need it, through a separate policy.'
    }
    'TransportRuleExternalCopy' = @{
        Severity = 'Warning'
        Title    = 'Transport rule sends mail outside the organisation'
        Summary  = 'An enabled mail-flow rule redirects, blind-copies or adds an external recipient. An attacker with admin access can use one to copy mail for the whole tenant.'
        Remedy   = 'Confirm the rule has a documented business owner; remove it if not.'
    }
    'SmtpAuthOrgEnabled' = @{
        Severity = 'Warning'
        Title    = 'SMTP AUTH enabled for the organisation'
        Summary  = 'Authenticated SMTP is a legacy protocol that bypasses Conditional Access prompts. Enabled org-wide, every mailbox accepts it.'
        Remedy   = 'Set-TransportConfig -SmtpClientAuthenticationDisabled $true, then enable it per mailbox only for devices that still need it (scanners, line-of-business apps).'
    }
    'SmtpAuthMailboxEnabled' = @{
        Severity = 'Info'
        Title    = 'SMTP AUTH explicitly enabled on mailboxes'
        Summary  = 'These mailboxes override the organisation setting and accept authenticated SMTP.'
        Remedy   = 'Confirm each one belongs to a device or app that needs it. Prefer a dedicated, MFA-excluded service mailbox over a user''s own.'
    }
    'AuditDisabled' = @{
        Severity = 'Error'
        Title    = 'Mailbox auditing is disabled for the organisation'
        Summary  = 'AuditDisabled = True: mailbox actions are not recorded, so an investigation after a compromise has nothing to read.'
        Remedy   = 'Set-OrganizationConfig -AuditDisabled $false.'
    }
    'FullAccessOnUserMailbox' = @{
        Severity = 'Info'
        Title    = 'Full Access granted on user mailboxes'
        Summary  = 'Delegation on shared mailboxes is normal; on a person''s own mailbox it should have a reason.'
        Remedy   = 'Review the delegates listed in the Delegation section and remove any without a current reason.'
    }
    'SpfMissing' = @{
        Severity = 'Error'
        Title    = 'No SPF record'
        Summary  = 'Without SPF, receiving servers cannot tell which servers may send for the domain, and spoofed mail is more likely to be delivered.'
        Remedy   = 'Publish a TXT record such as: v=spf1 include:spf.protection.outlook.com -all (add every other service that sends as the domain).'
    }
    'SpfInvalid' = @{
        Severity = 'Error'
        Title    = 'SPF record is broken or permits everyone'
        Summary  = 'More than one SPF record is a permanent error that receivers treat as no SPF at all; "+all" authorises every server on the internet.'
        Remedy   = 'Merge into a single v=spf1 record ending in ~all or -all.'
    }
    'SpfWeak' = @{
        Severity = 'Warning'
        Title    = 'SPF record does not reject unlisted senders'
        Summary  = 'The record ends in "?all" or has no "all" mechanism, so mail from unlisted servers is treated as neutral.'
        Remedy   = 'End the record with ~all (soft fail) or -all (fail).'
    }
    'DmarcMissing' = @{
        Severity = 'Error'
        Title    = 'No DMARC record'
        Summary  = 'Without DMARC, receivers have no instruction for mail that fails SPF and DKIM, and the domain gets no reports of who is sending as it.'
        Remedy   = 'Publish _dmarc TXT: v=DMARC1; p=none; rua=mailto:<reports address> -- then move to quarantine and reject once reports are clean.'
    }
    'DmarcInvalid' = @{
        Severity = 'Error'
        Title    = 'DMARC record is broken'
        Summary  = 'More than one DMARC record, or one without a valid p= policy, is ignored by receivers.'
        Remedy   = 'Publish exactly one _dmarc TXT record with p=none, quarantine or reject.'
    }
    'DmarcMonitorOnly' = @{
        Severity = 'Warning'
        Title    = 'DMARC is monitor-only'
        Summary  = 'p=none (or pct below 100) collects reports but tells receivers to deliver spoofed mail anyway.'
        Remedy   = 'Review the aggregate reports, then move to p=quarantine and p=reject.'
    }
    'DkimDisabled' = @{
        Severity = 'Warning'
        Title    = 'DKIM signing not enabled in Exchange Online'
        Summary  = 'Mail from the domain is not DKIM-signed with its own key, so DMARC can only pass on SPF, which breaks on forwarding.'
        Remedy   = 'Publish the selector1 / selector2 CNAMEs shown in the Defender portal, then enable DKIM signing for the domain.'
    }
    'DkimSelectorsMissing' = @{
        Severity = 'Info'
        Title    = 'Microsoft 365 DKIM selectors not published'
        Summary  = 'selector1._domainkey and selector2._domainkey do not resolve. Expected if the domain sends through another provider that uses its own selectors.'
        Remedy   = 'If the domain sends through Microsoft 365, publish both selector CNAMEs and enable DKIM signing.'
    }
    'DnsLookupFailed' = @{
        Severity = 'Warning'
        Title    = 'DNS lookup failed'
        Summary  = 'A lookup failed for a reason other than "no such record", so the result for this domain is incomplete.'
        Remedy   = 'Check this machine''s DNS resolution and rerun.'
    }
    'MailboxScanIncomplete' = @{
        Severity = 'Warning'
        Title    = 'Some mailboxes could not be scanned'
        Summary  = 'Inbox rules or permissions could not be read for part of the tenant, so the results below may be incomplete.'
        Remedy   = 'Rerun with an account holding Global Reader or View-Only Organization Management.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()

function Add-RavenFinding {
    param(
        [Parameter(Mandatory)][string]$Code,
        [string]$Subject = '',
        [string]$Detail = ''
    )
    $meta = $RavenFindings[$Code]
    if (-not $meta) {
        $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' }
    }
    [void]$Findings.Add([PSCustomObject]@{
        Code     = $Code
        Severity = $meta.Severity
        Title    = $meta.Title
        Summary  = $meta.Summary
        Remedy   = $meta.Remedy
        Subject  = $Subject
        Detail   = $Detail
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

function Show-RavenBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██████╗  █████╗ ██╗   ██╗███████╗███╗   ██╗
  ██╔══██╗██╔══██╗██║   ██║██╔════╝████╗  ██║
  ██████╔╝███████║██║   ██║█████╗  ██╔██╗ ██║
  ██╔══██╗██╔══██║╚██╗ ██╔╝██╔══╝  ██║╚██╗██║
  ██║  ██║██║  ██║ ╚████╔╝ ███████╗██║ ╚████║
  ╚═╝  ╚═╝╚═╝  ╚═╝  ╚═══╝  ╚══════╝╚═╝  ╚═══╝

"@ -ForegroundColor Cyan
    Write-Host "    R.A.V.E.N. — Reviews Auto-forwarding, Vulnerable Exchange settings & Nefarious rules" -ForegroundColor Cyan
    Write-Host "    Exchange Online Mailbox Security Audit Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PURE HELPERS
#
# No I/O until the COLLECTORS section -- the Pester suite extracts these by
# AST lookup and calls them directly with synthetic input.
# ─────────────────────────────────────────────────────────────────────────────

function Get-SmtpAddressFromRecipient {
    # Inbox-rule targets arrive as '"Name" [SMTP:a@b.com]', forwarding as
    # 'smtp:a@b.com', and an internal recipient may be '"Name" [EX:/o=...]'
    # with no SMTP address at all -- which is internal by construction.
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    $m = [regex]::Match($Value, '(?i)smtp:([^\]\s">]+)')
    if ($m.Success) { return $m.Groups[1].Value.Trim().ToLowerInvariant() }
    if ($Value -match '(?i)\[EX:') { return '' }
    $m = [regex]::Match($Value, "[A-Za-z0-9._%+'\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
    if ($m.Success) { return $m.Value.ToLowerInvariant() }
    return ''
}

function Test-ExternalAddress {
    # External means the address's domain is neither an accepted domain nor a
    # subdomain of one. An empty address (an EX: recipient) is internal.
    param([string]$Address, [string[]]$InternalDomains)
    if ([string]::IsNullOrWhiteSpace($Address) -or $Address -notmatch '@') { return $false }
    $dom = ($Address -split '@')[-1].Trim().TrimEnd('.').ToLowerInvariant()
    foreach ($d in $InternalDomains) {
        $d = "$d".Trim().TrimEnd('.').ToLowerInvariant()
        if (-not $d) { continue }
        if ($dom -eq $d -or $dom.EndsWith(".$d")) { return $false }
    }
    return $true
}

function Get-InboxRuleRisk {
    # Scores one inbox rule. Returns the worst severity and every reason, so the
    # report shows why a rule was flagged rather than only that it was.
    param([object]$Rule, [string[]]$InternalDomains)

    $reasons  = [System.Collections.Generic.List[string]]::new()
    $external = [System.Collections.Generic.List[string]]::new()
    $codes    = [System.Collections.Generic.List[string]]::new()

    foreach ($prop in 'ForwardTo', 'ForwardAsAttachmentTo', 'RedirectTo') {
        foreach ($target in @($Rule.$prop)) {
            if (-not $target) { continue }
            $addr = Get-SmtpAddressFromRecipient -Value "$target"
            if (Test-ExternalAddress -Address $addr -InternalDomains $InternalDomains) {
                if (-not $external.Contains($addr)) { [void]$external.Add($addr) }
            }
        }
    }
    if ($external.Count -gt 0) {
        [void]$reasons.Add("Sends to external address(es): $($external -join ', ')")
        [void]$codes.Add('InboxRuleExternalForward')
    }

    $conditionText = (@(
        $Rule.SubjectContainsWords, $Rule.SubjectOrBodyContainsWords, $Rule.BodyContainsWords,
        $Rule.FromAddressContainsWords
    ) | ForEach-Object { $_ } | Where-Object { $_ }) -join ' '
    $sensitive = $conditionText -match $SensitiveKeywordPattern

    $folder     = "$($Rule.MoveToFolder)"
    $hiddenMove = $folder -and ($folder -match $HiddenFolderPattern)
    $deletes    = [bool]$Rule.DeleteMessage -or [bool]$Rule.SoftDeleteMessage

    if ($hiddenMove -and [bool]$Rule.MarkAsRead) {
        [void]$reasons.Add("Moves mail to '$folder' and marks it read")
        [void]$codes.Add('InboxRuleHidesMail')
    } elseif ($hiddenMove -and $sensitive) {
        [void]$reasons.Add("Moves mail about payments or security to '$folder'")
        [void]$codes.Add('InboxRuleHidesMail')
    } elseif ($hiddenMove) {
        [void]$reasons.Add("Moves mail to rarely opened folder '$folder'")
        [void]$codes.Add('InboxRuleSuspicious')
    }

    if ($deletes -and $sensitive) {
        [void]$reasons.Add("Deletes mail matching: $conditionText")
        [void]$codes.Add('InboxRuleHidesMail')
    }

    if ("$($Rule.Name)" -match $SuspiciousRuleNamePattern) {
        [void]$reasons.Add("Throwaway rule name '$($Rule.Name)'")
        [void]$codes.Add('InboxRuleSuspicious')
    }

    $severity = ''
    if ($codes -contains 'InboxRuleExternalForward' -or $codes -contains 'InboxRuleHidesMail') { $severity = 'Error' }
    elseif ($codes.Count -gt 0) { $severity = 'Warning' }

    return [PSCustomObject]@{
        Severity  = $severity
        Codes     = @($codes | Select-Object -Unique)
        Reasons   = @($reasons)
        External  = @($external)
    }
}

function Get-SpfVerdict {
    # Takes every TXT string published at the domain apex.
    param([string[]]$Records)
    $spf = @($Records | Where-Object { $_ -match '^\s*"?v=spf1(\s|$)' })

    if ($spf.Count -eq 0) { return [PSCustomObject]@{ Status = 'Missing';  Code = 'SpfMissing'; Record = ''; IncludesM365 = $false } }
    if ($spf.Count -gt 1) { return [PSCustomObject]@{ Status = 'Multiple'; Code = 'SpfInvalid'; Record = ($spf -join ' | '); IncludesM365 = $false } }

    $rec  = $spf[0].Trim()
    $m365 = $rec -match '(?i)include:spf\.protection\.outlook\.com'
    $all  = [regex]::Match($rec, '(?i)(?:^|\s)([+\-~?]?)all(?:\s|$)')

    if (-not $all.Success) {
        if ($rec -match '(?i)(?:^|\s)redirect=') {
            return [PSCustomObject]@{ Status = 'Redirect'; Code = ''; Record = $rec; IncludesM365 = $m365 }
        }
        return [PSCustomObject]@{ Status = 'No all mechanism'; Code = 'SpfWeak'; Record = $rec; IncludesM365 = $m365 }
    }
    switch ($all.Groups[1].Value) {
        '-'     { return [PSCustomObject]@{ Status = 'Hard fail (-all)'; Code = '';           Record = $rec; IncludesM365 = $m365 } }
        '~'     { return [PSCustomObject]@{ Status = 'Soft fail (~all)'; Code = '';           Record = $rec; IncludesM365 = $m365 } }
        '?'     { return [PSCustomObject]@{ Status = 'Neutral (?all)';   Code = 'SpfWeak';    Record = $rec; IncludesM365 = $m365 } }
        default { return [PSCustomObject]@{ Status = 'Pass all (+all)';  Code = 'SpfInvalid'; Record = $rec; IncludesM365 = $m365 } }
    }
}

function Get-DmarcVerdict {
    # Takes every TXT string published at _dmarc.<domain>.
    param([string[]]$Records)
    $dmarc = @($Records | Where-Object { $_ -match '(?i)^\s*"?v=DMARC1' })

    if ($dmarc.Count -eq 0) { return [PSCustomObject]@{ Status = 'Missing';  Code = 'DmarcMissing'; Policy = ''; Pct = $null; Record = '' } }
    if ($dmarc.Count -gt 1) { return [PSCustomObject]@{ Status = 'Multiple'; Code = 'DmarcInvalid'; Policy = ''; Pct = $null; Record = ($dmarc -join ' | ') } }

    $rec    = $dmarc[0].Trim()
    $pMatch = [regex]::Match($rec, '(?i)(?:^|;)\s*p\s*=\s*(none|quarantine|reject)\s*(?:;|$)')
    if (-not $pMatch.Success) {
        return [PSCustomObject]@{ Status = 'No valid policy'; Code = 'DmarcInvalid'; Policy = ''; Pct = $null; Record = $rec }
    }
    $policy = $pMatch.Groups[1].Value.ToLowerInvariant()

    $pct = 100
    $pctMatch = [regex]::Match($rec, '(?i)(?:^|;)\s*pct\s*=\s*(\d+)')
    if ($pctMatch.Success) { $pct = [int]$pctMatch.Groups[1].Value }

    $code = ''
    if ($policy -eq 'none' -or $pct -lt 100) { $code = 'DmarcMonitorOnly' }
    return [PSCustomObject]@{ Status = "p=$policy" + $(if ($pct -lt 100) { " (pct=$pct)" } else { '' }); Code = $code; Policy = $policy; Pct = $pct; Record = $rec }
}

function Get-RavenVerdict {
    param([object[]]$FindingList)
    $sev = @($FindingList | ForEach-Object { $_.Severity })
    if ($sev -contains 'Error')   { return [PSCustomObject]@{ Verdict = 'At Risk'; Class = 'err'  } }
    if ($sev -contains 'Warning') { return [PSCustomObject]@{ Verdict = 'Review';  Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Clean'; Class = 'ok' }
}

# ─────────────────────────────────────────────────────────────────────────────
# MODULE + CONNECTION
# ─────────────────────────────────────────────────────────────────────────────

function Install-ExchangeModule {
    Write-Section "MODULE CHECK"
    if (Get-Module -ListAvailable -Name 'ExchangeOnlineManagement') {
        Write-Ok "ExchangeOnlineManagement — installed"
    } else {
        Write-Warn "ExchangeOnlineManagement — NOT found"
        $doInstall = $Unattended
        if (-not $Unattended) {
            $ans = Read-Host "  Install ExchangeOnlineManagement for current user? [Y/N]"
            $doInstall = $ans -match '^[Yy]'
        }
        if (-not $doInstall) {
            Write-Fail "ExchangeOnlineManagement is required for the full audit. Use -DnsOnly for the DNS checks alone."
            return $false
        }
        try {
            Install-Module -Name 'ExchangeOnlineManagement' -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            Write-Ok "ExchangeOnlineManagement installed."
        } catch {
            Write-Fail "Install failed: $($_.Exception.Message)"
            Write-TKError -ScriptName 'raven' -Message "ExchangeOnlineManagement install failed: $($_.Exception.Message)" -Category 'Module Install'
            return $false
        }
    }
    try {
        Import-Module 'ExchangeOnlineManagement' -ErrorAction Stop
        return $true
    } catch {
        Write-Fail "Could not import ExchangeOnlineManagement: $($_.Exception.Message)"
        return $false
    }
}

function Connect-RavenExchange {
    Write-Section "CONNECT TO EXCHANGE ONLINE"
    try {
        $existing = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' })
        if ($existing.Count -gt 0) {
            Write-Ok "Reusing existing session: $($existing[0].UserPrincipalName)"
            return $existing[0].UserPrincipalName
        }
    } catch {
        # Older module versions lack Get-ConnectionInformation -- fall through to connect.
        Write-Verbose "Get-ConnectionInformation unavailable: $($_.Exception.Message)"
    }
    Write-Step "Requesting interactive sign-in..."
    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $info = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' }) | Select-Object -First 1
        $who  = if ($info) { $info.UserPrincipalName } else { 'connected' }
        Write-Ok "Connected as: $who"
        return $who
    } catch {
        Write-Fail "Connect-ExchangeOnline failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'raven' -Message "Connect-ExchangeOnline failed: $($_.Exception.Message)" -Category 'EXO Auth'
        return $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS — DNS
# ─────────────────────────────────────────────────────────────────────────────

function Resolve-RavenDns {
    # Distinguishes "no such record" (an answer) from a failed lookup (no answer).
    param([string]$Name, [ValidateSet('TXT', 'CNAME')][string]$Type)
    try {
        $answers = @(Resolve-DnsName -Name $Name -Type $Type -DnsOnly -ErrorAction Stop | Where-Object { $_.Type -eq $Type })
        $values = foreach ($a in $answers) {
            if ($Type -eq 'TXT') { ($a.Strings -join '') } else { $a.NameHost }
        }
        return [PSCustomObject]@{ Values = @($values); Failed = $false; Error = '' }
    } catch {
        if ($_.Exception.Message -match 'does not exist|DNS name does not exist|9003|No records|9501') {
            return [PSCustomObject]@{ Values = @(); Failed = $false; Error = '' }
        }
        return [PSCustomObject]@{ Values = @(); Failed = $true; Error = $_.Exception.Message }
    }
}

function Get-RavenDomainAuth {
    param([string[]]$Domains, [hashtable]$DkimConfig)

    $rows = foreach ($d in $Domains) {
        Write-Step "Checking $d..."
        $txt   = Resolve-RavenDns -Name $d -Type TXT
        $dm    = Resolve-RavenDns -Name "_dmarc.$d" -Type TXT
        $sel1  = Resolve-RavenDns -Name "selector1._domainkey.$d" -Type CNAME
        $sel2  = Resolve-RavenDns -Name "selector2._domainkey.$d" -Type CNAME

        $failed = @(@($txt, $dm, $sel1, $sel2) | Where-Object { $_.Failed })
        if ($failed.Count -gt 0) { Add-RavenFinding -Code 'DnsLookupFailed' -Subject $d -Detail $failed[0].Error }

        $spf   = Get-SpfVerdict   -Records $txt.Values
        $dmarc = Get-DmarcVerdict -Records $dm.Values
        if (-not $txt.Failed -and $spf.Code)  { Add-RavenFinding -Code $spf.Code   -Subject $d -Detail $(if ($spf.Record) { $spf.Record } else { $spf.Status }) }
        if (-not $dm.Failed -and $dmarc.Code) { Add-RavenFinding -Code $dmarc.Code -Subject $d -Detail $(if ($dmarc.Record) { $dmarc.Record } else { $dmarc.Status }) }

        $selectorsPublished = ($sel1.Values.Count -gt 0 -and $sel2.Values.Count -gt 0)
        $dkimState = 'Not checked'
        if ($null -ne $DkimConfig) {
            if ($DkimConfig.ContainsKey($d.ToLowerInvariant()) -and $DkimConfig[$d.ToLowerInvariant()]) {
                $dkimState = 'Enabled'
            } else {
                $dkimState = 'Not enabled'
                Add-RavenFinding -Code 'DkimDisabled' -Subject $d -Detail $(if ($selectorsPublished) { 'Selector CNAMEs are published; signing is off.' } else { 'Selector CNAMEs are not published either.' })
            }
        } elseif (-not $selectorsPublished -and -not $sel1.Failed -and -not $sel2.Failed) {
            $dkimState = 'M365 selectors not published'
            Add-RavenFinding -Code 'DkimSelectorsMissing' -Subject $d
        } elseif ($selectorsPublished) {
            $dkimState = 'M365 selectors published'
        }

        [PSCustomObject]@{
            Domain       = $d
            SpfStatus    = $spf.Status
            SpfCode      = $spf.Code
            SpfRecord    = $spf.Record
            SpfM365      = $spf.IncludesM365
            DmarcStatus  = $dmarc.Status
            DmarcCode    = $dmarc.Code
            DmarcRecord  = $dmarc.Record
            Dkim         = $dkimState
            Selectors    = $selectorsPublished
        }
    }
    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS — EXCHANGE ONLINE
# ─────────────────────────────────────────────────────────────────────────────

function Get-RavenTenantSettings {
    param([string[]]$InternalDomains)

    Write-Section "TENANT SETTINGS"

    $org = $null; $transport = $null
    try { $org = Get-OrganizationConfig -ErrorAction Stop } catch { Write-Warn "Get-OrganizationConfig failed: $($_.Exception.Message)" }
    try { $transport = Get-TransportConfig -ErrorAction Stop } catch { Write-Warn "Get-TransportConfig failed: $($_.Exception.Message)" }

    if ($org -and $org.AuditDisabled) { Add-RavenFinding -Code 'AuditDisabled' -Subject $org.Name }
    if ($transport -and -not $transport.SmtpClientAuthenticationDisabled) {
        Add-RavenFinding -Code 'SmtpAuthOrgEnabled' -Subject 'Organisation' -Detail 'SmtpClientAuthenticationDisabled = False'
    }

    $spamPolicies = @()
    try {
        $spamPolicies = @(Get-HostedOutboundSpamFilterPolicy -ErrorAction Stop | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; IsDefault = [bool]$_.IsDefault; AutoForwardingMode = "$($_.AutoForwardingMode)" }
        })
        foreach ($p in $spamPolicies | Where-Object { $_.AutoForwardingMode -eq 'On' }) {
            Add-RavenFinding -Code 'AutoForwardAllowed' -Subject $p.Name -Detail 'AutoForwardingMode = On'
        }
    } catch {
        Write-Warn "Get-HostedOutboundSpamFilterPolicy failed: $($_.Exception.Message)"
    }

    $transportRules = @()
    try {
        foreach ($r in @(Get-TransportRule -ResultSize Unlimited -ErrorAction Stop)) {
            $targets = @(@($r.RedirectMessageTo) + @($r.BlindCopyTo) + @($r.AddToRecipients) + @($r.CopyTo) | Where-Object { $_ })
            $ext = @($targets | ForEach-Object { Get-SmtpAddressFromRecipient -Value "$_" } |
                Where-Object { Test-ExternalAddress -Address $_ -InternalDomains $InternalDomains } | Select-Object -Unique)
            if ($ext.Count -eq 0) { continue }
            $row = [PSCustomObject]@{ Name = $r.Name; State = "$($r.State)"; External = ($ext -join ', ') }
            $transportRules += $row
            if ($row.State -eq 'Enabled') {
                Add-RavenFinding -Code 'TransportRuleExternalCopy' -Subject $r.Name -Detail $row.External
            }
        }
    } catch {
        Write-Warn "Get-TransportRule failed: $($_.Exception.Message)"
    }

    Write-Info ("Mailbox auditing  : {0}" -f $(if (-not $org) { 'unknown' } elseif ($org.AuditDisabled) { 'DISABLED' } else { 'on' }))
    Write-Info ("SMTP AUTH (org)   : {0}" -f $(if (-not $transport) { 'unknown' } elseif ($transport.SmtpClientAuthenticationDisabled) { 'disabled' } else { 'ENABLED' }))
    Write-Info ("Outbound policies : {0}" -f $spamPolicies.Count)

    return [PSCustomObject]@{
        OrgName          = if ($org) { $org.Name } else { '' }
        AuditDisabled    = if ($org) { [bool]$org.AuditDisabled } else { $null }
        SmtpAuthDisabled = if ($transport) { [bool]$transport.SmtpClientAuthenticationDisabled } else { $null }
        SpamPolicies     = $spamPolicies
        TransportRules   = $transportRules
    }
}

function Get-RavenMailboxScan {
    param([string[]]$InternalDomains)

    Write-Section "MAILBOXES"
    Write-Step "Enumerating user and shared mailboxes..."
    $mailboxes = @()
    try {
        $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox `
            -Properties ForwardingSmtpAddress, ForwardingAddress, DeliverToMailboxAndForward -ErrorAction Stop)
    } catch {
        Write-Fail "Get-EXOMailbox failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'raven' -Message "Get-EXOMailbox failed: $($_.Exception.Message)" -Category 'EXO Query'
    }
    Write-Info "$($mailboxes.Count) mailbox(es)."

    # ── Mailbox-level forwarding ──
    $forwarding = [System.Collections.Generic.List[object]]::new()
    foreach ($mb in $mailboxes) {
        $targets = @()
        if ($mb.ForwardingSmtpAddress) {
            $targets += [PSCustomObject]@{ Via = 'ForwardingSmtpAddress'; Address = (Get-SmtpAddressFromRecipient -Value "$($mb.ForwardingSmtpAddress)") }
        }
        if ($mb.ForwardingAddress) {
            $addr = ''
            try {
                $rcp  = Get-EXORecipient -Identity "$($mb.ForwardingAddress)" -Properties ExternalEmailAddress -ErrorAction Stop
                $addr = if ($rcp.ExternalEmailAddress) { Get-SmtpAddressFromRecipient -Value "$($rcp.ExternalEmailAddress)" } else { "$($rcp.PrimarySmtpAddress)".ToLowerInvariant() }
            } catch {
                $addr = "$($mb.ForwardingAddress)"
            }
            $targets += [PSCustomObject]@{ Via = 'ForwardingAddress'; Address = $addr }
        }
        foreach ($t in $targets) {
            $isExternal = Test-ExternalAddress -Address $t.Address -InternalDomains $InternalDomains
            $row = [PSCustomObject]@{
                Mailbox    = "$($mb.UserPrincipalName)"
                Type       = "$($mb.RecipientTypeDetails)"
                Via        = $t.Via
                Target     = $t.Address
                External   = $isExternal
                KeepsCopy  = [bool]$mb.DeliverToMailboxAndForward
            }
            [void]$forwarding.Add($row)
            $copyNote = if ($row.KeepsCopy) { 'keeps a copy' } else { 'no copy kept in the mailbox' }
            if ($isExternal) { Add-RavenFinding -Code 'ExternalForwarding' -Subject $row.Mailbox -Detail "$($row.Target) via $($row.Via); $copyNote" }
            else             { Add-RavenFinding -Code 'InternalForwarding' -Subject $row.Mailbox -Detail "$($row.Target) via $($row.Via); $copyNote" }
        }
    }

    # ── SMTP AUTH per mailbox ──
    $smtpEnabled = @()
    try {
        $smtpEnabled = @(Get-EXOCASMailbox -ResultSize Unlimited -Properties SmtpClientAuthenticationDisabled -ErrorAction Stop |
            Where-Object { $_.SmtpClientAuthenticationDisabled -eq $false } |
            ForEach-Object { "$($_.PrimarySmtpAddress)" })
        if ($smtpEnabled.Count -gt 0) {
            Add-RavenFinding -Code 'SmtpAuthMailboxEnabled' -Subject "$($smtpEnabled.Count) mailbox(es)" -Detail (($smtpEnabled | Select-Object -First 20) -join ', ')
        }
    } catch {
        Write-Warn "Get-EXOCASMailbox failed: $($_.Exception.Message)"
    }

    # ── Inbox rules and delegation, per mailbox ──
    $flaggedRules = [System.Collections.Generic.List[object]]::new()
    $delegation   = [System.Collections.Generic.List[object]]::new()
    $rulesScanned = 0
    $failures     = 0
    $i = 0
    foreach ($mb in $mailboxes) {
        $i++
        $upn = "$($mb.UserPrincipalName)"
        Write-Progress -Activity 'R.A.V.E.N. mailbox scan' -Status $upn -PercentComplete ([int](($i / [math]::Max(1, $mailboxes.Count)) * 100))

        try {
            $rules = @(Get-InboxRule -Mailbox $upn -ErrorAction Stop -WarningAction SilentlyContinue)
            $rulesScanned += $rules.Count
            foreach ($r in $rules) {
                $risk = Get-InboxRuleRisk -Rule $r -InternalDomains $InternalDomains
                if (-not $risk.Severity) { continue }
                $row = [PSCustomObject]@{
                    Mailbox  = $upn
                    Rule     = "$($r.Name)"
                    Enabled  = [bool]$r.Enabled
                    Severity = $risk.Severity
                    Reasons  = ($risk.Reasons -join '; ')
                }
                [void]$flaggedRules.Add($row)
                foreach ($code in $risk.Codes) {
                    Add-RavenFinding -Code $code -Subject "$upn / $($row.Rule)" -Detail ("{0}{1}" -f $row.Reasons, $(if (-not $row.Enabled) { ' (rule is disabled)' } else { '' }))
                }
            }
        } catch {
            $failures++
        }

        if (-not $SkipDelegation) {
            try {
                foreach ($p in @(Get-EXOMailboxPermission -Identity $upn -ErrorAction Stop)) {
                    if ($p.IsInherited -or $p.Deny -or "$($p.User)" -match '^NT AUTHORITY\\SELF$|^S-1-5-') { continue }
                    if (@($p.AccessRights) -notcontains 'FullAccess') { continue }
                    [void]$delegation.Add([PSCustomObject]@{ Mailbox = $upn; Type = "$($mb.RecipientTypeDetails)"; Delegate = "$($p.User)"; Right = 'Full Access' })
                }
                foreach ($p in @(Get-EXORecipientPermission -Identity $upn -ErrorAction Stop)) {
                    if ($p.IsInherited -or "$($p.Trustee)" -match '^NT AUTHORITY\\SELF$|^S-1-5-') { continue }
                    if (@($p.AccessRights) -notcontains 'SendAs') { continue }
                    [void]$delegation.Add([PSCustomObject]@{ Mailbox = $upn; Type = "$($mb.RecipientTypeDetails)"; Delegate = "$($p.Trustee)"; Right = 'Send As' })
                }
            } catch {
                $failures++
            }
        }
    }
    Write-Progress -Activity 'R.A.V.E.N. mailbox scan' -Completed

    if ($failures -gt 0) {
        Add-RavenFinding -Code 'MailboxScanIncomplete' -Subject "$failures query failure(s)" -Detail "Across $($mailboxes.Count) mailbox(es)."
    }

    $userFullAccess = @($delegation | Where-Object { $_.Right -eq 'Full Access' -and $_.Type -eq 'UserMailbox' })
    if ($userFullAccess.Count -gt 0) {
        Add-RavenFinding -Code 'FullAccessOnUserMailbox' -Subject "$(@($userFullAccess | Select-Object -ExpandProperty Mailbox -Unique).Count) mailbox(es)" -Detail "$($userFullAccess.Count) grant(s)"
    }

    Write-Info "Inbox rules scanned: $rulesScanned  |  flagged: $($flaggedRules.Count)"
    Write-Info "Forwarding mailboxes: $($forwarding.Count)  |  external: $(@($forwarding | Where-Object { $_.External }).Count)"

    return [PSCustomObject]@{
        MailboxCount  = $mailboxes.Count
        Forwarding    = @($forwarding)
        SmtpEnabled   = $smtpEnabled
        RulesScanned  = $rulesScanned
        FlaggedRules  = @($flaggedRules)
        Delegation    = @($delegation)
        Failures      = $failures
        DelegationRan = -not $SkipDelegation
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# CONSOLE + HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Show-RavenFindings {
    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) {
        Write-Ok "Nothing to report."
        return
    }
    $order = @{ 'Error' = 0; 'Warning' = 1; 'Info' = 2 }
    foreach ($f in ($Findings | Sort-Object { $order[$_.Severity] })) {
        $line = "$($f.Title)" + $(if ($f.Subject) { " -- $($f.Subject)" } else { '' })
        switch ($f.Severity) {
            'Error'   { Write-Fail $line }
            'Warning' { Write-Warn $line }
            default   { Write-Info $line }
        }
    }
}

function Build-RavenReport {
    param([object]$Audit, [object]$Verdict)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $scope      = if ($Audit.DnsOnly) { 'DNS only' } else { 'Full audit' }
    $tenant     = if ($Audit.Tenant -and $Audit.Tenant.OrgName) { $Audit.Tenant.OrgName } elseif ($Audit.Domains.Count -gt 0) { $Audit.Domains[0] } else { 'Unknown' }

    $errCount  = @($Findings | Where-Object { $_.Severity -eq 'Error' }).Count
    $warnCount = @($Findings | Where-Object { $_.Severity -eq 'Warning' }).Count

    $order = @{ 'Error' = 0; 'Warning' = 1; 'Info' = 2 }
    $fRows = New-Object System.Text.StringBuilder
    if ($Findings.Count -eq 0) {
        [void]$fRows.Append("<tr><td colspan='4'>Nothing to report.</td></tr>")
    } else {
        foreach ($f in ($Findings | Sort-Object { $order[$_.Severity] })) {
            $badge = Get-SeverityClass -Severity $f.Severity
            [void]$fRows.Append(
                "<tr><td><span class='tk-badge-$badge'>$(EscHtml $f.Severity)</span></td>" +
                "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)</td>" +
                "<td>$(EscHtml $f.Subject)" + $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
                "<td>$(EscHtml $f.Remedy)</td></tr>"
            )
        }
    }

    function _statusBadge { param([string]$Code, [string]$Text)
        $cls = if (-not $Code) { 'ok' } elseif ($RavenFindings[$Code].Severity -eq 'Error') { 'err' } elseif ($RavenFindings[$Code].Severity -eq 'Warning') { 'warn' } else { 'info' }
        return "<span class='tk-badge-$cls'>$(EscHtml $Text)</span>"
    }

    $dRows = New-Object System.Text.StringBuilder
    if ($Audit.DomainAuth.Count -eq 0) {
        [void]$dRows.Append("<tr><td colspan='4'>No domains checked.</td></tr>")
    } else {
        foreach ($d in $Audit.DomainAuth) {
            $dkimCls = if ($d.Dkim -eq 'Enabled' -or $d.Dkim -eq 'M365 selectors published') { 'ok' } elseif ($d.Dkim -eq 'Not enabled') { 'warn' } else { 'info' }
            [void]$dRows.Append(
                "<tr><td class='tk-mono'>$(EscHtml $d.Domain)</td>" +
                "<td>$(_statusBadge $d.SpfCode $d.SpfStatus)" + $(if ($d.SpfRecord) { "<br/><span class='tk-mono'>$(EscHtml $d.SpfRecord)</span>" } else { '' }) + "</td>" +
                "<td>$(_statusBadge $d.DmarcCode $d.DmarcStatus)" + $(if ($d.DmarcRecord) { "<br/><span class='tk-mono'>$(EscHtml $d.DmarcRecord)</span>" } else { '' }) + "</td>" +
                "<td><span class='tk-badge-$dkimCls'>$(EscHtml $d.Dkim)</span></td></tr>"
            )
        }
    }

    $sections = New-Object System.Text.StringBuilder
    $nav      = @('Findings', 'Email Authentication')

    if (-not $Audit.DnsOnly) {
        $mb = $Audit.Mailboxes
        $t  = $Audit.Tenant

        $fwRows = New-Object System.Text.StringBuilder
        if ($mb.Forwarding.Count -eq 0) {
            [void]$fwRows.Append("<tr><td colspan='5'>No mailbox-level forwarding.</td></tr>")
        } else {
            foreach ($r in ($mb.Forwarding | Sort-Object External -Descending)) {
                $cls = if ($r.External) { 'err' } else { 'info' }
                [void]$fwRows.Append(
                    "<tr><td class='tk-mono'>$(EscHtml $r.Mailbox)</td><td>$(EscHtml $r.Type)</td>" +
                    "<td class='tk-mono'>$(EscHtml $r.Target)</td><td><span class='tk-badge-$cls'>$(if ($r.External) { 'External' } else { 'Internal' })</span></td>" +
                    "<td>$(if ($r.KeepsCopy) { 'Yes' } else { '<span class=''tk-badge-warn''>No</span>' })</td></tr>"
                )
            }
        }

        $rRows = New-Object System.Text.StringBuilder
        if ($mb.FlaggedRules.Count -eq 0) {
            [void]$rRows.Append("<tr><td colspan='4'>No suspicious inbox rules among $($mb.RulesScanned) scanned.</td></tr>")
        } else {
            foreach ($r in ($mb.FlaggedRules | Sort-Object { $order[$_.Severity] })) {
                $cls = Get-SeverityClass -Severity $r.Severity
                [void]$rRows.Append(
                    "<tr><td class='tk-mono'>$(EscHtml $r.Mailbox)</td><td><span class='tk-badge-$cls'>$(EscHtml $r.Rule)</span></td>" +
                    "<td>$(if ($r.Enabled) { 'Enabled' } else { 'Disabled' })</td><td>$(EscHtml $r.Reasons)</td></tr>"
                )
            }
        }

        $pRows = New-Object System.Text.StringBuilder
        foreach ($p in $t.SpamPolicies) {
            $cls = if ($p.AutoForwardingMode -eq 'On') { 'warn' } else { 'ok' }
            [void]$pRows.Append("<tr><td>$(EscHtml $p.Name)$(if ($p.IsDefault) { ' (default)' } else { '' })</td><td><span class='tk-badge-$cls'>$(EscHtml $p.AutoForwardingMode)</span></td></tr>")
        }
        if ($t.SpamPolicies.Count -eq 0) { [void]$pRows.Append("<tr><td colspan='2'>Could not read outbound spam policies.</td></tr>") }

        $trRows = New-Object System.Text.StringBuilder
        if ($t.TransportRules.Count -eq 0) {
            [void]$trRows.Append("<tr><td colspan='3'>No transport rule sends mail outside the organisation.</td></tr>")
        } else {
            foreach ($r in $t.TransportRules) {
                [void]$trRows.Append("<tr><td>$(EscHtml $r.Name)</td><td>$(EscHtml $r.State)</td><td class='tk-mono'>$(EscHtml $r.External)</td></tr>")
            }
        }

        $delRows = New-Object System.Text.StringBuilder
        if (-not $mb.DelegationRan) {
            [void]$delRows.Append("<tr><td colspan='4'>Skipped (-SkipDelegation).</td></tr>")
        } elseif ($mb.Delegation.Count -eq 0) {
            [void]$delRows.Append("<tr><td colspan='4'>No Full Access or Send As delegation.</td></tr>")
        } else {
            foreach ($d in ($mb.Delegation | Sort-Object Type, Mailbox)) {
                [void]$delRows.Append("<tr><td class='tk-mono'>$(EscHtml $d.Mailbox)</td><td>$(EscHtml $d.Type)</td><td class='tk-mono'>$(EscHtml $d.Delegate)</td><td>$(EscHtml $d.Right)</td></tr>")
            }
        }

        $auditText = if ($null -eq $t.AuditDisabled) { 'Unknown' } elseif ($t.AuditDisabled) { 'DISABLED' } else { 'On' }
        $smtpText  = if ($null -eq $t.SmtpAuthDisabled) { 'Unknown' } elseif ($t.SmtpAuthDisabled) { 'Disabled' } else { 'ENABLED' }

        [void]$sections.Append(@"

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Mailbox Forwarding</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Mailbox</th><th>Type</th><th>Forwards to</th><th>Destination</th><th>Keeps copy</th></tr></thead>
        <tbody>$($fwRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> Inbox Rules</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Mailbox</th><th>Rule</th><th>State</th><th>Why it was flagged</th></tr></thead>
        <tbody>$($rRows.ToString())</tbody>
      </table>
      <div class="tk-info-box"><span class="tk-info-label">Scanned</span> $($mb.RulesScanned) rule(s) across $($mb.MailboxCount) mailbox(es). Rules hidden at the MAPI level do not appear through Get-InboxRule; a clean result here does not rule those out.</div>
    </div>
  </div>

  <div class="tk-section" id="s05">
    <div class="tk-section-title"><span class="tk-section-num">05</span> Tenant Settings</div>
    <div class="tk-card">
      <div class="tk-info-box">
        <span class="tk-info-label">Mailbox auditing</span> $(EscHtml $auditText)<br/>
        <span class="tk-info-label">SMTP AUTH (organisation)</span> $(EscHtml $smtpText)<br/>
        <span class="tk-info-label">SMTP AUTH enabled per mailbox</span> $($mb.SmtpEnabled.Count)
      </div>
      <table class="tk-table">
        <thead><tr><th>Outbound spam policy</th><th>AutoForwardingMode</th></tr></thead>
        <tbody>$($pRows.ToString())</tbody>
      </table>
      <table class="tk-table">
        <thead><tr><th>Transport rule</th><th>State</th><th>External recipients</th></tr></thead>
        <tbody>$($trRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s06">
    <div class="tk-section-title"><span class="tk-section-num">06</span> Delegation</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Mailbox</th><th>Type</th><th>Delegate</th><th>Right</th></tr></thead>
        <tbody>$($delRows.ToString())</tbody>
      </table>
    </div>
  </div>
"@)
        $nav += @('Mailbox Forwarding', 'Inbox Rules', 'Tenant Settings', 'Delegation')
    }

    $meta = [ordered]@{
        'Tenant'    = $tenant
        'Generated' = $reportDate
        'Scope'     = $scope
        'Verdict'   = $Verdict.Verdict
    }
    if (-not $Audit.DnsOnly) {
        $meta['Signed in as'] = $Audit.ConnectedAs
        $meta['Mailboxes']    = $Audit.Mailboxes.MailboxCount
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'R.A.V.E.N. Exchange Online Mailbox Security Report' `
        -ScriptName 'R.A.V.E.N.' `
        -Subtitle   "${orgPrefix}Exchange Online Mailbox Security -- $tenant" `
        -MetaItems  $meta `
        -NavItems   $nav

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'R.A.V.E.N. v6.2'

    $mbCards = ''
    if (-not $Audit.DnsOnly) {
        $extFw  = @($Audit.Mailboxes.Forwarding | Where-Object { $_.External }).Count
        $badRul = @($Audit.Mailboxes.FlaggedRules | Where-Object { $_.Severity -eq 'Error' }).Count
        $mbCards = @"
    <div class="tk-summary-card $(if ($extFw -gt 0) { 'err' } else { 'ok' })"><div class="tk-summary-num">$extFw</div><div class="tk-summary-lbl">External Forwards</div></div>
    <div class="tk-summary-card $(if ($badRul -gt 0) { 'err' } elseif ($Audit.Mailboxes.FlaggedRules.Count -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Audit.Mailboxes.FlaggedRules.Count)</div><div class="tk-summary-lbl">Flagged Inbox Rules</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Audit.Mailboxes.MailboxCount)</div><div class="tk-summary-lbl">Mailboxes</div></div>
"@
    }

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Verdict</div></div>
    <div class="tk-summary-card $(if ($errCount -gt 0) { 'err' } else { 'ok' })"><div class="tk-summary-num">$errCount</div><div class="tk-summary-lbl">High-risk Findings</div></div>
    <div class="tk-summary-card $(if ($warnCount -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$warnCount</div><div class="tk-summary-lbl">Warnings</div></div>
$mbCards
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Audit.DomainAuth.Count)</div><div class="tk-summary-lbl">Domains Checked</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Severity</th><th>Finding</th><th>Where</th><th>Remedy</th></tr></thead>
        <tbody>$($fRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Email Authentication</div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Domain</th><th>SPF</th><th>DMARC</th><th>DKIM</th></tr></thead>
        <tbody>$($dRows.ToString())</tbody>
      </table>
    </div>
  </div>
$($sections.ToString())
"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# ORCHESTRATION
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-RavenRun {
    param([bool]$DnsOnlyMode, [string[]]$Domains)

    $Findings.Clear()

    $audit = [PSCustomObject]@{
        DnsOnly     = $DnsOnlyMode
        ConnectedAs = ''
        Domains     = @()
        DomainAuth  = @()
        Tenant      = $null
        Mailboxes   = $null
    }

    if ($DnsOnlyMode) {
        $targets = @($Domains | ForEach-Object { "$_".Trim().ToLowerInvariant() } | Where-Object { $_ } | Select-Object -Unique)
        if ($targets.Count -eq 0) {
            Write-Fail "DNS-only mode needs at least one domain (-Domain contoso.com)."
            return
        }
        $audit.Domains = $targets
        Write-Section "EMAIL AUTHENTICATION (DNS)"
        $audit.DomainAuth = Get-RavenDomainAuth -Domains $targets -DkimConfig $null
    } else {
        if (-not (Install-ExchangeModule)) {
            if ($Unattended) { exit 1 }
            return
        }
        $who = Connect-RavenExchange
        if (-not $who) {
            if ($Unattended) { exit 1 }
            return
        }
        $audit.ConnectedAs = $who

        Write-Step "Reading accepted domains..."
        $accepted = @()
        try {
            $accepted = @(Get-AcceptedDomain -ErrorAction Stop | ForEach-Object { "$($_.DomainName)".ToLowerInvariant() })
        } catch {
            Write-Fail "Get-AcceptedDomain failed: $($_.Exception.Message)"
        }
        if ($accepted.Count -eq 0) {
            Write-Fail "No accepted domains returned -- cannot tell internal from external addresses. Stopping."
            Write-TKError -ScriptName 'raven' -Message 'Get-AcceptedDomain returned nothing.' -Category 'EXO Query'
            if ($Unattended) { exit 1 }
            return
        }
        Write-Info "Accepted domains: $($accepted -join ', ')"

        $audit.Tenant    = Get-RavenTenantSettings -InternalDomains $accepted
        $audit.Mailboxes = Get-RavenMailboxScan    -InternalDomains $accepted

        # DNS: the named domains if given, else every accepted domain that is
        # not a Microsoft-managed onmicrosoft.com name.
        $targets = if ($Domains) {
            @($Domains | ForEach-Object { "$_".Trim().ToLowerInvariant() } | Where-Object { $_ } | Select-Object -Unique)
        } else {
            @($accepted | Where-Object { $_ -notlike '*.onmicrosoft.com' })
        }
        $audit.Domains = $targets

        $dkim = @{}
        try {
            foreach ($cfg in @(Get-DkimSigningConfig -ErrorAction Stop)) { $dkim["$($cfg.Domain)".ToLowerInvariant()] = [bool]$cfg.Enabled }
        } catch {
            Write-Warn "Get-DkimSigningConfig failed: $($_.Exception.Message) -- falling back to DNS selector checks."
            $dkim = $null
        }

        Write-Section "EMAIL AUTHENTICATION (DNS)"
        $audit.DomainAuth = Get-RavenDomainAuth -Domains $targets -DkimConfig $dkim
    }

    Show-RavenFindings

    $verdict = Get-RavenVerdict -FindingList @($Findings)
    Write-Section "VERDICT"
    switch ($verdict.Class) {
        'err'   { Write-Fail $verdict.Verdict }
        'warn'  { Write-Warn $verdict.Verdict }
        default { Write-Ok   $verdict.Verdict }
    }

    Add-TKNote -Text ("RAVEN {0} audit: verdict {1} ({2} finding(s))." -f $(if ($DnsOnlyMode) { 'DNS-only' } else { 'full' }), $verdict.Verdict, $Findings.Count) -Category 'Info' -ScriptName 'raven'

    Write-Step "Generating HTML report..."
    $html      = Build-RavenReport -Audit $audit -Verdict $verdict
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "RAVEN_${timestamp}.html"

    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'raven' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-RavenBanner
    if ($DnsOnly -and -not $Domain) {
        Write-Fail "-DnsOnly needs -Domain (for example: -Domain contoso.com)."
        exit 1
    }
    Invoke-RavenRun -DnsOnlyMode ([bool]$DnsOnly) -Domains $Domain
} else {
    $choice = ''

    do {
        Show-RavenBanner

        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] Full audit  -  sign in to Exchange Online, audit mailboxes + DNS" -ForegroundColor $C.Info
        Write-Host "  [2] DNS only  -  SPF / DKIM / DMARC for domains you name, no sign-in" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' { Invoke-RavenRun -DnsOnlyMode $false -Domains $Domain }
            '2' {
                $names = $Domain
                if (-not $names) {
                    Write-Host -NoNewline "  Domain(s), comma-separated: " -ForegroundColor $C.Header
                    $names = @((Read-Host) -split '[,;\s]+' | Where-Object { $_ })
                }
                Invoke-RavenRun -DnsOnlyMode $true -Domains $names
            }
            'Q' {
                Write-Host ""
                Write-Host "  Closing R.A.V.E.N." -ForegroundColor $C.Header
                Write-Host ""
            }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1, 2 or Q." -ForegroundColor $C.Warning
                Start-Sleep -Seconds 1
            }
        }

        if ($choice -ne 'Q') {
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

    } while ($choice -ne 'Q')
}

# The note buffer ends with this run, so save it now -- one file per run.
$null = Save-TKNoteReport -Prefix 'RAVEN' -Directory (Resolve-LogDirectory -FallbackPath $ScriptPath)

if ($Transcript) { Stop-TKTranscript }
