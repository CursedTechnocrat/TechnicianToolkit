# wyrm.ps1 - W.Y.R.M. — Whole-drive encrYption & Recovery Manager
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
    W.Y.R.M. — Whole-drive encrYption & Recovery Manager
    BitLocker Drive Encryption Tool for PowerShell 5.1+

.DESCRIPTION
    Manages BitLocker on the local machine: show drive status, enable or
    disable encryption, suspend or resume protection, show recovery keys, back
    them up to Active Directory or Entra ID, and export a status + recovery-key
    report.

    Every change goes through manage-bde.exe and every read through the
    Win32_EncryptableVolume CIM class. Neither depends on the BitLocker
    PowerShell module, so the tool behaves the same under Windows PowerShell
    5.1 and PowerShell 7.

    Enable uses one protector set: TPM + recovery password on the operating
    system drive, recovery password + auto-unlock on any other drive.

.USAGE
    PS C:\> .\wyrm.ps1                                           # Must be run as Administrator
    PS C:\> .\wyrm.ps1 -WhatIf                                   # Preview actions without making changes
    PS C:\> .\wyrm.ps1 -Unattended -Action Status                # Show drive status and exit
    PS C:\> .\wyrm.ps1 -Unattended -Action Enable -Drive C       # Encrypt C: (TPM + recovery password)
    PS C:\> .\wyrm.ps1 -Unattended -Action Disable -Drive C      # Decrypt C:
    PS C:\> .\wyrm.ps1 -Unattended -Action Suspend -Drive C      # Suspend BitLocker on C: for one reboot
    PS C:\> .\wyrm.ps1 -Unattended -Action BackupAD -Drive C     # Backup recovery key to AD
    PS C:\> .\wyrm.ps1 -Unattended -Action Export                # Export status + recovery keys to HTML
    PS C:\> .\wyrm.ps1 -Unattended -Action Export -OutputPath D:\Reports

.NOTES
    Version : 6.2

    Credits : Thanks to Steve the Killer for help and letting me use his
              script BERET: https://tools.thekiller.net/killer-scripts

#>

param(
    [switch]$Unattended,
    [switch]$WhatIf,
    [ValidateSet('Status','Enable','Disable','Suspend','Resume','BackupAD','BackupEntraID','Export')]
    [string]$Action = "Status",
    [ValidatePattern('^[A-Za-z]:?$')]
    [string]$Drive  = "C",
    [string]$OutputPath,
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

function Assert-AdminPrivilege {
    if (-not (Test-IsAdmin)) {
        Write-Host "  This script must be run as Administrator." -ForegroundColor Red
        exit 1
    }
}
#endregion TK-SHARED
Assert-AdminPrivilege

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $PSScriptRoot) }

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
# BITLOCKER PLUMBING
# ─────────────────────────────────────────────────────────────────────────────

# Sysnative reaches the real System32 from a 32-bit host, where manage-bde
# would otherwise not be found at all.
$ManageBde = Join-Path $env:windir 'Sysnative\manage-bde.exe'
if (-not (Test-Path $ManageBde)) { $ManageBde = Join-Path $env:windir 'System32\manage-bde.exe' }

$BdeNamespace = 'root/CIMV2/Security/MicrosoftVolumeEncryption'

$ConversionStatusNames = @{
    0 = 'FullyDecrypted'; 1 = 'FullyEncrypted'; 2 = 'EncryptionInProgress'
    3 = 'DecryptionInProgress'; 4 = 'EncryptionPaused'; 5 = 'DecryptionPaused'
}
$ProtectorTypeNames = @{
    0 = 'Unknown'; 1 = 'Tpm'; 2 = 'ExternalKey'; 3 = 'RecoveryPassword'; 4 = 'TpmPin'
    5 = 'TpmStartupKey'; 6 = 'TpmPinStartupKey'; 7 = 'PublicKey'; 8 = 'Password'
    9 = 'TpmNetworkKey'; 10 = 'AdAccountOrGroup'
}

# Runs manage-bde, echoes its output, and returns $true on exit code 0. Under
# -WhatIf it only prints the command. Every change the tool makes goes through
# here.
function Invoke-ManageBde {
    param([Parameter(Mandatory)][string[]]$Arguments)

    if ($WhatIf) {
        Write-Host "  [~] Would run: manage-bde $($Arguments -join ' ')" -ForegroundColor Cyan
        return $true
    }

    $output = & $ManageBde @Arguments 2>&1
    $exit   = $LASTEXITCODE
    foreach ($line in $output) {
        if ("$line".Trim()) { Write-Host "      $line" -ForegroundColor $ColorSchema.Info }
    }
    if ($exit -ne 0) {
        Write-TKError -ScriptName 'wyrm' -Message "manage-bde $($Arguments -join ' ') exited $exit" -Category 'BitLocker'
    }
    return ($exit -eq 0)
}

# Returns one object per lettered volume: MountPoint, Status, Protection,
# Percent and Protectors (Id / Type / RecoveryPassword).
function Get-WyrmVolume {
    param([string]$MountPoint)

    $volumes = Get-CimInstance -Namespace $BdeNamespace -ClassName Win32_EncryptableVolume -ErrorAction Stop |
        Where-Object { $_.DriveLetter -and (-not $MountPoint -or $_.DriveLetter -eq $MountPoint) }

    foreach ($v in $volumes) {
        $conv = Invoke-CimMethod -InputObject $v -MethodName GetConversionStatus
        $ids  = (Invoke-CimMethod -InputObject $v -MethodName GetKeyProtectors).VolumeKeyProtectorID

        $protectors = foreach ($id in $ids) {
            $idArg    = @{ VolumeKeyProtectorID = $id }
            $typeCode = [int](Invoke-CimMethod -InputObject $v -MethodName GetKeyProtectorType -Arguments $idArg).KeyProtectorType
            $password = $null
            if ($typeCode -eq 3) {
                $password = (Invoke-CimMethod -InputObject $v -MethodName GetKeyProtectorNumericalPassword -Arguments $idArg).NumericalPassword
            }
            [PSCustomObject]@{ Id = $id; Type = $ProtectorTypeNames[$typeCode]; RecoveryPassword = $password }
        }

        # A locked volume fails GetConversionStatus; its zeroed output must not
        # read as FullyDecrypted.
        $status = if ($conv.ReturnValue -eq 0) { $ConversionStatusNames[[int]$conv.ConversionStatus] }
        [PSCustomObject]@{
            MountPoint = $v.DriveLetter
            Status     = if ($status) { $status } else { 'Unknown' }
            Protection = switch ([int]$v.ProtectionStatus) { 0 { 'Off' } 1 { 'On' } default { 'Unknown' } }
            Percent    = [int]$conv.EncryptionPercentage
            Protectors = @($protectors)
        }
    }
}

# Returns the volume for one drive, or $null after saying why.
function Get-TargetVolume {
    param([Parameter(Mandatory)][string]$MountPoint)
    $vol = Get-WyrmVolume -MountPoint $MountPoint | Select-Object -First 1
    if (-not $vol) { Write-Fail "Drive $MountPoint was not found, or BitLocker cannot manage it." }
    return $vol
}

# ─────────────────────────────────────────────────────────────────────────────
# ACTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Show-DriveStatus {
    Write-Section 'DRIVE ENCRYPTION STATUS'
    foreach ($vol in Get-WyrmVolume) {
        $statusColor = switch ($vol.Status) {
            'FullyEncrypted' { $ColorSchema.Success }
            'FullyDecrypted' { $ColorSchema.Warning }
            default          { $ColorSchema.Progress }
        }
        $protColor = if ($vol.Protection -eq 'On') { $ColorSchema.Success } else { $ColorSchema.Warning }
        $types     = if ($vol.Protectors.Count) { ($vol.Protectors.Type) -join ', ' } else { 'None' }

        Write-Host "  Drive $($vol.MountPoint)" -ForegroundColor $ColorSchema.Header
        Write-Host ("    Status      : {0}" -f $vol.Status) -ForegroundColor $statusColor
        Write-Host ("    Protection  : {0}" -f $vol.Protection) -ForegroundColor $protColor
        Write-Host ("    Encrypted   : {0}%" -f $vol.Percent) -ForegroundColor $ColorSchema.Info
        Write-Host ("    Protectors  : {0}" -f $types) -ForegroundColor $ColorSchema.Info
        Write-Host ""
    }
}

function Show-RecoveryKey {
    param([Parameter(Mandatory)][string]$MountPoint)

    $vol = Get-TargetVolume -MountPoint $MountPoint
    if (-not $vol) { return }

    $keys = @($vol.Protectors | Where-Object { $_.Type -eq 'RecoveryPassword' })
    if (-not $keys) {
        Write-Warn "No recovery password on $MountPoint."
        return
    }
    Write-Host ""
    Write-Host "  RECOVERY KEY(S) FOR $MountPoint" -ForegroundColor $ColorSchema.Warning
    foreach ($key in $keys) {
        Write-Host "    ID  : $($key.Id)" -ForegroundColor $ColorSchema.Warning
        Write-Host "    Key : $($key.RecoveryPassword)" -ForegroundColor $ColorSchema.Warning
    }
    Write-Host ""
}

# Encrypts a decrypted drive, or turns protection back on for an encrypted one.
# Either way the drive ends up with a recovery password plus TPM (OS drive) or
# auto-unlock (any other drive).
function Enable-DriveEncryption {
    param([Parameter(Mandatory)][string]$MountPoint)

    Write-Section "ENABLE BITLOCKER ON $MountPoint"
    $vol = Get-TargetVolume -MountPoint $MountPoint
    if (-not $vol) { return }

    if ($vol.Status -eq 'Unknown') {
        Write-Fail "$MountPoint is locked or unreadable. Unlock it first (manage-bde -unlock)."
        return
    }
    if ($vol.Status -in 'DecryptionInProgress', 'DecryptionPaused') {
        Write-Warn "$MountPoint is still decrypting ($($vol.Percent)% encrypted). Run Enable again once it reads FullyDecrypted."
        return
    }
    if ($vol.Status -ne 'FullyDecrypted' -and $vol.Protection -eq 'On') {
        Write-Ok "BitLocker is already on for $MountPoint - nothing to do."
        Show-RecoveryKey -MountPoint $MountPoint
        return
    }

    $isOsDrive = $MountPoint -eq $env:SystemDrive
    $types     = @($vol.Protectors.Type)

    if ($types -notcontains 'RecoveryPassword') {
        Write-Step 'Adding a recovery password...'
        if (-not (Invoke-ManageBde '-protectors', '-add', $MountPoint, '-RecoveryPassword')) {
            Write-Fail 'Could not add a recovery password. Nothing else was changed.'
            return
        }
    }
    if ($isOsDrive -and -not ($types | Where-Object { $_ -like 'Tpm*' })) {
        Write-Step 'Adding a TPM protector...'
        if (-not (Invoke-ManageBde '-protectors', '-add', $MountPoint, '-TPM')) {
            Write-Fail 'Could not add a TPM protector. Check that the TPM is present and ready (tpm.msc).'
            return
        }
    }

    if ($vol.Status -eq 'FullyDecrypted') {
        Write-Step "Starting XTS-AES 256 encryption on $MountPoint..."
        $started = Invoke-ManageBde '-on', $MountPoint, '-EncryptionMethod', 'XtsAes256', '-UsedSpaceOnly', '-SkipHardwareTest'
        if (-not $started) {
            Write-Step 'Retrying as full-volume encryption...'
            $started = Invoke-ManageBde '-on', $MountPoint, '-EncryptionMethod', 'XtsAes256', '-SkipHardwareTest'
        }
        if (-not $started) {
            Write-Fail "Encryption did not start on $MountPoint."
            return
        }
    } else {
        # Encrypted but protection is off: suspended, or an OEM device-encryption
        # volume still holding a clear key. -enable removes the clear key. If it
        # refuses, the protectors are half-suspended; a clean -disable first puts
        # them back into a state -enable accepts. -disable never decrypts.
        Write-Step "Turning protection on for $MountPoint..."
        $on = Invoke-ManageBde '-protectors', '-enable', $MountPoint
        if (-not $on) {
            Write-Step 'Re-suspending cleanly and retrying...'
            $on = (Invoke-ManageBde '-protectors', '-disable', $MountPoint) -and
                  (Invoke-ManageBde '-protectors', '-enable', $MountPoint)
        }
        if (-not $on) {
            Write-Fail "Protection is still off on $MountPoint."
            return
        }
    }

    if (-not $isOsDrive) {
        Write-Step "Enabling auto-unlock so $MountPoint opens at sign-in..."
        if (-not (Invoke-ManageBde '-autounlock', '-enable', $MountPoint)) {
            Write-Warn 'Auto-unlock was not enabled (the OS drive must be encrypted first). The drive will ask for its recovery password after a reboot.'
        }
    }

    Write-Ok "BitLocker is enabled on $MountPoint."
    if (-not $WhatIf) {
        Write-Warn 'Save the recovery key below before closing this window.'
        Show-RecoveryKey -MountPoint $MountPoint
    }
}

function Disable-DriveEncryption {
    param([Parameter(Mandatory)][string]$MountPoint)
    Write-Section "DISABLE BITLOCKER ON $MountPoint"
    if (Invoke-ManageBde '-off', $MountPoint) {
        Write-Ok "Decryption started on $MountPoint. It runs in the background."
    }
}

function Suspend-DriveProtection {
    param([Parameter(Mandatory)][string]$MountPoint)
    Write-Section "SUSPEND BITLOCKER ON $MountPoint"
    Write-Info 'Use before BIOS / firmware updates. Protection resumes after one reboot.'
    if (Invoke-ManageBde '-protectors', '-disable', $MountPoint, '-RebootCount', '1') {
        Write-Ok "BitLocker suspended on $MountPoint until the next reboot."
    }
}

function Resume-DriveProtection {
    param([Parameter(Mandatory)][string]$MountPoint)
    Write-Section "RESUME BITLOCKER ON $MountPoint"
    if (Invoke-ManageBde '-protectors', '-enable', $MountPoint) {
        Write-Ok "BitLocker protection resumed on $MountPoint."
    }
}

function Backup-RecoveryKey {
    param(
        [Parameter(Mandatory)][string]$MountPoint,
        [Parameter(Mandatory)][ValidateSet('AD','EntraID')][string]$Target
    )
    Write-Section "BACK UP RECOVERY KEY FOR $MountPoint TO $Target"
    $vol = Get-TargetVolume -MountPoint $MountPoint
    if (-not $vol) { return }

    $keys = @($vol.Protectors | Where-Object { $_.Type -eq 'RecoveryPassword' })
    if (-not $keys) {
        Write-Fail "No recovery password on $MountPoint. Enable BitLocker first."
        return
    }
    $verb = if ($Target -eq 'AD') { '-adbackup' } else { '-aadbackup' }
    foreach ($key in $keys) {
        if (Invoke-ManageBde '-protectors', $verb, $MountPoint, '-id', $key.Id) {
            Write-Ok "Recovery key $($key.Id) backed up to $Target."
        } else {
            Write-Warn "Backup failed. Check the machine is joined to $Target and can reach it."
        }
    }
}

function Export-EncryptionReport {
    Write-Section 'EXPORT ENCRYPTION REPORT'
    Write-Warn 'The report contains the recovery passwords. Store it somewhere secure, not on the encrypted drive.'

    if ($WhatIf) {
        Write-Host "  [~] Would write an HTML report with drive status and recovery keys." -ForegroundColor Cyan
        return
    }

    $volumes = @(Get-WyrmVolume)

    $reportDir = if ($OutputPath) { $OutputPath } else { Resolve-LogDirectory -FallbackPath $PSScriptRoot }
    if (-not (Test-Path $reportDir)) { New-Item -ItemType Directory -Path $reportDir -Force | Out-Null }
    $reportPath = Join-Path $reportDir "WYRM_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

    $cfg      = Get-TKConfig
    $subtitle = if ($cfg.OrgName) { "$($cfg.OrgName) — $env:COMPUTERNAME" } else { $env:COMPUTERNAME }

    $statusRows = foreach ($vol in $volumes) {
        $badge = if ($vol.Protection -eq 'On') { "<span class='tk-badge-ok'>On</span>" } else { "<span class='tk-badge-warn'>$(EscHtml $vol.Protection)</span>" }
        $types = if ($vol.Protectors.Count) { ($vol.Protectors.Type) -join ', ' } else { 'None' }
        "<tr><td class='tk-mono'>$(EscHtml $vol.MountPoint)</td><td>$(EscHtml $vol.Status)</td><td>$badge</td><td>$($vol.Percent)%</td><td>$(EscHtml $types)</td></tr>"
    }
    $keyRows = foreach ($vol in $volumes) {
        foreach ($key in ($vol.Protectors | Where-Object { $_.Type -eq 'RecoveryPassword' })) {
            "<tr><td class='tk-mono'>$(EscHtml $vol.MountPoint)</td><td class='tk-mono'>$(EscHtml $key.Id)</td><td class='tk-mono'>$(EscHtml $key.RecoveryPassword)</td></tr>"
        }
    }
    if (-not $keyRows) { $keyRows = "<tr><td colspan='3'>No recovery passwords on any drive.</td></tr>" }

    $html  = Get-TKHtmlHead -Title 'BitLocker Encryption Report' -ScriptName 'W.Y.R.M.' `
                 -Subtitle $subtitle `
                 -MetaItems ([ordered]@{
                     'Generated' = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                     'Run As'    = "$env:USERDOMAIN\$env:USERNAME"
                     'Volumes'   = $volumes.Count
                 }) `
                 -NavItems @('Drive Status', 'Recovery Keys')
    $html += @"
<div class="tk-section" id="s01">
  <div class="tk-section-title"><span class="tk-section-num">01</span> Drive Status</div>
  <div class="tk-info-box"><span class="tk-info-label">SENSITIVE</span> This document contains BitLocker recovery passwords. Anyone who reads it can unlock the listed drives.</div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>Drive</th><th>Volume Status</th><th>Protection</th><th>Encrypted</th><th>Key Protectors</th></tr></thead>
      <tbody>$($statusRows -join "`n")</tbody>
    </table>
  </div>
</div>
<div class="tk-section" id="s02">
  <div class="tk-section-title"><span class="tk-section-num">02</span> Recovery Keys</div>
  <div class="tk-card">
    <table class="tk-table">
      <thead><tr><th>Drive</th><th>Key Protector ID</th><th>Recovery Password</th></tr></thead>
      <tbody>$($keyRows -join "`n")</tbody>
    </table>
  </div>
</div>
"@
    $html += Get-TKHtmlFoot -ScriptName 'W.Y.R.M. v6.2'

    $html | Out-File -FilePath $reportPath -Encoding UTF8
    Show-TKReportResult -Path $reportPath -Unattended:$Unattended
}

# Runs one action against one drive. Shared by the unattended path and the menu.
function Invoke-WyrmAction {
    param([Parameter(Mandatory)][string]$Name, [string]$MountPoint)
    switch ($Name) {
        'Status'        { Show-DriveStatus }
        'Enable'        { Enable-DriveEncryption  -MountPoint $MountPoint }
        'Disable'       { Disable-DriveEncryption -MountPoint $MountPoint }
        'Suspend'       { Suspend-DriveProtection -MountPoint $MountPoint }
        'Resume'        { Resume-DriveProtection  -MountPoint $MountPoint }
        'ShowKeys'      { Show-RecoveryKey        -MountPoint $MountPoint }
        'BackupAD'      { Backup-RecoveryKey      -MountPoint $MountPoint -Target AD }
        'BackupEntraID' { Backup-RecoveryKey      -MountPoint $MountPoint -Target EntraID }
        'Export'        { Export-EncryptionReport }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

$DefaultMount = "$($Drive.ToUpper().TrimEnd(':')):"

try {
    if ($Unattended) {
        if ($WhatIf) { Write-Host "  [~] DRY RUN - no changes will be made." -ForegroundColor Cyan }
        Invoke-WyrmAction -Name $Action -MountPoint $DefaultMount
    } else {
        $menu = [ordered]@{
            '1' = @('Enable',        'Enable BitLocker')
            '2' = @('Disable',       'Disable BitLocker (decrypt)')
            '3' = @('Suspend',       'Suspend protection for one reboot')
            '4' = @('Resume',        'Resume protection')
            '5' = @('ShowKeys',      'Show recovery keys')
            '6' = @('BackupAD',      'Back up recovery key to Active Directory')
            '7' = @('BackupEntraID', 'Back up recovery key to Entra ID')
            '8' = @('Export',        'Export report (status + recovery keys)')
        }
        $destructive = @('Disable', 'Suspend')

        while ($true) {
            Clear-Host
            Write-Host ""
            Write-Host "  W.Y.R.M. — Whole-drive encrYption & Recovery Manager" -ForegroundColor $ColorSchema.Header
            if ($WhatIf) { Write-Host "  [~] DRY RUN - no changes will be made." -ForegroundColor Cyan }
            Show-DriveStatus

            foreach ($key in $menu.Keys) { Write-Host "  [$key] $($menu[$key][1])" -ForegroundColor $ColorSchema.Info }
            Write-Host "  [Q] Quit" -ForegroundColor $ColorSchema.Info
            Write-Host ""
            $choice = (Read-Host '  Selection').Trim().ToUpper()
            if ($choice -eq 'Q') { break }
            if (-not $menu.Contains($choice)) { continue }

            $name  = $menu[$choice][0]
            $mount = $DefaultMount
            if ($name -ne 'Export') {
                $letter = (Read-Host "  Drive letter [$($DefaultMount.TrimEnd(':'))]").Trim().TrimEnd(':').ToUpper()
                if ($letter -match '^[A-Z]$') { $mount = "${letter}:" }
            }
            $confirmed = $WhatIf -or ($name -notin $destructive) -or
                         ((Read-Host "  $name BitLocker on $mount? (Y/N)").Trim().ToUpper() -eq 'Y')
            if ($confirmed) { Invoke-WyrmAction -Name $name -MountPoint $mount }

            Write-Host ""
            Read-Host '  Press Enter to return to the menu' | Out-Null
        }
    }
} catch {
    Write-Fail "BitLocker could not be read: $($_.Exception.Message)"
    Write-Info 'BitLocker needs a Pro, Enterprise or Education edition of Windows.'
    Write-TKError -ScriptName 'wyrm' -Message $_.Exception.Message -Category 'BitLocker'
}

if ($Transcript) { Stop-TKTranscript }
