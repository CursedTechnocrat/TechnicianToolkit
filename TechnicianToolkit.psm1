#Requires -Version 5.1
<#
.SYNOPSIS
    Shared helpers for the TechnicianToolkit suite.
.DESCRIPTION
    Provides logging, HTML utility, and privilege management functions shared
    across all TechnicianToolkit scripts. Import at the top of each script:

        Import-Module "$PSScriptRoot\TechnicianToolkit.psm1" -Force
#>

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

#region ── Logging Helpers ───────────────────────────────────────────────────

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host ("  " + ("─" * 62)) -ForegroundColor Cyan
    Write-Host ""
}

function Write-Step { param([string]$Msg) Write-Host ("  [*] {0}" -f $Msg) -ForegroundColor Magenta }
function Write-Ok   { param([string]$Msg) Write-Host ("  [+] {0}" -f $Msg) -ForegroundColor Green   }
function Write-Warn { param([string]$Msg) Write-Host ("  [!] {0}" -f $Msg) -ForegroundColor Yellow  }
function Write-Fail { param([string]$Msg) Write-Host ("  [-] {0}" -f $Msg) -ForegroundColor Red     }
function Write-Info { param([string]$Msg) Write-Host ("      {0}" -f $Msg) -ForegroundColor Gray    }

#endregion

#region ── Error Telemetry ───────────────────────────────────────────────────

function Write-TKError {
    <#
    .SYNOPSIS
        Logs a structured error to the central toolkit error log and optionally
        posts to a Teams incoming webhook configured as TeamsWebhook in config.json.
    .EXAMPLE
        Write-TKError -ScriptName 'sigil' -Message $_.Exception.Message -Category 'Registry'
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

    # Optional Teams webhook notification
    try {
        $cfg = Get-TKConfig
        $webhook = $cfg.TeamsWebhook
        if (-not [string]::IsNullOrWhiteSpace($webhook)) {
            $card = @{
                '@type'      = 'MessageCard'
                '@context'   = 'http://schema.org/extensions'
                themeColor   = 'FF0000'
                summary      = "TechnicianToolkit error in $ScriptName"
                sections     = @(@{
                    activityTitle    = "TechnicianToolkit — $ScriptName [$Category]"
                    activitySubtitle = "$($entry.Host) / $($entry.User)  •  $($entry.Timestamp)"
                    activityText     = $Message
                })
            } | ConvertTo-Json -Depth 5

            Invoke-RestMethod -Uri $webhook -Method Post -Body $card `
                -ContentType 'application/json' -ErrorAction Stop | Out-Null
        }
    }
    catch { <# webhook failures are silent — never interrupt the caller #> }
}

#endregion

#region ── HTML Utilities ────────────────────────────────────────────────────

function EscHtml {
    param([string]$s)
    if (-not $s) { return '' }
    $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

#endregion

#region ── Config Helpers ────────────────────────────────────────────────────

function Get-TKConfig {
    <#
    .SYNOPSIS
        Returns the toolkit configuration as a PSCustomObject.
        Reads config.json from the module directory; missing keys are filled
        with empty-string defaults so callers never receive null.
    #>
    $configPath = Join-Path $PSScriptRoot 'config.json'

    $defaults = [PSCustomObject]@{
        OrgName       = ''
        LogDirectory  = ''
        TeamsWebhook  = ''
        Archive       = [PSCustomObject]@{ DefaultDestination   = '' }
        Phantom       = [PSCustomObject]@{ DefaultDestination   = '' }
        Covenant      = [PSCustomObject]@{ DefaultTimezone = ''; DefaultLocalAdminUser = '' }
    }

    if (-not (Test-Path $configPath)) { return $defaults }

    try {
        $raw = Get-Content $configPath -Raw -ErrorAction Stop | ConvertFrom-Json

        # Ensure top-level keys exist
        foreach ($key in ($defaults | Get-Member -MemberType NoteProperty).Name) {
            if ($null -eq $raw.$key) {
                $raw | Add-Member -NotePropertyName $key -NotePropertyValue $defaults.$key -Force
            }
        }
        # Ensure nested keys exist
        foreach ($section in @('Archive','Phantom','Covenant')) {
            if ($raw.$section -isnot [PSCustomObject]) {
                $raw | Add-Member -NotePropertyName $section -NotePropertyValue $defaults.$section -Force
            } else {
                foreach ($key in ($defaults.$section | Get-Member -MemberType NoteProperty).Name) {
                    if ($null -eq $raw.$section.$key) {
                        $raw.$section | Add-Member -NotePropertyName $key -NotePropertyValue '' -Force
                    }
                }
            }
        }
        return $raw
    }
    catch { return $defaults }
}

function Set-TKConfig {
    <#
    .SYNOPSIS
        Writes a single value to config.json.
    .EXAMPLE
        Set-TKConfig -Key 'OrgName'             -Value 'Contoso'
        Set-TKConfig -Key 'DefaultDestination'  -Value '\\srv\backups' -Section 'Archive'
    #>
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)][string]$Value,
        [string]$Section = ''
    )

    $configPath = Join-Path $PSScriptRoot 'config.json'
    $cfg = if (Test-Path $configPath) {
        Get-Content $configPath -Raw | ConvertFrom-Json
    } else {
        [PSCustomObject]@{}
    }

    if ($Section) {
        if ($null -eq $cfg.$Section) {
            $cfg | Add-Member -NotePropertyName $Section -NotePropertyValue ([PSCustomObject]@{}) -Force
        }
        $cfg.$Section | Add-Member -NotePropertyName $Key -NotePropertyValue $Value -Force
    } else {
        $cfg | Add-Member -NotePropertyName $Key -NotePropertyValue $Value -Force
    }

    $cfg | ConvertTo-Json -Depth 4 | Set-Content $configPath -Encoding UTF8
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

#endregion

#region ── Privilege Management ──────────────────────────────────────────────

function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Assert-AdminPrivilege {
    if (-not (Test-IsAdmin)) {
        Write-Host "  This script must be run as Administrator." -ForegroundColor Red
        exit 1
    }
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

#endregion

Export-ModuleMember -Function Write-Section, Write-Step, Write-Ok, Write-Warn, Write-Fail, Write-Info, EscHtml, Test-IsAdmin, Assert-AdminPrivilege, Invoke-AdminElevation, Get-TKConfig, Set-TKConfig, Resolve-LogDirectory, Start-TKTranscript, Stop-TKTranscript, Write-TKError
