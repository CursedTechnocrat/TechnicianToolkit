<#
.SYNOPSIS
    E.X.H.U.M.E. — Enumerates, eXposes & Hunts Unmigrated Mail Entries
    Outlook PST / OST Discovery Tool for PowerShell 5.1+

.DESCRIPTION
    Scans the machine for Outlook data files and attached profiles before a
    mail migration. Inventories every .pst (and optionally .ost) on local
    drives, cross-references them against configured Outlook profiles,
    flags orphaned and stale archives, and highlights large files that
    exceed Exchange Online import limits. Produces a dark-themed HTML
    report with a migration readiness summary.

.USAGE
    PS C:\> .\exhume.ps1                                 # Interactive: scans all local drives, skips .ost
    PS C:\> .\exhume.ps1 -Unattended                    # Silent mode: exports HTML and exits
    PS C:\> .\exhume.ps1 -ScanDrives C:,D: -IncludeOst  # Custom drive list and include .ost

.NOTES
    Version : 3.6

#>

param(
    [switch]$Unattended,
    [switch]$Transcript,
    [string[]]$ScanDrives = @(),
    [switch]$IncludeOst
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

function Show-ExhumeBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  E.X.H.U.M.E. — Enumerates, eXposes & Hunts Unmigrated Mail Entries" -ForegroundColor Cyan
    Write-Host "  Outlook PST / OST Discovery Tool  v3.6" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Functions and main entry point appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — OUTLOOK PROFILE DISCOVERY
# ─────────────────────────────────────────────────────────────────────────────

# Walks every Outlook profile registered under the current user and extracts
# the paths of any PST / OST stores each profile has mounted. Works across
# Outlook 2016 / 2019 / 2021 / 365 (all use the 16.0 hive).
function Get-OutlookProfiles {
    $roots = @(
        'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles',
        'HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles',
        'HKCU:\Software\Microsoft\Office\14.0\Outlook\Profiles'
    )

    $defaultProfile = $null
    try {
        $defaultProfile = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\16.0\Outlook' -Name 'DefaultProfile' -ErrorAction SilentlyContinue).DefaultProfile
    } catch {
        # DefaultProfile key missing is normal on machines without Outlook configured.
    }

    $profiles = foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        $version = if     ($root -match '16\.0') { 'Outlook 2016+' }
                   elseif ($root -match '15\.0') { 'Outlook 2013' }
                   elseif ($root -match '14\.0') { 'Outlook 2010' }
                   else { 'Unknown' }

        foreach ($p in Get-ChildItem $root -ErrorAction SilentlyContinue) {
            $stores = [System.Collections.Generic.List[string]]::new()
            # Each profile has N subkeys for each account / store. PST path lives in
            # the '001f6700' binary value (UTF-16 encoded) on store entries.
            try {
                $all = Get-ChildItem $p.PSPath -Recurse -ErrorAction SilentlyContinue
                foreach ($sub in $all) {
                    $props = Get-ItemProperty -Path $sub.PSPath -ErrorAction SilentlyContinue
                    if (-not $props) { continue }
                    foreach ($pn in $props.PSObject.Properties.Name) {
                        if ($pn -eq '001f6700') {
                            $raw  = $props.$pn
                            if ($raw -is [byte[]]) {
                                $str = [System.Text.Encoding]::Unicode.GetString($raw) -replace "`0",''
                                if ($str -match '\.(pst|ost)$') {
                                    $stores.Add($str)
                                }
                            }
                        }
                    }
                }
            } catch {
                # Some sub-keys are unreadable under constrained token ACLs. Move on
                # and surface whatever stores we managed to collect.
            }

            [PSCustomObject]@{
                Name        = $p.PSChildName
                IsDefault   = ($p.PSChildName -eq $defaultProfile)
                Version     = $version
                Stores      = @($stores | Select-Object -Unique)
            }
        }
    }

    return @($profiles)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — DATA FILE INVENTORY (DRIVE SCAN)
# ─────────────────────────────────────────────────────────────────────────────

# Default exclusions: system dirs and OneDrive sync cache (PSTs should not live
# in OneDrive — if one is there it's a migration landmine, but the scan is
# already slow and these caches are noisy).
$script:ScanExclusions = @(
    'Windows', 'Program Files', 'Program Files (x86)', 'ProgramData',
    '$Recycle.Bin', 'System Volume Information', 'Recovery'
)

function Get-DriveRoots {
    param([string[]]$Override)

    if ($Override -and $Override.Count -gt 0) {
        return @($Override | ForEach-Object {
            $d = $_.TrimEnd(':','\','/')
            "${d}:\"
        })
    }

    # Fixed local drives only — skip CD-ROM, network, and removable by default.
    $drives = Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^[A-Z]$' -and $_.Used -ne $null }
    return @($drives | ForEach-Object { "$($_.Name):\" })
}

function Find-DataFiles {
    param(
        [string[]]$Roots,
        [bool]$IncludeOstFlag
    )

    $patterns = if ($IncludeOstFlag) { @('*.pst', '*.ost') } else { @('*.pst') }
    $found    = [System.Collections.Generic.List[object]]::new()

    foreach ($root in $Roots) {
        if (-not (Test-Path $root)) { continue }
        Write-Step "Scanning $root for Outlook data files..."

        # Enumerate top-level children so we can skip the standard exclusions without
        # paying for a full recursion into them.
        $topLevel = Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction SilentlyContinue |
                    Where-Object { $script:ScanExclusions -notcontains $_.Name }

        # Also hit loose files in the drive root (rare but possible).
        foreach ($pat in $patterns) {
            Get-ChildItem -LiteralPath $root -Filter $pat -File -Force -ErrorAction SilentlyContinue |
                ForEach-Object { $found.Add($_) }
        }

        foreach ($dir in $topLevel) {
            foreach ($pat in $patterns) {
                try {
                    Get-ChildItem -LiteralPath $dir.FullName -Filter $pat -Recurse -File -Force -ErrorAction SilentlyContinue |
                        ForEach-Object { $found.Add($_) }
                } catch {
                    # Access-denied on a sub-tree is expected for other-user profiles
                    # when not running as admin; continue with the next directory.
                }
            }
        }
    }

    $now = Get-Date
    $rows = foreach ($f in $found) {
        $daysSinceModified = [math]::Round(($now - $f.LastWriteTime).TotalDays, 0)
        $daysSinceAccessed = $null
        try {
            if ($f.LastAccessTime) { $daysSinceAccessed = [math]::Round(($now - $f.LastAccessTime).TotalDays, 0) }
        } catch {
            # LastAccessTime can be $null or raise on some FS configs; optional field.
        }

        [PSCustomObject]@{
            FullPath          = $f.FullName
            Name              = $f.Name
            Extension         = $f.Extension.ToLower()
            Bytes             = $f.Length
            Size              = Format-Bytes $f.Length
            LastModified      = $f.LastWriteTime.ToString('yyyy-MM-dd')
            DaysSinceModified = $daysSinceModified
            DaysSinceAccessed = $daysSinceAccessed
        }
    }

    return @($rows | Sort-Object Bytes -Descending)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — PROFILE/FILE CORRELATION
# ─────────────────────────────────────────────────────────────────────────────

# Annotates the disk inventory with "attached to which profile" and marks any
# file that no profile references as orphaned.
function Add-ProfileAttachment {
    param([array]$Files, [array]$Profiles)

    # Build a case-insensitive lookup of store-path -> profile-names.
    $map = @{}
    foreach ($p in $Profiles) {
        foreach ($s in $p.Stores) {
            $key = $s.ToLower()
            if (-not $map.ContainsKey($key)) { $map[$key] = [System.Collections.Generic.List[string]]::new() }
            [void]$map[$key].Add($p.Name)
        }
    }

    foreach ($f in $Files) {
        $attached = @()
        if ($map.ContainsKey($f.FullPath.ToLower())) {
            $attached = $map[$f.FullPath.ToLower()]
        }
        $f | Add-Member -NotePropertyName 'AttachedTo'  -NotePropertyValue (@($attached) -join ', ') -Force
        $f | Add-Member -NotePropertyName 'IsOrphaned'  -NotePropertyValue ($attached.Count -eq 0) -Force
    }
    return $Files
}

# ─────────────────────────────────────────────────────────────────────────────
# READINESS VERDICT
# ─────────────────────────────────────────────────────────────────────────────

function Get-Verdict {
    param([array]$Files, [array]$Profiles)

    $issues = [System.Collections.Generic.List[string]]::new()
    $warns  = [System.Collections.Generic.List[string]]::new()

    $psts       = @($Files | Where-Object { $_.Extension -eq '.pst' })
    $orphaned   = @($Files | Where-Object { $_.IsOrphaned -and $_.Extension -eq '.pst' })
    $oversize   = @($Files | Where-Object { $_.Extension -eq '.pst' -and $_.Bytes -ge 50GB })
    $large      = @($Files | Where-Object { $_.Extension -eq '.pst' -and $_.Bytes -ge 10GB -and $_.Bytes -lt 50GB })
    $stale      = @($Files | Where-Object { $_.Extension -eq '.pst' -and $_.DaysSinceAccessed -and $_.DaysSinceAccessed -ge 365 })

    if ($orphaned.Count -gt 0) {
        $warns.Add("$($orphaned.Count) PST(s) on disk are not attached to any Outlook profile — review whether each should be ingested or deleted.")
    }
    foreach ($f in $oversize) {
        $issues.Add("PST '$($f.FullPath)' is $($f.Size) — Exchange Online Import Service has a 50 GB hard limit per file; split before ingest.")
    }
    if ($large.Count -gt 0) {
        $warns.Add("$($large.Count) PST(s) are 10-50 GB — migrations will be slow; consider splitting or ingesting overnight.")
    }
    if ($stale.Count -gt 0) {
        $warns.Add("$($stale.Count) PST(s) have not been accessed in 365+ days — candidates for archive-on-ingest rather than primary-mailbox import.")
    }
    if ($psts.Count -eq 0) {
        $warns.Add('No PSTs found on disk — either already migrated or never existed on this machine.')
    }

    $verdict = if ($issues.Count -gt 0) { 'ACTION REQUIRED' }
               elseif ($warns.Count -gt 0) { 'REVIEW BEFORE MIGRATION' }
               else { 'READY TO MIGRATE' }
    $class   = if ($issues.Count -gt 0) { 'err' }
               elseif ($warns.Count -gt 0) { 'warn' }
               else { 'ok' }

    return [PSCustomObject]@{
        Verdict   = $verdict
        Class     = $class
        Issues    = @($issues)
        Warns     = @($warns)
        PstCount  = $psts.Count
        Orphans   = $orphaned.Count
        Oversize  = $oversize.Count
        Large     = $large.Count
        Stale     = $stale.Count
        TotalPstBytes = (($psts | Measure-Object -Property Bytes -Sum).Sum)
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param([array]$Files, [array]$Profiles, $Verdict, [array]$Roots, [bool]$IncludeOstFlag)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $runUser    = "$env:USERDOMAIN\$env:USERNAME"

    $tkCfg     = Get-TKConfig
    $orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    # Profile table
    $profileRows = [System.Text.StringBuilder]::new()
    if ($Profiles.Count -eq 0) {
        [void]$profileRows.Append("<tr><td colspan='4' class='tk-badge-info' style='text-align:center;'>No Outlook profiles configured for the current user.</td></tr>")
    } else {
        foreach ($p in $Profiles) {
            $defBadge = if ($p.IsDefault) { "<span class='tk-badge-ok'>Default</span>" } else { "<span class='tk-badge-info'>Secondary</span>" }
            $storeList = if ($p.Stores.Count -gt 0) {
                ($p.Stores | ForEach-Object { "<code>$(EscHtml $_)</code>" }) -join '<br>'
            } else { "<span class='tk-badge-info'>(none)</span>" }
            [void]$profileRows.Append("<tr><td>$(EscHtml $p.Name)</td><td>$defBadge</td><td>$(EscHtml $p.Version)</td><td>$storeList</td></tr>`n")
        }
    }

    # File inventory table
    $fileRows = [System.Text.StringBuilder]::new()
    if ($Files.Count -eq 0) {
        [void]$fileRows.Append("<tr><td colspan='6' class='tk-badge-ok' style='text-align:center;'>No Outlook data files found on the scanned drives.</td></tr>")
    } else {
        foreach ($f in $Files) {
            $sizeClass = if ($f.Bytes -ge 50GB) { 'tk-badge-err' }
                         elseif ($f.Bytes -ge 10GB) { 'tk-badge-warn' }
                         elseif ($f.Bytes -ge 1GB)  { 'tk-badge-info' }
                         else { 'tk-badge-ok' }
            $orphanCell = if ($f.IsOrphaned) { "<span class='tk-badge-warn'>Orphaned</span>" } else { "<span class='tk-badge-ok'>$(EscHtml $f.AttachedTo)</span>" }
            $extCell    = if ($f.Extension -eq '.ost') { "<span class='tk-badge-info'>OST</span>" } else { "<span class='tk-badge-ok'>PST</span>" }
            [void]$fileRows.Append("<tr><td><code>$(EscHtml $f.FullPath)</code></td><td>$extCell</td><td><span class='$sizeClass'>$(EscHtml $f.Size)</span></td><td>$(EscHtml $f.LastModified)</td><td>$orphanCell</td><td>$($f.DaysSinceAccessed)</td></tr>`n")
        }
    }

    # Verdict block
    $verdictBlock = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$verdictBlock.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>`n") }
    foreach ($w in $Verdict.Warns)  { [void]$verdictBlock.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>`n") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$verdictBlock.Append("<li class='tk-badge-ok'>All pre-migration checks passed.</li>")
    }

    $scopeLabel = "$($Roots -join ', ') ($(if ($IncludeOstFlag) { 'PST + OST' } else { 'PST only' }))"

    $htmlHead = Get-TKHtmlHead `
        -Title      'E.X.H.U.M.E. PST Discovery Report' `
        -ScriptName 'E.X.H.U.M.E.' `
        -Subtitle   "${orgPrefix}Outlook PST / OST Pre-Migration Discovery -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Run As'    = $runUser
            'Generated' = $reportDate
            'Scope'     = $scopeLabel
            'Verdict'   = $Verdict.Verdict
        }) `
        -NavItems   @('Verdict', 'Outlook Profiles', 'Data Files')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'E.X.H.U.M.E. v3.6'

    $totalSize = Format-Bytes $Verdict.TotalPstBytes

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Migration Readiness</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Verdict.PstCount)</div><div class="tk-summary-lbl">PST Files Found</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$totalSize</div><div class="tk-summary-lbl">Total PST Size</div></div>
    <div class="tk-summary-card $(if ($Verdict.Orphans -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Verdict.Orphans)</div><div class="tk-summary-lbl">Orphaned (not in any profile)</div></div>
    <div class="tk-summary-card $(if ($Verdict.Oversize -gt 0) { 'err' } else { 'ok' })"><div class="tk-summary-num">$($Verdict.Oversize)</div><div class="tk-summary-lbl">Over 50 GB (import blocker)</div></div>
    <div class="tk-summary-card $(if ($Verdict.Stale -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Verdict.Stale)</div><div class="tk-summary-lbl">Stale (&gt;365 days)</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Verdict &amp; Findings</span>
      <span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span>
    </div>
    <div class="tk-card"><ul class="tk-info-box" style="list-style:none;padding-left:0;">$($verdictBlock.ToString())</ul></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Outlook Profiles</span><span class="tk-section-num">$($Profiles.Count) profile(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Profile</th><th>Default</th><th>Version</th><th>Attached Stores</th></tr></thead>
        <tbody>$($profileRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Data Files (.pst$(if ($IncludeOstFlag) { ' + .ost' }))</span><span class="tk-section-num">$($Files.Count) file(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Path</th><th>Type</th><th>Size</th><th>Last Modified</th><th>Attached To</th><th>Days Since Last Access</th></tr></thead>
        <tbody>$($fileRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# CONSOLE OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

function Write-ConsoleSummary {
    param([array]$Files, [array]$Profiles, $Verdict, [array]$Roots, [bool]$IncludeOstFlag)

    Write-Section "OUTLOOK PROFILES"
    if ($Profiles.Count -eq 0) {
        Write-Warn "No Outlook profiles configured for the current user."
    } else {
        foreach ($p in $Profiles) {
            $defLabel = if ($p.IsDefault) { ' [DEFAULT]' } else { '' }
            Write-Host "  $($p.Name)$defLabel  ($($p.Version))" -ForegroundColor $C.Header
            if ($p.Stores.Count -eq 0) {
                Write-Host "    (no attached stores)" -ForegroundColor $C.Info
            } else {
                foreach ($s in $p.Stores) { Write-Host "    * $s" -ForegroundColor $C.Info }
            }
        }
    }

    Write-Section "DATA FILE INVENTORY"
    $scopeLabel = "$($Roots -join ', ') ($(if ($IncludeOstFlag) { 'PST + OST' } else { 'PST only' }))"
    Write-Info "Scanned: $scopeLabel"
    if ($Files.Count -eq 0) {
        Write-Ok "No Outlook data files found."
    } else {
        foreach ($f in $Files | Select-Object -First 30) {
            $color = if ($f.Bytes -ge 50GB) { $C.Error }
                     elseif ($f.Bytes -ge 10GB) { $C.Warning }
                     elseif ($f.IsOrphaned) { $C.Warning }
                     else { $C.Info }
            $orphanTag = if ($f.IsOrphaned) { '  [ORPHAN]' } else { "  [$($f.AttachedTo)]" }
            Write-Host ("  {0,10}  {1}{2}" -f $f.Size, $f.FullPath, $orphanTag) -ForegroundColor $color
        }
        if ($Files.Count -gt 30) {
            Write-Info "... and $($Files.Count - 30) more (see HTML report for full list)."
        }
    }

    Write-Section "MIGRATION READINESS VERDICT"
    $vColor = switch ($Verdict.Class) { 'ok' { $C.Success } 'warn' { $C.Warning } default { $C.Error } }
    Write-Host "  $($Verdict.Verdict)" -ForegroundColor $vColor
    foreach ($i in $Verdict.Issues) { Write-Host "    [!!] $i" -ForegroundColor $C.Error }
    foreach ($w in $Verdict.Warns)  { Write-Host "    [~ ] $w" -ForegroundColor $C.Warning }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        Write-Host "    [+ ] All checks passed." -ForegroundColor $C.Success
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-ExhumeBanner

$roots = Get-DriveRoots -Override $ScanDrives
if ($roots.Count -eq 0) {
    Write-Fail "No local drives detected to scan."
    exit 1
}

Write-Section "OUTLOOK DATA FILE DISCOVERY"
Write-Info "Drives: $($roots -join ', ')"
Write-Info "Include OST: $IncludeOst"
Write-Host ""

$profiles = Get-OutlookProfiles
$files    = Find-DataFiles -Roots $roots -IncludeOstFlag:$IncludeOst
$files    = Add-ProfileAttachment -Files $files -Profiles $profiles
$verdict  = Get-Verdict -Files $files -Profiles $profiles

Write-ConsoleSummary -Files $files -Profiles $profiles -Verdict $verdict -Roots $roots -IncludeOstFlag:$IncludeOst

Write-Host ""
Write-Step "Generating HTML report..."
$html      = Build-HtmlReport -Files $files -Profiles $profiles -Verdict $verdict -Roots $roots -IncludeOstFlag:$IncludeOst
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "EXHUME_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
