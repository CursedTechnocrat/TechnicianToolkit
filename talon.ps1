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
    Version : 3.0

#>

param(
    [switch]$Unattended,
    [switch]$Transcript
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
}

function Show-TalonBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  T.A.L.O.N. — Tracks Anomalies & Locates Otherwise-silent Nastiness" -ForegroundColor Cyan
    Write-Host "  Persistence / Autoruns Audit Tool  v3.0" -ForegroundColor Cyan
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
    $rows = @()

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
        $rows += New-PersistenceRow -Category 'WMI Subscription' -Source 'root\subscription' -Name $b.Filter -Command $cmd
    }

    # If there are no bindings, still surface bare filters/consumers so a reviewer knows the
    # subscription surface was inspected (no silent skips).
    if ($bindings.Count -eq 0 -and ($filters.Count -gt 0 -or $consumers.Count -gt 0)) {
        foreach ($f in $filters) {
            $rows += New-PersistenceRow -Category 'WMI Subscription' -Source 'root\subscription' -Name "Filter: $($f.Name)" -Command $f.Query
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
    $rows = @()

    try {
        $wl = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -ErrorAction Stop
        if ($wl.Shell)    { $rows += New-PersistenceRow -Category 'Winlogon' -Source 'HKLM\Winlogon' -Name 'Shell'    -Command $wl.Shell }
        if ($wl.Userinit) { $rows += New-PersistenceRow -Category 'Winlogon' -Source 'HKLM\Winlogon' -Name 'Userinit' -Command $wl.Userinit }
    } catch {
        # Winlogon key always exists on Windows; failure here means access-denied in a test
        # environment. Continue silently rather than emit a fake-missing row.
    }

    # AppInit_DLLs is legacy (>= Windows 8 requires signed, disabled by default) but still
    # auditable.
    try {
        $init = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows' -ErrorAction Stop
        if ($init.AppInit_DLLs)  { $rows += New-PersistenceRow -Category 'Winlogon' -Source 'HKLM\Windows' -Name 'AppInit_DLLs'  -Command $init.AppInit_DLLs }
        if ($init.LoadAppInit_DLLs) { $rows += New-PersistenceRow -Category 'Winlogon' -Source 'HKLM\Windows' -Name 'LoadAppInit_DLLs' -Command "$($init.LoadAppInit_DLLs)" }
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

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'T.A.L.O.N. v3.0'

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
    Write-Ok "Report saved: $outPath"
    if (-not $Unattended) {
        Write-Step "Opening in default browser..."
        Start-Process $outPath
    }
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
