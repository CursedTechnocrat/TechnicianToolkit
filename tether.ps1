<#
.SYNOPSIS
    T.E.T.H.E.R. — Tests Endpoint Tethering: Hosted Environment Readiness
    OneDrive Known-Folder-Move Pre-Migration Validator for PowerShell 5.1+

.DESCRIPTION
    Local machine audit that answers the pre-migration question: "Is this
    user's data actually going to be in the cloud when we hand them a new
    laptop?" Checks OneDrive client state, signed-in accounts, Known
    Folder Move engagement for Desktop / Documents / Pictures, content
    volumes, and recent sync errors. Produces a dark-themed HTML report
    with a red / yellow / green readiness verdict.

.USAGE
    PS C:\> .\tether.ps1                    # Interactive run against current user
    PS C:\> .\tether.ps1 -Unattended        # Silent mode, export HTML and exit

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

function Show-TetherBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  T.E.T.H.E.R. — Tests Endpoint Tethering: Hosted Environment Readiness" -ForegroundColor Cyan
    Write-Host "  OneDrive Known-Folder-Move Pre-Migration Validator  v3.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Data collection and report-building functions are appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — ONEDRIVE CLIENT STATE
# ─────────────────────────────────────────────────────────────────────────────

function Get-OneDriveClient {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe'),
        'C:\Program Files\Microsoft OneDrive\OneDrive.exe',
        'C:\Program Files (x86)\Microsoft OneDrive\OneDrive.exe'
    )

    $exe = $null
    foreach ($p in $candidates) {
        if (Test-Path $p) { $exe = $p; break }
    }

    $version = $null
    if ($exe) {
        try {
            $version = (Get-Item $exe).VersionInfo.ProductVersion
        } catch {
            $version = 'unknown'
        }
    }

    $running = $false
    try {
        $running = [bool](Get-Process -Name 'OneDrive' -ErrorAction SilentlyContinue)
    } catch {
        # Process enumeration may fail under constrained contexts — keep default $false.
    }

    return [PSCustomObject]@{
        Installed = [bool]$exe
        Path      = $exe
        Version   = $version
        Running   = $running
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — SIGNED-IN ACCOUNTS
# ─────────────────────────────────────────────────────────────────────────────

function Get-OneDriveAccounts {
    $root = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
    if (-not (Test-Path $root)) { return @() }

    $accounts = foreach ($child in Get-ChildItem $root -ErrorAction SilentlyContinue) {
        $key = $child.PSPath
        $props = $null
        try {
            $props = Get-ItemProperty -Path $key -ErrorAction Stop
        } catch {
            continue
        }

        $accountType = if ($child.PSChildName -like 'Business*') { 'Business' }
                       elseif ($child.PSChildName -eq 'Personal') { 'Personal' }
                       else { 'Unknown' }

        [PSCustomObject]@{
            AccountKey  = $child.PSChildName
            AccountType = $accountType
            UserEmail   = $props.UserEmail
            UserFolder  = $props.UserFolder
            TenantId    = $props.ConfiguredTenantId
            DisplayName = $props.DisplayName
        }
    }

    return @($accounts)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — KNOWN FOLDER MOVE STATUS
# ─────────────────────────────────────────────────────────────────────────────

# Shell folder GUIDs (User Shell Folders registry uses named values, but the
# KFM check is the simpler "does the resolved path contain 'OneDrive'" rule).
$script:KfmFolders = @(
    @{ Label = 'Desktop';   ShellName = 'Desktop' },
    @{ Label = 'Documents'; ShellName = 'Personal' },
    @{ Label = 'Pictures';  ShellName = 'My Pictures' }
)

function Get-KfmStatus {
    $shellKey = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    $results = foreach ($folder in $script:KfmFolders) {
        $path = $null
        try {
            $raw = (Get-ItemProperty -Path $shellKey -Name $folder.ShellName -ErrorAction Stop).($folder.ShellName)
            $path = [Environment]::ExpandEnvironmentVariables($raw)
        } catch {
            $path = $null
        }

        $inOneDrive = $false
        if ($path) {
            $inOneDrive = ($path -match '(?i)\\OneDrive') -or
                          ($path -like "$env:OneDrive*") -or
                          ($path -like "$env:OneDriveCommercial*")
        }

        [PSCustomObject]@{
            Label      = $folder.Label
            Path       = if ($path) { $path } else { '(not set)' }
            Redirected = $inOneDrive
        }
    }

    return @($results)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — FOLDER CONTENT VOLUME
# ─────────────────────────────────────────────────────────────────────────────

function Get-FolderVolume {
    param([array]$KfmStatus)

    $results = foreach ($f in $KfmStatus) {
        $count = 0
        $bytes = 0
        $errMsg = $null

        if ($f.Path -and (Test-Path -LiteralPath $f.Path)) {
            try {
                $items = Get-ChildItem -LiteralPath $f.Path -Recurse -File -Force -ErrorAction SilentlyContinue
                $count = @($items).Count
                $bytes = ($items | Measure-Object -Property Length -Sum).Sum
                if (-not $bytes) { $bytes = 0 }
            } catch {
                $errMsg = $_.Exception.Message
            }
        } else {
            $errMsg = 'Path does not exist'
        }

        [PSCustomObject]@{
            Label     = $f.Label
            Path      = $f.Path
            FileCount = $count
            Bytes     = $bytes
            Size      = Format-Bytes $bytes
            Error     = $errMsg
        }
    }

    return @($results)
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 — RECENT SYNC ERRORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-SyncErrors {
    $since = (Get-Date).AddDays(-7)
    $events = @()

    # OneDrive errors historically land in the Application log with provider
    # 'Microsoft-Windows-User Profiles Service' or the generic 'OneDrive' source.
    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName   = 'Application'
            ProviderName = @('OneDrive', 'Microsoft-Windows-User Profiles Service')
            Level     = 1, 2, 3  # Critical, Error, Warning
            StartTime = $since
        } -ErrorAction SilentlyContinue
    } catch {
        # Filtering on non-existent providers throws — fall back to a broader sweep.
        try {
            $events = Get-WinEvent -LogName Application -MaxEvents 500 -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.TimeCreated -ge $since -and
                    $_.LevelDisplayName -in @('Error','Warning','Critical') -and
                    ($_.ProviderName -like '*OneDrive*' -or $_.Message -match 'OneDrive')
                }
        } catch {
            $events = @()
        }
    }

    $rows = foreach ($e in $events) {
        [PSCustomObject]@{
            Time     = $e.TimeCreated.ToString('yyyy-MM-dd HH:mm')
            Level    = $e.LevelDisplayName
            Provider = $e.ProviderName
            EventId  = $e.Id
            Message  = ($e.Message -split "`n" | Select-Object -First 1).Trim()
        }
    }

    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# READINESS VERDICT
# ─────────────────────────────────────────────────────────────────────────────

function Get-ReadinessVerdict {
    param($Client, [array]$Accounts, [array]$Kfm, [array]$Volume, [array]$Errors)

    $issues  = [System.Collections.Generic.List[string]]::new()
    $warns   = [System.Collections.Generic.List[string]]::new()

    if (-not $Client.Installed) { $issues.Add('OneDrive client is not installed.') }
    elseif (-not $Client.Running) { $warns.Add('OneDrive client is installed but not currently running.') }

    $businessAccounts = @($Accounts | Where-Object { $_.AccountType -eq 'Business' -and $_.UserEmail })
    if ($businessAccounts.Count -eq 0) {
        $issues.Add('No Business/Work OneDrive account is signed in.')
    }

    $notRedirected = @($Kfm | Where-Object { -not $_.Redirected })
    foreach ($r in $notRedirected) {
        $issues.Add("Known Folder '$($r.Label)' is not redirected to OneDrive.")
    }

    foreach ($v in $Volume) {
        if ($v.Bytes -gt 25GB) {
            $warns.Add("Folder '$($v.Label)' is $($v.Size) — large uploads may take time to complete.")
        }
    }

    $errCount = @($Errors).Count
    if ($errCount -gt 0) {
        $warns.Add("$errCount OneDrive Application-log errors/warnings in the last 7 days.")
    }

    $verdict = if ($issues.Count -gt 0) { 'NOT READY' }
               elseif ($warns.Count -gt 0) { 'READY WITH WARNINGS' }
               else { 'READY' }
    $class = if ($issues.Count -gt 0) { 'err' }
             elseif ($warns.Count -gt 0) { 'warn' }
             else { 'ok' }

    return [PSCustomObject]@{
        Verdict = $verdict
        Class   = $class
        Issues  = @($issues)
        Warns   = @($warns)
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param($Client, [array]$Accounts, [array]$Kfm, [array]$Volume, [array]$Errors, $Verdict)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $runUser    = "$env:USERDOMAIN\$env:USERNAME"

    $tkCfg     = Get-TKConfig
    $orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $clientRow = if ($Client.Installed) {
        $runBadge = if ($Client.Running) { "<span class='tk-badge-ok'>Running</span>" } else { "<span class='tk-badge-warn'>Not running</span>" }
        "<tr><td><span class='tk-badge-ok'>Installed</span></td><td>$(EscHtml $Client.Version)</td><td>$runBadge</td><td><code>$(EscHtml $Client.Path)</code></td></tr>"
    } else {
        "<tr><td colspan='4' class='tk-badge-err' style='text-align:center;'>OneDrive client is NOT installed on this machine.</td></tr>"
    }

    $acctRows = [System.Text.StringBuilder]::new()
    if ($Accounts.Count -eq 0) {
        [void]$acctRows.Append("<tr><td colspan='4' class='tk-badge-err' style='text-align:center;'>No OneDrive accounts signed in.</td></tr>")
    } else {
        foreach ($a in $Accounts) {
            $typeBadge = if ($a.AccountType -eq 'Business') { "<span class='tk-badge-ok'>Business</span>" }
                         elseif ($a.AccountType -eq 'Personal') { "<span class='tk-badge-warn'>Personal</span>" }
                         else { "<span class='tk-badge-info'>Unknown</span>" }
            [void]$acctRows.Append("<tr><td>$typeBadge</td><td>$(EscHtml $a.UserEmail)</td><td><code>$(EscHtml $a.UserFolder)</code></td><td>$(EscHtml $a.TenantId)</td></tr>`n")
        }
    }

    $kfmRows = [System.Text.StringBuilder]::new()
    foreach ($k in $Kfm) {
        $badge = if ($k.Redirected) { "<span class='tk-badge-ok'>Redirected</span>" } else { "<span class='tk-badge-err'>Local only</span>" }
        [void]$kfmRows.Append("<tr><td>$(EscHtml $k.Label)</td><td>$badge</td><td><code>$(EscHtml $k.Path)</code></td></tr>`n")
    }

    $volRows = [System.Text.StringBuilder]::new()
    foreach ($v in $Volume) {
        $sizeClass = if ($v.Bytes -ge 25GB) { 'tk-badge-err' }
                     elseif ($v.Bytes -ge 5GB) { 'tk-badge-warn' }
                     else { 'tk-badge-ok' }
        $errCell = if ($v.Error) { "<span class='tk-badge-warn'>$(EscHtml $v.Error)</span>" } else { '' }
        [void]$volRows.Append("<tr><td>$(EscHtml $v.Label)</td><td>$($v.FileCount)</td><td><span class='$sizeClass'>$(EscHtml $v.Size)</span></td><td>$errCell</td></tr>`n")
    }

    $errRows = [System.Text.StringBuilder]::new()
    if ($Errors.Count -eq 0) {
        [void]$errRows.Append("<tr><td colspan='4' class='tk-badge-ok' style='text-align:center;'>No OneDrive errors in the last 7 days.</td></tr>")
    } else {
        foreach ($e in $Errors | Select-Object -First 50) {
            $levelClass = if ($e.Level -eq 'Error' -or $e.Level -eq 'Critical') { 'tk-badge-err' } else { 'tk-badge-warn' }
            [void]$errRows.Append("<tr><td>$(EscHtml $e.Time)</td><td><span class='$levelClass'>$(EscHtml $e.Level)</span></td><td>$($e.EventId)</td><td>$(EscHtml $e.Message)</td></tr>`n")
        }
    }

    $verdictBlock = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$verdictBlock.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>`n") }
    foreach ($w in $Verdict.Warns)  { [void]$verdictBlock.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>`n") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$verdictBlock.Append("<li class='tk-badge-ok'>All pre-migration checks passed.</li>")
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'T.E.T.H.E.R. OneDrive Pre-Migration Report' `
        -ScriptName 'T.E.T.H.E.R.' `
        -Subtitle   "${orgPrefix}OneDrive KFM Readiness -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Run As'    = $runUser
            'Generated' = $reportDate
            'Verdict'   = $Verdict.Verdict
        }) `
        -NavItems   @('Verdict', 'Client', 'Accounts', 'Known Folders', 'Content Volume', 'Sync Errors')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'T.E.T.H.E.R. v3.0'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Migration Readiness</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Accounts.Count)</div><div class="tk-summary-lbl">OneDrive Accounts</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(@($Kfm | Where-Object { $_.Redirected }).Count) / $($Kfm.Count)</div><div class="tk-summary-lbl">Folders Redirected</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(Format-Bytes (($Volume | Measure-Object -Property Bytes -Sum).Sum))</div><div class="tk-summary-lbl">Total Content</div></div>
    <div class="tk-summary-card $(if ($Errors.Count -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($Errors.Count)</div><div class="tk-summary-lbl">Sync Errors (7d)</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header">
      <span class="tk-section-title">Verdict &amp; Findings</span>
      <span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span>
    </div>
    <div class="tk-card"><ul class="tk-info-box" style="list-style:none;padding-left:0;">$($verdictBlock.ToString())</ul></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">OneDrive Client</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Installed</th><th>Version</th><th>Process</th><th>Path</th></tr></thead>
        <tbody>$clientRow</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Signed-In Accounts</span><span class="tk-section-num">$($Accounts.Count) account(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Type</th><th>Email</th><th>Sync Root</th><th>Tenant ID</th></tr></thead>
        <tbody>$($acctRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Known Folder Redirection</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Folder</th><th>Status</th><th>Resolved Path</th></tr></thead>
        <tbody>$($kfmRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Content Volume</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Folder</th><th>File Count</th><th>Size</th><th>Error</th></tr></thead>
        <tbody>$($volRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Sync Errors (last 7 days)</span><span class="tk-section-num">$($Errors.Count) event(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Time</th><th>Level</th><th>Event ID</th><th>Message</th></tr></thead>
        <tbody>$($errRows.ToString())</tbody>
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
    param($Client, [array]$Accounts, [array]$Kfm, [array]$Volume, [array]$Errors, $Verdict)

    Write-Section "ONEDRIVE CLIENT"
    if ($Client.Installed) {
        Write-Ok "Installed: $($Client.Path) (v$($Client.Version))"
        if ($Client.Running) { Write-Ok "Process is running." } else { Write-Warn "Process is not currently running." }
    } else {
        Write-Fail "OneDrive client is NOT installed."
    }

    Write-Section "SIGNED-IN ACCOUNTS"
    if ($Accounts.Count -eq 0) {
        Write-Fail "No OneDrive accounts are signed in."
    } else {
        foreach ($a in $Accounts) {
            $label = "[$($a.AccountType)] $($a.UserEmail)"
            Write-Host "  $label" -ForegroundColor $(if ($a.AccountType -eq 'Business') { $C.Success } else { $C.Warning })
            Write-Host "    Sync root : $($a.UserFolder)" -ForegroundColor $C.Info
        }
    }

    Write-Section "KNOWN FOLDER REDIRECTION"
    foreach ($k in $Kfm) {
        if ($k.Redirected) {
            Write-Ok "$($k.Label) -> redirected to OneDrive"
        } else {
            Write-Fail "$($k.Label) -> local only ($($k.Path))"
        }
    }

    Write-Section "CONTENT VOLUME"
    foreach ($v in $Volume) {
        $color = if ($v.Bytes -ge 25GB) { $C.Error } elseif ($v.Bytes -ge 5GB) { $C.Warning } else { $C.Success }
        Write-Host ("  {0,-12} {1,8} file(s)  {2}" -f $v.Label, $v.FileCount, $v.Size) -ForegroundColor $color
    }

    Write-Section "SYNC ERRORS (LAST 7 DAYS)"
    if ($Errors.Count -eq 0) {
        Write-Ok "No OneDrive errors in the last 7 days."
    } else {
        Write-Warn "$($Errors.Count) event(s) in the last 7 days."
        foreach ($e in $Errors | Select-Object -First 10) {
            $color = if ($e.Level -in 'Error','Critical') { $C.Error } else { $C.Warning }
            Write-Host ("  [{0}] {1} (id={2}): {3}" -f $e.Time, $e.Level, $e.EventId, $e.Message) -ForegroundColor $color
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

Show-TetherBanner

$client   = Get-OneDriveClient
$accounts = Get-OneDriveAccounts
$kfm      = Get-KfmStatus
$volume   = Get-FolderVolume -KfmStatus $kfm
$errors   = Get-SyncErrors
$verdict  = Get-ReadinessVerdict -Client $client -Accounts $accounts -Kfm $kfm -Volume $volume -Errors $errors

Write-ConsoleSummary -Client $client -Accounts $accounts -Kfm $kfm -Volume $volume -Errors $errors -Verdict $verdict

Write-Host ""
Write-Step "Generating HTML report..."
$html = Build-HtmlReport -Client $client -Accounts $accounts -Kfm $kfm -Volume $volume -Errors $errors -Verdict $verdict
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "TETHER_${timestamp}.html"

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
