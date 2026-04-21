<#
.SYNOPSIS
    S.C.R.Y.E.R. -- System Consolidated Report Yielding Exhaustive Results
    Unified Diagnostic Report for PowerShell 5.1+

.DESCRIPTION
    Runs all major diagnostic queries in a single pass and produces one
    comprehensive HTML report suitable for machine handoffs or audit records.
    Covers system information, user accounts, disk space, disk health (SMART),
    and service/scheduled task status.

.USAGE
    PS C:\> .\scryer.ps1                          # Must be run as Administrator
    PS C:\> .\scryer.ps1 -Unattended              # Silent mode, no prompts
    PS C:\> .\scryer.ps1 -OutputPath "D:\Reports" # Write report to a specific directory

.NOTES
    Version : 1.0

    Tools Available
    -----------------------------------------------------------------
    G.R.I.M.O.I.R.E.       -- Technician Toolkit hub and central launcher
    A.U.S.P.E.X.           -- System diagnostics and HTML report generation
    W.A.R.D.               -- User account and local security audit
    T.H.R.E.S.H.O.L.D.     -- Disk space monitor and volume usage
    A.U.G.U.R.             -- Disk wear and health assessment, SMART status
    G.A.R.G.O.Y.L.E.       -- Service and scheduled task monitoring
    S.C.R.Y.E.R.           -- Unified diagnostic report (all of the above)

    Color Schema
    -----------------------------------------
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success messages
    Yellow   Warnings and cautions
    Red      Critical errors
    Gray     Information and details
#>

param(
    [switch]$Unattended,
    [string]$OutputPath = "",
    [switch]$Transcript
)

# ---------------------------------------------------------------------------
# INITIALIZATION
# ---------------------------------------------------------------------------

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

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $PSScriptRoot) }

# ---------------------------------------------------------------------------
# COLOR SCHEMA
# ---------------------------------------------------------------------------

$ColorSchema = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
    Accent   = 'Blue'
}

# ---------------------------------------------------------------------------
# BANNER
# ---------------------------------------------------------------------------

function Show-ScryerBanner {
    [Console]::Clear()
    Write-Host @"

  ███████╗ ██████╗██████╗ ██╗   ██╗███████╗██████╗
  ██╔════╝██╔════╝██╔══██╗╚██╗ ██╔╝██╔════╝██╔══██╗
  ███████╗██║     ██████╔╝ ╚████╔╝ █████╗  ██████╔╝
  ╚════██║██║     ██╔══██╗  ╚██╔╝  ██╔══╝  ██╔══██╗
  ███████║╚██████╗██║  ██║   ██║   ███████╗██║  ██║
  ╚══════╝ ╚═════╝╚═╝  ╚═╝   ╚═╝   ╚══════╝╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    S.C.R.Y.E.R. -- System Consolidated Report Yielding Exhaustive Results" -ForegroundColor Cyan
    Write-Host "    Unified Diagnostic Report Tool" -ForegroundColor Cyan
    Write-Host ""
}

if (-not $Unattended) { Show-ScryerBanner }

Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  UNIFIED DIAGNOSTIC REPORT" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  Machine   : $env:COMPUTERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  Run As    : $env:USERDOMAIN\$env:USERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  Started   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor $ColorSchema.Info
Write-Host ""

# ---------------------------------------------------------------------------
# SECTION 1 -- SYSTEM OVERVIEW
# ---------------------------------------------------------------------------

Write-Host "  [1/5] Collecting system information..." -ForegroundColor $ColorSchema.Progress

$sysOS   = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
$sysCS   = Get-CimInstance -ClassName Win32_ComputerSystem  -ErrorAction SilentlyContinue
$sysCPU  = Get-CimInstance -ClassName Win32_Processor       -ErrorAction SilentlyContinue | Select-Object -First 1

$osCaption    = if ($sysOS)  { $sysOS.Caption }       else { "Unknown" }
$osBuild      = if ($sysOS)  { $sysOS.BuildNumber }   else { "" }
$manufacturer = if ($sysCS)  { $sysCS.Manufacturer }  else { "" }
$model        = if ($sysCS)  { $sysCS.Model }          else { "" }
$cpuName      = if ($sysCPU) { $sysCPU.Name.Trim() }  else { "Unknown" }
$cpuCores     = if ($sysCS)  { $sysCS.NumberOfLogicalProcessors } else { "" }
$totalRAM_GB  = if ($sysCS)  { [math]::Round($sysCS.TotalPhysicalMemory / 1GB, 1) } else { 0 }
$freeRAM_GB   = if ($sysOS)  { [math]::Round($sysOS.FreePhysicalMemory / 1MB, 1) }  else { 0 }
$lastBoot     = if ($sysOS)  { $sysOS.LastBootUpTime.ToString("yyyy-MM-dd HH:mm") }  else { "" }
$uptimeSpan   = if ($sysOS)  { (Get-Date) - $sysOS.LastBootUpTime } else { $null }
$uptimeStr    = if ($uptimeSpan) { "$([int]$uptimeSpan.TotalDays)d $($uptimeSpan.Hours)h" } else { "" }
$psVersion    = $PSVersionTable.PSVersion.ToString()

Write-Host "  [+] System info collected." -ForegroundColor $ColorSchema.Success

# ---------------------------------------------------------------------------
# SECTION 2 -- USER ACCOUNTS
# ---------------------------------------------------------------------------

Write-Host "  [2/5] Auditing local user accounts..." -ForegroundColor $ColorSchema.Progress

$localUsers   = Get-LocalUser -ErrorAction SilentlyContinue
$adminMembers = @()
try { $adminMembers = (Get-LocalGroupMember -Group "Administrators" -ErrorAction SilentlyContinue).Name }
catch {}

$userRows = foreach ($u in ($localUsers | Sort-Object { -not ($adminMembers -contains "$env:COMPUTERNAME\$($u.Name)") }, { $_.LastLogon } -Descending)) {
    $isAdmin = ($adminMembers -contains "$env:COMPUTERNAME\$($u.Name)") -or ($adminMembers -contains $u.Name)
    [PSCustomObject]@{
        Name        = $u.Name
        FullName    = $u.FullName
        Enabled     = $u.Enabled
        LastLogon   = if ($u.LastLogon) { $u.LastLogon.ToString("yyyy-MM-dd") } else { "Never" }
        IsAdmin     = $isAdmin
        PwdRequired = $u.PasswordRequired
    }
}

Write-Host "  [+] User accounts audited ($($userRows.Count) users)." -ForegroundColor $ColorSchema.Success

# ---------------------------------------------------------------------------
# SECTION 3 -- DISK SPACE
# ---------------------------------------------------------------------------

Write-Host "  [3/5] Checking disk space..." -ForegroundColor $ColorSchema.Progress

$volumes = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue
$volumeRows = foreach ($v in $volumes) {
    $totalGB = [math]::Round($v.Size / 1GB, 1)
    $freeGB  = [math]::Round($v.FreeSpace / 1GB, 1)
    $usedGB  = [math]::Round(($v.Size - $v.FreeSpace) / 1GB, 1)
    $pctUsed = if ($v.Size -gt 0) { [math]::Round(($v.Size - $v.FreeSpace) / $v.Size * 100, 0) } else { 0 }
    $health  = if ($pctUsed -gt 95) { "err" } elseif ($pctUsed -gt 85) { "warn" } else { "ok" }
    [PSCustomObject]@{
        Letter  = $v.DeviceID
        Label   = $v.VolumeName
        TotalGB = $totalGB
        UsedGB  = $usedGB
        FreeGB  = $freeGB
        PctUsed = $pctUsed
        Health  = $health
    }
}

$volWarnCount = ($volumeRows | Where-Object { $_.Health -ne "ok" }).Count
Write-Host "  [+] Disk space checked ($($volumeRows.Count) volumes, $volWarnCount warnings)." -ForegroundColor $ColorSchema.Success

# ---------------------------------------------------------------------------
# SECTION 4 -- DISK HEALTH
# ---------------------------------------------------------------------------

Write-Host "  [4/5] Assessing disk health..." -ForegroundColor $ColorSchema.Progress

$diskRows      = @()
$smartAvailable = $false
try {
    $physDisks      = Get-PhysicalDisk -ErrorAction Stop
    $smartAvailable = $true
    $diskRows = foreach ($d in $physDisks) {
        $sizeGB = [math]::Round($d.Size / 1GB, 0)
        $wear   = $null
        $temp   = $null
        try {
            $rel = Get-StorageReliabilityCounter -PhysicalDisk $d -ErrorAction SilentlyContinue
            if ($rel) {
                $wear = $rel.Wear
                $temp = $rel.Temperature
            }
        } catch {}
        $healthClass = switch -Regex ($d.HealthStatus) {
            'Healthy' { "ok"   }
            'Warning' { "warn" }
            default   { "err"  }
        }
        [PSCustomObject]@{
            Name        = $d.FriendlyName
            MediaType   = $d.MediaType
            SizeGB      = $sizeGB
            Health      = $d.HealthStatus
            HealthClass = $healthClass
            Temp        = $temp
            Wear        = $wear
        }
    }
} catch {
    $smartAvailable = $false
}

$diskWarnCount = ($diskRows | Where-Object { $_.HealthClass -ne "ok" }).Count
Write-Host "  [+] Disk health assessed ($($diskRows.Count) disks)." -ForegroundColor $ColorSchema.Success

# ---------------------------------------------------------------------------
# SECTION 5 -- SERVICES & SCHEDULED TASKS
# ---------------------------------------------------------------------------

Write-Host "  [5/5] Checking services and scheduled tasks..." -ForegroundColor $ColorSchema.Progress

$triggerExclusions = @('gupdate','gupdatem','edgeupdate','edgeupdatem','MapsBroker',
    'RemoteRegistry','SharedAccess','TabletInputService','WbioSrvc','lfsvc',
    'SCardSvr','SensrSvc','WSearch','wuauserv','BITS','DoSvc','UsoSvc','WerSvc',
    'AppReadiness','tiledatamodelsvc','CDPSvc','OneSyncSvc','PimIndexMaintenanceSvc',
    'MessagingService','cbdhsvc','DevicesFlowUserSvc')

$stoppedSvcs = Get-Service -ErrorAction SilentlyContinue |
    Where-Object { $_.StartType -eq 'Automatic' -and $_.Status -eq 'Stopped' -and
                   $_.Name -notin $triggerExclusions } |
    Select-Object Name, DisplayName, Status, StartType |
    Sort-Object DisplayName

$failedTasks = @()
try {
    $failedTasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
        Where-Object { $_.State -ne 'Disabled' -and $_.TaskPath -notlike '\Microsoft\*' } |
        ForEach-Object {
            $info = $_ | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue
            if ($info -and $info.LastTaskResult -ne 0 -and $info.LastRunTime -gt [datetime]::MinValue) {
                [PSCustomObject]@{
                    TaskName    = $_.TaskName
                    TaskPath    = $_.TaskPath
                    LastRunTime = $info.LastRunTime.ToString("yyyy-MM-dd HH:mm")
                    LastResult  = "0x{0:X8}" -f $info.LastTaskResult
                }
            }
        } | Where-Object { $_ -ne $null } | Select-Object -First 20
} catch {}

Write-Host "  [+] Services/tasks checked ($($stoppedSvcs.Count) svc issues, $($failedTasks.Count) task failures)." -ForegroundColor $ColorSchema.Success

# ---------------------------------------------------------------------------
# HTML REPORT BUILD
# ---------------------------------------------------------------------------

$_cfg = Get-TKConfig
$reportDir = if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath }
             elseif (-not [string]::IsNullOrWhiteSpace($_cfg.LogDirectory)) { $_cfg.LogDirectory }
             else { $PSScriptRoot }
if (-not (Test-Path $reportDir)) { New-Item -ItemType Directory -Path $reportDir -Force | Out-Null }
$reportFile = Join-Path $reportDir "SCRYER_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

Write-Host ""
Write-Host "  [*] Building HTML report..." -ForegroundColor $ColorSchema.Progress

$html = Get-TKHtmlHead -Title "SCRYER Unified Report" `
    -ScriptName "S.C.R.Y.E.R." `
    -Subtitle $env:COMPUTERNAME `
    -MetaItems ([ordered]@{
        'Generated' = (Get-Date -Format 'yyyy-MM-dd HH:mm')
        'OS'        = "$osCaption (Build $osBuild)"
        'Run As'    = "$env:USERDOMAIN\$env:USERNAME"
    }) `
    -NavItems @('System Overview','User Accounts','Disk Space','Disk Health','Services and Tasks')

# ------------------------------------------------------------------
# Section 01 -- System Overview
# ------------------------------------------------------------------

$html += @"

<div class="tk-section" id="system-overview">
  <div class="tk-section-title"><span class="tk-section-num">01</span> System Overview</div>
  <div class="tk-card">
    <table class="tk-table">
      <tbody>
        <tr><td class="tk-card-label">Hostname</td><td>$(EscHtml $env:COMPUTERNAME)</td></tr>
        <tr><td class="tk-card-label">Operating System</td><td>$(EscHtml $osCaption) (Build $osBuild)</td></tr>
        <tr><td class="tk-card-label">Manufacturer / Model</td><td>$(EscHtml $manufacturer) $(EscHtml $model)</td></tr>
        <tr><td class="tk-card-label">CPU</td><td>$(EscHtml $cpuName) ($cpuCores logical cores)</td></tr>
        <tr><td class="tk-card-label">RAM</td><td>$totalRAM_GB GB total / $freeRAM_GB GB free</td></tr>
        <tr><td class="tk-card-label">Last Boot</td><td>$(EscHtml $lastBoot)</td></tr>
        <tr><td class="tk-card-label">Uptime</td><td>$(EscHtml $uptimeStr)</td></tr>
        <tr><td class="tk-card-label">PowerShell Version</td><td>$(EscHtml $psVersion)</td></tr>
      </tbody>
    </table>
  </div>
</div>

"@

# ------------------------------------------------------------------
# Section 02 -- User Accounts
# ------------------------------------------------------------------

$userTableRows = ""
foreach ($u in $userRows) {
    $statusBadge = if ($u.Enabled) {
        "<span class='tk-badge-ok'>Enabled</span>"
    } else {
        "<span class='tk-badge-err'>Disabled</span>"
    }
    $adminCell = if ($u.IsAdmin) { "<span class='tk-badge-err'>Yes</span>" } else { "" }
    $pwdCell   = if ($u.PwdRequired) { "Yes" } else { "No" }
    $userTableRows += @"
        <tr>
          <td>$(EscHtml $u.Name)</td>
          <td>$(EscHtml $u.FullName)</td>
          <td>$statusBadge</td>
          <td>$(EscHtml $u.LastLogon)</td>
          <td>$adminCell</td>
          <td>$pwdCell</td>
        </tr>
"@
}

$html += @"

<div class="tk-section" id="user-accounts">
  <div class="tk-section-title"><span class="tk-section-num">02</span> User Accounts</div>
  <div class="tk-card">
    <table class="tk-table">
      <thead>
        <tr>
          <th>User</th>
          <th>Full Name</th>
          <th>Status</th>
          <th>Last Logon</th>
          <th>Admin</th>
          <th>Pwd Required</th>
        </tr>
      </thead>
      <tbody>
$userTableRows
      </tbody>
    </table>
  </div>
</div>

"@

# ------------------------------------------------------------------
# Section 03 -- Disk Space
# ------------------------------------------------------------------

$diskSpaceCards = ""
foreach ($vol in $volumeRows) {
    $letterLabel = if ($vol.Label) { "$(EscHtml $vol.Letter) -- $(EscHtml $vol.Label)" } else { "$(EscHtml $vol.Letter)" }
    $diskSpaceCards += @"
  <div class="tk-card">
    <div class="tk-card-header">$letterLabel</div>
    <div class="tk-progress-wrap">
      <div class="tk-progress-bar $($vol.Health)" style="width: $($vol.PctUsed)%"></div>
    </div>
    <div style="font-size:0.85em; color:var(--tk-text-dim); margin-top:4px">
      $($vol.PctUsed)% used &nbsp;|&nbsp; $($vol.UsedGB) GB used / $($vol.TotalGB) GB total &nbsp;|&nbsp; $($vol.FreeGB) GB free
    </div>
  </div>
"@
}

if (-not $diskSpaceCards) {
    $diskSpaceCards = "<div class='tk-card'><div class='tk-info-box'>No fixed volumes found.</div></div>"
}

$html += @"

<div class="tk-section" id="disk-space">
  <div class="tk-section-title"><span class="tk-section-num">03</span> Disk Space</div>
$diskSpaceCards
</div>

"@

# ------------------------------------------------------------------
# Section 04 -- Disk Health
# ------------------------------------------------------------------

if (-not $smartAvailable) {
    $diskHealthBody = @"
  <div class="tk-card">
    <div class="tk-info-box">SMART data unavailable -- Storage module not accessible on this system.</div>
  </div>
"@
} else {
    $diskHealthRows = ""
    foreach ($d in $diskRows) {
        $healthBadge = "<span class='tk-badge-$($d.HealthClass)'>$(EscHtml $d.Health)</span>"
        $tempVal     = if ($null -ne $d.Temp) { $d.Temp } else { "--" }
        $wearVal     = if ($null -ne $d.Wear) { "$($d.Wear)%" } else { "--" }
        $diskHealthRows += @"
        <tr>
          <td>$(EscHtml $d.Name)</td>
          <td>$(EscHtml $d.MediaType)</td>
          <td>$($d.SizeGB)</td>
          <td>$healthBadge</td>
          <td>$tempVal</td>
          <td>$wearVal</td>
        </tr>
"@
    }
    $diskHealthBody = @"
  <div class="tk-card">
    <table class="tk-table">
      <thead>
        <tr>
          <th>Drive</th>
          <th>Type</th>
          <th>Size (GB)</th>
          <th>Health</th>
          <th>Temp (C)</th>
          <th>Wear (%)</th>
        </tr>
      </thead>
      <tbody>
$diskHealthRows
      </tbody>
    </table>
  </div>
"@
}

$html += @"

<div class="tk-section" id="disk-health">
  <div class="tk-section-title"><span class="tk-section-num">04</span> Disk Health</div>
$diskHealthBody
</div>

"@

# ------------------------------------------------------------------
# Section 05 -- Services and Tasks
# ------------------------------------------------------------------

# Stopped automatic services card
if (-not $stoppedSvcs -or $stoppedSvcs.Count -eq 0) {
    $svcCardBody = "    <div class='tk-info-box'><span class='tk-badge-ok'>OK</span> No stopped automatic services found.</div>"
} else {
    $svcRows = ""
    foreach ($s in $stoppedSvcs) {
        $svcRows += @"
        <tr>
          <td>$(EscHtml $s.Name)</td>
          <td>$(EscHtml $s.DisplayName)</td>
          <td><span class='tk-badge-warn'>Stopped</span></td>
          <td>$(EscHtml $s.StartType)</td>
        </tr>
"@
    }
    $svcCardBody = @"
    <table class="tk-table">
      <thead>
        <tr>
          <th>Name</th>
          <th>Display Name</th>
          <th>Status</th>
          <th>Start Type</th>
        </tr>
      </thead>
      <tbody>
$svcRows
      </tbody>
    </table>
"@
}

# Failed scheduled tasks card
if (-not $failedTasks -or $failedTasks.Count -eq 0) {
    $taskCardBody = "    <div class='tk-info-box'><span class='tk-badge-ok'>OK</span> No failed scheduled tasks found.</div>"
} else {
    $taskRows = ""
    foreach ($t in $failedTasks) {
        $taskRows += @"
        <tr>
          <td>$(EscHtml $t.TaskName)</td>
          <td>$(EscHtml $t.TaskPath)</td>
          <td>$(EscHtml $t.LastRunTime)</td>
          <td><span class='tk-badge-warn'>$(EscHtml $t.LastResult)</span></td>
        </tr>
"@
    }
    $taskCardBody = @"
    <table class="tk-table">
      <thead>
        <tr>
          <th>Task Name</th>
          <th>Path</th>
          <th>Last Run</th>
          <th>Last Result</th>
        </tr>
      </thead>
      <tbody>
$taskRows
      </tbody>
    </table>
"@
}

$html += @"

<div class="tk-section" id="services-and-tasks">
  <div class="tk-section-title"><span class="tk-section-num">05</span> Services and Tasks</div>
  <div class="tk-card">
    <div class="tk-card-header">Stopped Automatic Services</div>
$svcCardBody
  </div>
  <div class="tk-card">
    <div class="tk-card-header">Failed Scheduled Tasks</div>
$taskCardBody
  </div>
</div>

"@

# ------------------------------------------------------------------
# Footer
# ------------------------------------------------------------------

$html += Get-TKHtmlFoot -ScriptName "S.C.R.Y.E.R. v1.0"

# ------------------------------------------------------------------
# Write report and open
# ------------------------------------------------------------------

$html | Out-File -FilePath $reportFile -Encoding UTF8
Write-Host "  [+] Report saved: $reportFile" -ForegroundColor $ColorSchema.Success

if (-not $Unattended) {
    try { Start-Process $reportFile } catch {}
}

# ---------------------------------------------------------------------------
# CONSOLE SUMMARY
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host "  SCRYER COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host "  Report    : $reportFile" -ForegroundColor $ColorSchema.Info
Write-Host "  Users     : $($userRows.Count)" -ForegroundColor $ColorSchema.Info
Write-Host "  Volumes   : $($volumeRows.Count) ($volWarnCount with warnings)" -ForegroundColor $(if ($volWarnCount -gt 0) { $ColorSchema.Warning } else { $ColorSchema.Info })
Write-Host "  Disks     : $($diskRows.Count) ($diskWarnCount with issues)" -ForegroundColor $(if ($diskWarnCount -gt 0) { $ColorSchema.Warning } else { $ColorSchema.Info })
Write-Host "  Svc Issues: $($stoppedSvcs.Count)" -ForegroundColor $(if ($stoppedSvcs.Count -gt 0) { $ColorSchema.Warning } else { $ColorSchema.Info })
Write-Host "  Task Fails: $($failedTasks.Count)" -ForegroundColor $(if ($failedTasks.Count -gt 0) { $ColorSchema.Warning } else { $ColorSchema.Info })
Write-Host ""
Write-Host ("  " + ("=" * 62)) -ForegroundColor $ColorSchema.Header
Write-Host ""

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
