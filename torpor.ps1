# torpor.ps1 - T.O.R.P.O.R. — Traces Overload, Resource Pressure & Offending Routines
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
    T.O.R.P.O.R. — Traces Overload, Resource Pressure & Offending Routines
    Slow-Machine Triage Tool for PowerShell 5.1+

.DESCRIPTION
    Answers "why is this PC slow?" by watching the machine for a short sample
    window and naming what is using it:

      - CPU load, processor queue length, and whether the CPU is being held
        below full speed by power or thermal limits
      - Memory: available RAM, commit charge against the commit limit, the
        page file, and how much RAM is installed in the first place
      - Physical disks: busy time, response time, IOPS and throughput, and
        whether the system disk is a spinning drive
      - The processes responsible, grouped by name (thirty browser processes
        are one browser), with a hint for well-known offenders and the
        services living inside any busy svchost
      - Context that explains a slow machine without a busy one: uptime, the
        power plan, free space on the system drive, startup programs, boot
        times and the apps / drivers / services Windows blamed for slow boots,
        and applications that keep hanging

    Counters are read through WMI performance classes rather than Get-Counter,
    whose counter paths are translated on non-English Windows.

    Read-only -- nothing on the machine is changed. T.O.R.P.O.R. lists startup
    programs for their cost, not their safety: T.A.L.O.N. is the persistence
    audit.

.USAGE
    PS C:\> .\torpor.ps1                               # Interactive menu
    PS C:\> .\torpor.ps1 -Unattended                   # 30-second sample + HTML report
    PS C:\> .\torpor.ps1 -Unattended -SampleSeconds 120   # Longer sample for intermittent slowness

.NOTES
    Version : 5.1

#>

param(
    [switch]$Unattended,
    [ValidateRange(5, 300)]
    [int]$SampleSeconds = 30,
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
# THRESHOLDS
# ─────────────────────────────────────────────────────────────────────────────

# Averages over the sample window, not single-second peaks: a machine that
# touches 100% for a moment is not slow, one that sits there is.
$CpuErrorPercent          = 90
$CpuWarningPercent        = 70
$CpuLimitWarningPercent   = 80     # % Performance Limit below this = held back by power / thermal policy
$QueuePerCoreWarning      = 2      # processor queue length per logical processor
$MemAvailErrorPercent     = 5
$MemAvailWarningPercent   = 15
$CommitErrorPercent       = 90
$CommitWarningPercent     = 80
$DiskBusyErrorPercent     = 90
$DiskBusyWarningPercent   = 60
$DiskLatencyErrorMs       = 50
$DiskLatencyWarningMs     = 25
$SystemDriveErrorPercent  = 5      # free space on the system drive
$SystemDriveWarningPercent = 10
$LowMemoryWarningGB       = 4
$LowMemoryInfoGB          = 8
$UptimeWarningDays        = 14
$SlowBootWarningSeconds   = 120
$AppHangThreshold         = 3      # hangs of one app in $AppHangDays
$AppHangDays              = 7
$BootLookbackDays         = 30
$StartupItemsInfo         = 15
$HeavyProcessCpuPercent   = 25     # one process group's share of the whole machine
$HeavyProcessMemPercent   = 25

$PowerSaverSchemeGuid = 'a1841308-3541-4fab-bc81-f71556f20b4a'

# ─────────────────────────────────────────────────────────────────────────────
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

# Well-known processes and what their load usually means. Keyed by process
# name, lower-case, without .exe.
$ProcessHints = @{
    'msmpeng'             = 'Microsoft Defender scanning. Check for a scheduled or running scan, or a busy folder it keeps rescanning (G.R.I.F.F.I.N. lists exclusions).'
    'mssense'             = 'Microsoft Defender for Endpoint sensor.'
    'tiworker'            = 'Windows servicing installing updates or features. Settles when the install finishes; a restart may be pending.'
    'trustedinstaller'    = 'Windows Modules Installer (servicing). Settles when the update or feature install finishes.'
    'searchindexer'       = 'Windows Search indexing. Heavy after a migration or a large sync; settles once the index catches up.'
    'searchprotocolhost'  = 'Windows Search reading files to index them.'
    'searchfilterhost'    = 'Windows Search extracting file contents to index them.'
    'onedrive'            = 'OneDrive sync. A large first sync or a sync loop; P.H.Y.L.A.C.T.E.R.Y. shows sync errors.'
    'wmiprvse'            = 'WMI provider host. Something is querying WMI hard, often a monitoring or management agent.'
    'svchost'             = 'Windows service host. The services inside it are listed with the process.'
    'compattelrunner'     = 'Windows compatibility telemetry. Runs on a schedule and exits.'
    'mrt'                 = 'Malicious Software Removal Tool, the monthly scan after Patch Tuesday.'
    'officeclicktorun'    = 'Microsoft 365 Apps updating or streaming.'
    'msedge'              = 'Microsoft Edge. Cost grows with open tabs and extensions; Edge''s own Task Manager (Shift+Esc) names the tab.'
    'chrome'              = 'Google Chrome. Cost grows with open tabs and extensions; Chrome''s Task Manager (Shift+Esc) names the tab.'
    'firefox'             = 'Mozilla Firefox. about:processes names the tab.'
    'ms-teams'            = 'Microsoft Teams. Meetings with video and many chats are expensive; clearing the cache helps a client that stays heavy.'
    'teams'               = 'Microsoft Teams (classic). Retired by Microsoft; move to the new Teams client.'
    'outlook'             = 'Outlook. A large mailbox cache, a slow add-in, or a sync in progress.'
    'dwm'                 = 'Desktop Window Manager. High use points at the graphics driver or several high-resolution displays.'
    'system'              = 'Kernel and drivers. Sustained load here usually means a driver (interrupts / DPCs); NECROPSY and LatencyMon-style tools narrow it down.'
    'memory compression'  = 'Memory compression. A symptom of memory pressure, not its cause: see the memory findings.'
    'audiodg'             = 'Windows audio engine. Audio enhancements or a faulty audio driver.'
    'msiexec'             = 'Windows Installer. A software install or repair is running.'
}

# Diagnostics-Performance events that name what slowed a boot. 100 is the boot
# itself; the others each blame one component.
$BootDegradationKinds = @{
    101 = 'Application'
    102 = 'Driver'
    103 = 'Service'
    106 = 'Background optimisation'
    107 = 'Machine Group Policy'
    108 = 'User Group Policy'
    109 = 'Device'
    110 = 'Session manager'
}

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
# ─────────────────────────────────────────────────────────────────────────────

$TorporFindings = @{
    'CpuSaturated' = @{
        Severity = 'Error'
        Title    = 'CPU saturated'
        Summary  = 'The processor averaged 90% or more across the sample. Everything on the machine is waiting for CPU time.'
        Remedy   = 'Start with the top processes table: end or fix the process responsible. If the load is spread across many ordinary apps, the CPU is undersized for the workload.'
    }
    'CpuHigh' = @{
        Severity = 'Warning'
        Title    = 'CPU heavily loaded'
        Summary  = 'The processor averaged 70% or more across the sample. Responsiveness suffers as soon as anything else starts.'
        Remedy   = 'Check the top processes table for a process that should not be that busy.'
    }
    'ProcessorQueue' = @{
        Severity = 'Warning'
        Title    = 'Threads queuing for the CPU'
        Summary  = 'More than two threads per logical processor were waiting to run on average -- the CPU cannot keep up even if the percentage looks moderate.'
        Remedy   = 'Reduce the concurrent load (see the top processes), or move the workload to a machine with more cores.'
    }
    'CpuLimited' = @{
        Severity = 'Warning'
        Title    = 'CPU held below full speed'
        Summary  = 'Windows reported the processor running under a performance limit -- a power plan, battery saver, or thermal throttling.'
        Remedy   = 'Check the power plan and power mode, run on AC power, and clear vents / fans. A laptop that throttles on AC with a clean cooler needs a vendor diagnostic (thermal paste, fan).'
    }
    'MemoryExhausted' = @{
        Severity = 'Error'
        Title    = 'Memory exhausted'
        Summary  = 'Available RAM fell below 5% or the commit charge passed 90% of the commit limit. Windows is paging heavily and allocations may start to fail.'
        Remedy   = 'Close or fix the largest memory consumers in the table. If ordinary use fills RAM, the machine needs more memory.'
    }
    'MemoryPressure' = @{
        Severity = 'Warning'
        Title    = 'Memory under pressure'
        Summary  = 'Available RAM fell below 15% or the commit charge passed 80% of the commit limit. The machine is close to paging.'
        Remedy   = 'Review the largest memory consumers; a browser with many tabs or a leaking process is the usual cause.'
    }
    'LowInstalledMemory' = @{
        Severity = 'Warning'
        Title    = 'Very little RAM installed'
        Summary  = 'The machine has 4 GB of RAM or less, which current Windows and Microsoft 365 outgrow under ordinary use.'
        Remedy   = 'Add memory if the hardware allows it; otherwise plan a replacement.'
    }
    'ModestInstalledMemory' = @{
        Severity = 'Info'
        Title    = 'Under 8 GB of RAM installed'
        Summary  = 'Enough for light use, but a browser, Teams and Outlook together will fill it.'
        Remedy   = 'Consider an upgrade to 16 GB if memory findings recur.'
    }
    'PagefileDisabled' = @{
        Severity = 'Warning'
        Title    = 'No page file'
        Summary  = 'There is no page file and Windows is not managing one, so the commit limit equals physical RAM and memory runs out sooner. Crash dumps also need a page file.'
        Remedy   = 'Re-enable a system-managed page file: System Properties > Advanced > Performance > Virtual memory.'
    }
    'DiskSaturated' = @{
        Severity = 'Error'
        Title    = 'Disk saturated'
        Summary  = 'A physical disk was busy 90% or more of the time, or took 50 ms or longer per request on average. Applications stall waiting on storage.'
        Remedy   = 'Find the process doing the I/O in the top processes table. A spinning disk under ordinary load needs replacing with an SSD; check A.U.G.U.R. for a failing drive.'
    }
    'DiskBusy' = @{
        Severity = 'Warning'
        Title    = 'Disk heavily used'
        Summary  = 'A physical disk was busy 60% or more of the time, or took 25 ms or longer per request on average.'
        Remedy   = 'Check the top processes by I/O. Response times like this on an SSD are worth an A.U.G.U.R. health check.'
    }
    'SpinningSystemDisk' = @{
        Severity = 'Warning'
        Title    = 'Windows is on a spinning hard disk'
        Summary  = 'The system drive is an HDD. Current Windows with Defender, updates and indexing is slow on one regardless of anything else.'
        Remedy   = 'Replace the system disk with an SSD -- usually the single biggest improvement available.'
    }
    'SystemDriveCritical' = @{
        Severity = 'Error'
        Title    = 'System drive almost full'
        Summary  = 'Less than 5% of the system drive is free. Paging, updates and temporary files all compete for the space.'
        Remedy   = 'Free space now: C.L.E.A.N.S.E. clears temp and update caches, F.A.T.H.O.M. finds large folders and old profiles.'
    }
    'SystemDriveLow' = @{
        Severity = 'Warning'
        Title    = 'System drive low on space'
        Summary  = 'Less than 10% of the system drive is free.'
        Remedy   = 'Run C.L.E.A.N.S.E., and F.A.T.H.O.M. to find what is using the space.'
    }
    'LongUptime' = @{
        Severity = 'Warning'
        Title    = 'Not restarted in a long time'
        Summary  = 'The machine has been up for more than two weeks. Leaks and pending updates accumulate; with Fast Startup on, "Shut down" does not reset this -- only Restart does.'
        Remedy   = 'Restart (not Shut down) and re-test before chasing anything else.'
    }
    'PowerSaverPlan' = @{
        Severity = 'Warning'
        Title    = 'Power saving plan or mode active'
        Summary  = 'The Power saver plan, or the Best power efficiency power mode, caps the processor to save energy.'
        Remedy   = 'Switch to Balanced (or Best performance) in Settings > System > Power, unless the battery life is the point.'
    }
    'SlowBoot' = @{
        Severity = 'Warning'
        Title    = 'Slow boots'
        Summary  = 'Recent boots took over two minutes to reach an idle desktop, as measured by Windows itself.'
        Remedy   = 'See the boot degradation table for what Windows blamed. Trim startup programs and check for slow Group Policy or logon scripts.'
    }
    'BootDegradation' = @{
        Severity = 'Info'
        Title    = 'Windows named components that slowed boot'
        Summary  = 'The Diagnostics-Performance log recorded applications, drivers or services that took longer than usual during boot.'
        Remedy   = 'Update or remove the components named most often; a driver that recurs here is worth updating with F.O.R.G.E.'
    }
    'AppHangs' = @{
        Severity = 'Warning'
        Title    = 'Applications repeatedly hanging'
        Summary  = 'An application stopped responding several times in the last week ("Not Responding"), which users report as the machine being slow.'
        Remedy   = 'Repair, update or reset the application named. Hangs in Office apps often trace to an add-in.'
    }
    'ManyStartupItems' = @{
        Severity = 'Info'
        Title    = 'Many programs start at sign-in'
        Summary  = 'More than fifteen enabled startup entries compete for CPU and disk right after sign-in.'
        Remedy   = 'Disable the ones the user does not need in Task Manager > Startup apps.'
    }
    'HeavyProcess' = @{
        Severity = 'Info'
        Title    = 'One program dominates the machine'
        Summary  = 'A single program used a quarter or more of the machine''s CPU or memory during the sample.'
        Remedy   = 'See the hint against the program in the top processes table.'
    }
    'CountersUnavailable' = @{
        Severity = 'Warning'
        Title    = 'Performance counters could not be read'
        Summary  = 'The WMI performance classes returned nothing, so CPU, memory or disk load could not be measured.'
        Remedy   = 'Rebuild the counters from an elevated prompt (lodctr /R, then winmgmt /resyncperf) and run T.O.R.P.O.R. again.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()

function Add-TorporFinding {
    param([Parameter(Mandatory)][string]$Code, [string]$Detail = '')
    $meta = $TorporFindings[$Code]
    if (-not $meta) { $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' } }
    [void]$Findings.Add([PSCustomObject]@{
        Code = $Code; Severity = $meta.Severity; Title = $meta.Title; Summary = $meta.Summary; Remedy = $meta.Remedy; Detail = $Detail
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

function Show-TorporBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ████████╗ ██████╗ ██████╗ ██████╗  ██████╗ ██████╗
  ╚══██╔══╝██╔═══██╗██╔══██╗██╔══██╗██╔═══██╗██╔══██╗
     ██║   ██║   ██║██████╔╝██████╔╝██║   ██║██████╔╝
     ██║   ██║   ██║██╔══██╗██╔═══╝ ██║   ██║██╔══██╗
     ██║   ╚██████╔╝██║  ██║██║     ╚██████╔╝██║  ██║
     ╚═╝    ╚═════╝ ╚═╝  ╚═╝╚═╝      ╚═════╝ ╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    T.O.R.P.O.R. — Traces Overload, Resource Pressure & Offending Routines" -ForegroundColor Cyan
    Write-Host "    Slow-Machine Triage Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PURE HELPERS — no I/O; the Pester suite calls these with captured data.
# ─────────────────────────────────────────────────────────────────────────────

function Get-TorporAverage {
    # Mean of the non-null values, or $null when there are none.
    param([object[]]$Values)
    $nums = @($Values | Where-Object { $null -ne $_ } | ForEach-Object { [double]$_ })
    if ($nums.Count -eq 0) { return $null }
    return ($nums | Measure-Object -Average).Average
}

function Get-TorporLevel {
    # 'Error', 'Warning' or 'Ok' for a measured value against two thresholds.
    # -LowerIsWorse flips the comparison for "free" / "available" style values.
    param($Value, [double]$Warning, [double]$ErrorAt, [switch]$LowerIsWorse)
    if ($null -eq $Value) { return 'Ok' }
    $v = [double]$Value
    if ($LowerIsWorse) {
        if ($v -lt $ErrorAt) { return 'Error' }
        if ($v -lt $Warning) { return 'Warning' }
        return 'Ok'
    }
    if ($v -ge $ErrorAt) { return 'Error' }
    if ($v -ge $Warning) { return 'Warning' }
    return 'Ok'
}

function Get-ProcessHint {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $key = ($Name -replace '(?i)\.exe$', '').Trim().ToLowerInvariant()
    if ($ProcessHints.ContainsKey($key)) { return $ProcessHints[$key] }
    return ''
}

function Get-ProcessCpuUsage {
    # Turns two process snapshots into per-process usage over the window.
    # CPU is the share of the whole machine (all logical processors), so the
    # rows add up to the total CPU figure. A process is matched on Id and Name
    # together so a reused PID is not credited with another process's time; a
    # process that started inside the window is credited with all of its time.
    param([object[]]$Before, [object[]]$After, [double]$ElapsedSeconds, [int]$LogicalProcessors)
    $prior = @{}
    foreach ($b in @($Before)) { $prior["$($b.Id)|$($b.Name)"] = $b }
    $seconds  = [math]::Max(0.001, $ElapsedSeconds)
    $capacity = $seconds * [math]::Max(1, $LogicalProcessors)
    foreach ($a in @($After)) {
        $b = $prior["$($a.Id)|$($a.Name)"]
        $cpuDelta = 0.0
        if ($null -ne $a.CpuSeconds) {
            $start = 0.0
            if ($b -and $null -ne $b.CpuSeconds) { $start = [double]$b.CpuSeconds }
            $cpuDelta = [math]::Max(0.0, [double]$a.CpuSeconds - $start)
        }
        $ioDelta = [double]$a.IoBytes
        if ($b) { $ioDelta = [math]::Max(0.0, [double]$a.IoBytes - [double]$b.IoBytes) }
        [PSCustomObject]@{
            Id            = $a.Id
            Name          = $a.Name
            CpuPercent    = [math]::Round(100.0 * $cpuDelta / $capacity, 1)
            WorkingSet    = [double]$a.WorkingSet
            Private       = [double]$a.Private
            IoBytesPerSec = $ioDelta / $seconds
        }
    }
}

function Group-TorporProcess {
    # One row per program name: thirty msedge processes are one browser.
    param([object[]]$Rows)
    $groups = @($Rows) | Where-Object { $_ } | Group-Object -Property Name
    foreach ($g in $groups) {
        $items = @($g.Group)
        [PSCustomObject]@{
            Name          = $g.Name
            Count         = $items.Count
            CpuPercent    = [math]::Round((($items | Measure-Object -Property CpuPercent -Sum).Sum), 1)
            WorkingSet    = [double](($items | Measure-Object -Property WorkingSet -Sum).Sum)
            Private       = [double](($items | Measure-Object -Property Private -Sum).Sum)
            IoBytesPerSec = [double](($items | Measure-Object -Property IoBytesPerSec -Sum).Sum)
            Ids           = @($items | ForEach-Object { $_.Id })
        }
    }
}

function Get-DiskRawDelta {
    # Averages over the window from two Win32_PerfRawData_PerfDisk_PhysicalDisk
    # snapshots of one disk. Raw counters are used because the formatted class
    # rounds response time to whole seconds.
    #   % Idle Time      PERF_PRECISION_100NS_TIMER -> delta / delta(base)
    #   Avg. sec/Transfer PERF_AVERAGE_TIMER         -> delta(ticks) / freq / delta(ops)
    #   Transfers/sec, Bytes/sec PERF_COUNTER_*     -> delta / elapsed seconds
    param([object]$Before, [object]$After)
    $freq    = [double]$After.Frequency_PerfTime
    $dt100   = [double]$After.Timestamp_Sys100NS - [double]$Before.Timestamp_Sys100NS
    $seconds = 0.0
    if ($freq -gt 0) { $seconds = ([double]$After.Timestamp_PerfTime - [double]$Before.Timestamp_PerfTime) / $freq }
    if ($seconds -le 0 -and $dt100 -gt 0) { $seconds = $dt100 / 1e7 }

    $idleBase = $dt100
    if ($null -ne $After.PercentIdleTime_Base -and $null -ne $Before.PercentIdleTime_Base) {
        $idleBase = [double]$After.PercentIdleTime_Base - [double]$Before.PercentIdleTime_Base
    }
    $busy = $null
    if ($idleBase -gt 0) {
        $idle = 100.0 * ([double]$After.PercentIdleTime - [double]$Before.PercentIdleTime) / $idleBase
        $busy = [math]::Round([math]::Min(100.0, [math]::Max(0.0, 100.0 - $idle)), 1)
    }

    $ops = [double]$After.AvgDisksecPerTransfer_Base - [double]$Before.AvgDisksecPerTransfer_Base
    $latency = 0.0
    if ($ops -gt 0 -and $freq -gt 0) {
        $latency = 1000.0 * (([double]$After.AvgDisksecPerTransfer - [double]$Before.AvgDisksecPerTransfer) / $freq) / $ops
    }

    $iops = 0.0; $bps = 0.0
    if ($seconds -gt 0) {
        $iops = ([double]$After.DiskTransfersPersec - [double]$Before.DiskTransfersPersec) / $seconds
        $bps  = ([double]$After.DiskBytesPersec - [double]$Before.DiskBytesPersec) / $seconds
    }
    return [PSCustomObject]@{
        BusyPercent = $busy
        LatencyMs   = [math]::Round([math]::Max(0.0, $latency), 1)
        Iops        = [math]::Round([math]::Max(0.0, $iops), 1)
        BytesPerSec = [math]::Max(0.0, $bps)
    }
}

function ConvertFrom-PowercfgScheme {
    # powercfg /getactivescheme: "Power Scheme GUID: <guid>  (Balanced)". The
    # label is translated on non-English Windows; the GUID is not.
    param([string[]]$Lines)
    $m = [regex]::Match(($Lines -join "`n"), '(?i)([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\s*(?:\((.+?)\))?')
    if (-not $m.Success) { return [PSCustomObject]@{ Guid = ''; Name = '' } }
    return [PSCustomObject]@{ Guid = $m.Groups[1].Value.ToLowerInvariant(); Name = $m.Groups[2].Value.Trim() }
}

function Get-PowerModeLabel {
    # Windows 10/11 power mode slider overlay, read from the registry.
    param([string]$Guid)
    switch ("$Guid".ToLowerInvariant()) {
        '961cc777-2547-4f9d-8174-7d86181b8a7a' { return 'Best power efficiency' }
        '3af9b8d9-7c97-431d-ad78-34a8bfea439f' { return 'Better performance' }
        'ded574b5-45a0-4f42-8737-46345c09c238' { return 'Best performance' }
        '00000000-0000-0000-0000-000000000000' { return 'Balanced' }
        ''                                     { return '' }
        default                                { return "Unrecognised ($Guid)" }
    }
}

function Test-StartupApprovedEnabled {
    # Explorer\StartupApproved values: first byte even (02, 06) = enabled,
    # odd (03, 07) = disabled in Task Manager. No value means never toggled,
    # which is enabled.
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Count -eq 0) { return $true }
    return (($Bytes[0] -band 1) -eq 0)
}

function Get-BootDegradationKind {
    param([int]$EventId)
    if ($BootDegradationKinds.ContainsKey($EventId)) { return $BootDegradationKinds[$EventId] }
    return 'Other'
}

function Get-TorporVerdict {
    param([object[]]$FindingList)
    $sev = @($FindingList | ForEach-Object { $_.Severity })
    if ($sev -contains 'Error')   { return [PSCustomObject]@{ Verdict = 'Struggling'; Class = 'err'  } }
    if ($sev -contains 'Warning') { return [PSCustomObject]@{ Verdict = 'Strained';   Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Healthy'; Class = 'ok' }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-CimSafe {
    param([string]$ClassName, [string]$Filter = '', [string]$Namespace = 'root/cimv2')
    try {
        if ($Filter) { return @(Get-CimInstance -Namespace $Namespace -ClassName $ClassName -Filter $Filter -ErrorAction Stop) }
        return @(Get-CimInstance -Namespace $Namespace -ClassName $ClassName -ErrorAction Stop)
    } catch {
        return @()
    }
}

function Get-EventsSafe {
    # Get-WinEvent throws when nothing matches; that is an empty result.
    param([hashtable]$Filter, [int]$MaxEvents = 500)
    try {
        return @(Get-WinEvent -FilterHashtable $Filter -MaxEvents $MaxEvents -ErrorAction Stop)
    } catch {
        return @()
    }
}

function Get-EventDataMap {
    # Named EventData values as a hashtable.
    param($Record)
    $map = @{}
    try {
        $xml = [xml]$Record.ToXml()
        foreach ($d in @($xml.Event.EventData.Data)) {
            if ($d -and $d.Name) { $map[$d.Name] = $d.'#text' }
        }
    } catch {
        return $map
    }
    return $map
}

function Get-TorporContext {
    $os  = Get-CimSafe -ClassName 'Win32_OperatingSystem' | Select-Object -First 1
    $cs  = Get-CimSafe -ClassName 'Win32_ComputerSystem'  | Select-Object -First 1
    $cpu = @(Get-CimSafe -ClassName 'Win32_Processor')
    $bat = @(Get-CimSafe -ClassName 'Win32_Battery')

    $uptime = $null
    if ($os -and $os.LastBootUpTime) { $uptime = (Get-Date) - $os.LastBootUpTime }

    $hiberboot = $null
    try {
        $hiberboot = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' -Name 'HiberbootEnabled' -ErrorAction Stop).HiberbootEnabled
    } catch { $hiberboot = $null }

    $logical = 1
    if ($cs -and $cs.NumberOfLogicalProcessors) { $logical = [int]$cs.NumberOfLogicalProcessors }

    $onBattery = $false
    if ($bat.Count -gt 0) { $onBattery = [bool](@($bat | Where-Object { $_.BatteryStatus -eq 1 }).Count) }

    return [PSCustomObject]@{
        Computer       = $env:COMPUTERNAME
        OS             = $(if ($os) { "$($os.Caption) (build $($os.BuildNumber))" } else { 'Unknown' })
        Model          = $(if ($cs) { ("$($cs.Manufacturer) $($cs.Model)").Trim() } else { 'Unknown' })
        Cpu            = $(if ($cpu.Count) { ($cpu[0].Name -replace '\s+', ' ').Trim() } else { 'Unknown' })
        Cores          = [int](($cpu | Measure-Object -Property NumberOfCores -Sum).Sum)
        Logical        = $logical
        TotalRamBytes  = $(if ($cs -and $cs.TotalPhysicalMemory) { [double]$cs.TotalPhysicalMemory } else { 0 })
        AutoPagefile   = $(if ($cs) { [bool]$cs.AutomaticManagedPagefile } else { $true })
        LastBoot       = $(if ($os) { $os.LastBootUpTime } else { $null })
        Uptime         = $uptime
        FastStartup    = ($hiberboot -eq 1)
        HasBattery     = ($bat.Count -gt 0)
        OnBattery      = $onBattery
        CommitLimitKB  = $(if ($os) { [double]$os.TotalVirtualMemorySize } else { 0 })
        CommitFreeKB   = $(if ($os) { [double]$os.FreeVirtualMemory } else { 0 })
    }
}

function Get-TorporProcessSnapshot {
    $io = @{}
    foreach ($p in (Get-CimSafe -ClassName 'Win32_Process')) {
        $io[[int]$p.ProcessId] = [double]$p.ReadTransferCount + [double]$p.WriteTransferCount + [double]$p.OtherTransferCount
    }
    $rows = foreach ($p in @(Get-Process -ErrorAction SilentlyContinue)) {
        if ($p.Id -eq 0) { continue }   # the Idle pseudo-process
        $cpu = $null
        try { if ($p.TotalProcessorTime) { $cpu = $p.TotalProcessorTime.TotalSeconds } } catch { $cpu = $null }
        $pid32 = [int]$p.Id
        [PSCustomObject]@{
            Id         = $pid32
            Name       = $p.ProcessName
            CpuSeconds = $cpu
            WorkingSet = [double]$p.WorkingSet64
            Private    = [double]$p.PrivateMemorySize64
            IoBytes    = $(if ($io.ContainsKey($pid32)) { $io[$pid32] } else { 0 })
        }
    }
    return @($rows)
}

function Get-TorporDiskRaw {
    $rows = @{}
    foreach ($d in (Get-CimSafe -ClassName 'Win32_PerfRawData_PerfDisk_PhysicalDisk')) {
        if ($d.Name -eq '_Total') { continue }
        $rows[$d.Name] = $d
    }
    return $rows
}

function Invoke-TorporSample {
    # Samples the system-wide counters once a second and snapshots processes
    # and disks at both ends of the window.
    param([int]$Seconds, [int]$LogicalProcessors)

    $cpuSamples   = [System.Collections.Generic.List[object]]::new()
    $queueSamples = [System.Collections.Generic.List[object]]::new()
    $availSamples = [System.Collections.Generic.List[object]]::new()
    $commitSamples = [System.Collections.Generic.List[object]]::new()
    $limitSamples = [System.Collections.Generic.List[object]]::new()
    $faultSamples = [System.Collections.Generic.List[object]]::new()

    # Formatted rate counters need a previous sample inside the provider; the
    # first read after a pause can come back zero, so it is discarded.
    $null = Get-CimSafe -ClassName 'Win32_PerfFormattedData_PerfOS_Processor' -Filter "Name='_Total'"
    $null = Get-CimSafe -ClassName 'Win32_PerfFormattedData_PerfOS_Memory'
    $null = Get-CimSafe -ClassName 'Win32_PerfFormattedData_Counters_ProcessorInformation' -Filter "Name='_Total'"

    $procBefore = Get-TorporProcessSnapshot
    $diskBefore = Get-TorporDiskRaw
    $watch = [System.Diagnostics.Stopwatch]::StartNew()

    for ($i = 1; $i -le $Seconds; $i++) {
        $tick = [System.Diagnostics.Stopwatch]::StartNew()
        Write-Progress -Activity 'T.O.R.P.O.R. sampling' -Status "Second $i of $Seconds" -PercentComplete ([int](100 * $i / $Seconds))

        $p = Get-CimSafe -ClassName 'Win32_PerfFormattedData_PerfOS_Processor' -Filter "Name='_Total'" | Select-Object -First 1
        if ($p) { [void]$cpuSamples.Add([double]$p.PercentProcessorTime) }

        $s = Get-CimSafe -ClassName 'Win32_PerfFormattedData_PerfOS_System' | Select-Object -First 1
        if ($s) { [void]$queueSamples.Add([double]$s.ProcessorQueueLength) }

        $m = Get-CimSafe -ClassName 'Win32_PerfFormattedData_PerfOS_Memory' | Select-Object -First 1
        if ($m) {
            [void]$availSamples.Add([double]$m.AvailableMBytes)
            [void]$commitSamples.Add([double]$m.PercentCommittedBytesInUse)
            [void]$faultSamples.Add([double]$m.PagesInputPersec)
        }

        $pi = Get-CimSafe -ClassName 'Win32_PerfFormattedData_Counters_ProcessorInformation' -Filter "Name='_Total'" | Select-Object -First 1
        if ($pi -and $null -ne $pi.PercentPerformanceLimit) { [void]$limitSamples.Add([double]$pi.PercentPerformanceLimit) }

        $remaining = 1000 - $tick.ElapsedMilliseconds
        if ($remaining -gt 0 -and $i -lt $Seconds) { Start-Sleep -Milliseconds $remaining }
    }
    Write-Progress -Activity 'T.O.R.P.O.R. sampling' -Completed

    $elapsed    = $watch.Elapsed.TotalSeconds
    $procAfter  = Get-TorporProcessSnapshot
    $diskAfter  = Get-TorporDiskRaw

    $disks = foreach ($name in $diskAfter.Keys) {
        if (-not $diskBefore.ContainsKey($name)) { continue }
        $delta = Get-DiskRawDelta -Before $diskBefore[$name] -After $diskAfter[$name]
        [PSCustomObject]@{
            Instance    = $name
            Number      = $(if ($name -match '^(\d+)') { $Matches[1] } else { '' })
            BusyPercent = $delta.BusyPercent
            LatencyMs   = $delta.LatencyMs
            Iops        = $delta.Iops
            BytesPerSec = $delta.BytesPerSec
        }
    }

    $processes = @(Get-ProcessCpuUsage -Before $procBefore -After $procAfter -ElapsedSeconds $elapsed -LogicalProcessors $LogicalProcessors)

    return [PSCustomObject]@{
        Elapsed      = $elapsed
        CpuAvg       = Get-TorporAverage -Values $cpuSamples.ToArray()
        CpuPeak      = $(if ($cpuSamples.Count) { ($cpuSamples | Measure-Object -Maximum).Maximum } else { $null })
        QueueAvg     = Get-TorporAverage -Values $queueSamples.ToArray()
        AvailAvgMB   = Get-TorporAverage -Values $availSamples.ToArray()
        AvailMinMB   = $(if ($availSamples.Count) { ($availSamples | Measure-Object -Minimum).Minimum } else { $null })
        CommitAvg    = Get-TorporAverage -Values $commitSamples.ToArray()
        CommitPeak   = $(if ($commitSamples.Count) { ($commitSamples | Measure-Object -Maximum).Maximum } else { $null })
        FaultsAvg    = Get-TorporAverage -Values $faultSamples.ToArray()
        LimitAvg     = Get-TorporAverage -Values $limitSamples.ToArray()
        SampleCount  = $cpuSamples.Count
        Disks        = @($disks)
        Processes    = $processes
        Groups       = @(Group-TorporProcess -Rows $processes)
    }
}

function Get-TorporPhysicalDiskInfo {
    # DeviceId -> media type / bus / model, to label the perf counter instances.
    $map = @{}
    try {
        foreach ($d in @(Get-PhysicalDisk -ErrorAction Stop)) {
            $map["$($d.DeviceId)"] = [PSCustomObject]@{ MediaType = "$($d.MediaType)"; BusType = "$($d.BusType)"; Model = "$($d.FriendlyName)" }
        }
    } catch {
        return $map
    }
    return $map
}

function Get-TorporSystemDrive {
    $letter = "$env:SystemDrive"
    $v = Get-CimSafe -ClassName 'Win32_LogicalDisk' -Filter "DeviceID='$letter'" | Select-Object -First 1
    if (-not $v -or -not $v.Size) { return $null }
    return [PSCustomObject]@{
        Drive       = $letter
        SizeBytes   = [double]$v.Size
        FreeBytes   = [double]$v.FreeSpace
        FreePercent = [math]::Round(100.0 * [double]$v.FreeSpace / [double]$v.Size, 1)
    }
}

function Get-TorporPowerState {
    $scheme = ConvertFrom-PowercfgScheme -Lines @(& powercfg.exe /getactivescheme 2>$null | ForEach-Object { "$_" })
    $overlay = ''
    try {
        $overlay = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes' -Name 'ActiveOverlayAcPowerScheme' -ErrorAction Stop).ActiveOverlayAcPowerScheme
    } catch { $overlay = '' }
    return [PSCustomObject]@{
        SchemeGuid = $scheme.Guid
        SchemeName = $scheme.Name
        PowerMode  = Get-PowerModeLabel -Guid "$overlay"
    }
}

function Get-TorporStartupItem {
    $items = [System.Collections.Generic.List[object]]::new()
    $runKeys = @(
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run';             Scope = 'All users';          Approved = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run' }
        @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'; Scope = 'All users (32-bit)'; Approved = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32' }
        @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run';             Scope = 'Current user';       Approved = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run' }
    )
    foreach ($rk in $runKeys) {
        $key = Get-Item -Path $rk.Path -ErrorAction SilentlyContinue
        if (-not $key) { continue }
        $approved = Get-Item -Path $rk.Approved -ErrorAction SilentlyContinue
        foreach ($name in $key.GetValueNames()) {
            if ([string]::IsNullOrEmpty($name)) { continue }
            $bytes = $null
            if ($approved) { $bytes = [byte[]]$approved.GetValue($name, $null) }
            [void]$items.Add([PSCustomObject]@{
                Name    = $name
                Scope   = $rk.Scope
                Source  = 'Run key'
                Command = "$($key.GetValue($name, ''))"
                Enabled = Test-StartupApprovedEnabled -Bytes $bytes
            })
        }
    }

    $folders = @(
        @{ Path = [Environment]::GetFolderPath('CommonStartup'); Scope = 'All users';    Approved = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder' }
        @{ Path = [Environment]::GetFolderPath('Startup');       Scope = 'Current user'; Approved = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder' }
    )
    foreach ($f in $folders) {
        if ([string]::IsNullOrWhiteSpace($f.Path) -or -not (Test-Path -LiteralPath $f.Path)) { continue }
        $approved = Get-Item -Path $f.Approved -ErrorAction SilentlyContinue
        foreach ($file in @(Get-ChildItem -LiteralPath $f.Path -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne 'desktop.ini' })) {
            $bytes = $null
            if ($approved) { $bytes = [byte[]]$approved.GetValue($file.Name, $null) }
            [void]$items.Add([PSCustomObject]@{
                Name    = $file.BaseName
                Scope   = $f.Scope
                Source  = 'Startup folder'
                Command = $file.FullName
                Enabled = Test-StartupApprovedEnabled -Bytes $bytes
            })
        }
    }
    return $items.ToArray()
}

function Get-TorporBootHistory {
    $since  = (Get-Date).AddDays(-$BootLookbackDays)
    $log    = 'Microsoft-Windows-Diagnostics-Performance/Operational'
    $boots  = foreach ($e in (Get-EventsSafe -Filter @{ LogName = $log; Id = 100; StartTime = $since } -MaxEvents 20)) {
        $d = Get-EventDataMap -Record $e
        $total = 0.0
        if ($d['BootTime']) { $total = [double]$d['BootTime'] / 1000.0 }
        $main = 0.0
        if ($d['MainPathBootTime']) { $main = [double]$d['MainPathBootTime'] / 1000.0 }
        [PSCustomObject]@{ Time = $e.TimeCreated; TotalSeconds = [math]::Round($total, 1); MainPathSeconds = [math]::Round($main, 1) }
    }

    $ids = @($BootDegradationKinds.Keys | ForEach-Object { [int]$_ })
    $blamed = foreach ($e in (Get-EventsSafe -Filter @{ LogName = $log; Id = $ids; StartTime = $since } -MaxEvents 200)) {
        $d = Get-EventDataMap -Record $e
        $name = $d['FriendlyName']
        if ([string]::IsNullOrWhiteSpace($name)) { $name = $d['Name'] }
        if ([string]::IsNullOrWhiteSpace($name)) { $name = $d['FileName'] }
        if ([string]::IsNullOrWhiteSpace($name)) { $name = '(unnamed)' }
        $deg = 0.0
        if ($d['DegradationTime']) { $deg = [double]$d['DegradationTime'] / 1000.0 }
        [PSCustomObject]@{ Kind = (Get-BootDegradationKind -EventId $e.Id); Name = $name; DegradationSeconds = $deg }
    }
    $grouped = foreach ($g in (@($blamed) | Where-Object { $_ } | Group-Object -Property Kind, Name)) {
        $first = $g.Group[0]
        [PSCustomObject]@{
            Kind     = $first.Kind
            Name     = $first.Name
            Count    = $g.Count
            WorstSec = [math]::Round((($g.Group | Measure-Object -Property DegradationSeconds -Maximum).Maximum), 1)
        }
    }
    return [PSCustomObject]@{
        Boots  = @($boots | Sort-Object Time -Descending)
        Blamed = @($grouped | Sort-Object Count, WorstSec -Descending)
    }
}

function Get-TorporAppHang {
    $since = (Get-Date).AddDays(-$AppHangDays)
    $rows = foreach ($e in (Get-EventsSafe -Filter @{ LogName = 'Application'; ProviderName = 'Application Hang'; Id = 1002; StartTime = $since } -MaxEvents 500)) {
        $app = ''
        try { $app = "$($e.Properties[0].Value)" } catch { $app = '' }
        if ([string]::IsNullOrWhiteSpace($app)) { continue }
        [PSCustomObject]@{ App = $app; Time = $e.TimeCreated }
    }
    $grouped = foreach ($g in (@($rows) | Where-Object { $_ } | Group-Object -Property App)) {
        [PSCustomObject]@{
            App  = $g.Name
            Count = $g.Count
            Last = ($g.Group | Sort-Object Time -Descending | Select-Object -First 1).Time
        }
    }
    return @($grouped | Sort-Object Count -Descending)
}

function Get-TorporSvchostService {
    # PID -> service names, for the svchost rows in the process table.
    param([int[]]$ProcessIds)
    $map = @{}
    if (-not $ProcessIds -or $ProcessIds.Count -eq 0) { return $map }
    foreach ($s in (Get-CimSafe -ClassName 'Win32_Service' -Filter 'ProcessId > 0')) {
        $id = [int]$s.ProcessId
        if ($ProcessIds -notcontains $id) { continue }
        if (-not $map.ContainsKey($id)) { $map[$id] = [System.Collections.Generic.List[string]]::new() }
        [void]$map[$id].Add("$($s.Name)")
    }
    return $map
}

function Get-TorporPagefile {
    return @(Get-CimSafe -ClassName 'Win32_PageFileUsage' | ForEach-Object {
        [PSCustomObject]@{ Path = "$($_.Name)"; AllocatedMB = [int]$_.AllocatedBaseSize; CurrentMB = [int]$_.CurrentUsage; PeakMB = [int]$_.PeakUsage }
    })
}

# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-TorporTriage {
    param([int]$Seconds)

    Write-Section "MACHINE"
    $ctx = Get-TorporContext
    Write-Info "$($ctx.Computer)  |  $($ctx.Model)  |  $($ctx.OS)"
    Write-Info ("CPU: {0}  ({1} cores / {2} logical)" -f $ctx.Cpu, $ctx.Cores, $ctx.Logical)
    Write-Info ("RAM: {0}" -f (Format-Bytes $ctx.TotalRamBytes))
    if ($ctx.Uptime) { Write-Info ("Up for {0:N1} days{1}" -f $ctx.Uptime.TotalDays, $(if ($ctx.FastStartup) { ' (Fast Startup on: Shut down does not reset this)' } else { '' })) }

    Write-Section "SAMPLING ($Seconds s)"
    Write-Step "Watching CPU, memory, disk and processes. Leave the machine doing whatever is slow."
    $sample = Invoke-TorporSample -Seconds $Seconds -LogicalProcessors $ctx.Logical
    if ($sample.SampleCount -eq 0) { Add-TorporFinding -Code 'CountersUnavailable' -Detail 'Win32_PerfFormattedData_PerfOS_Processor returned no data.' }

    $diskInfo = Get-TorporPhysicalDiskInfo
    $sysDrive = Get-TorporSystemDrive
    $power    = Get-TorporPowerState
    $pagefile = Get-TorporPagefile
    Write-Step "Reading startup programs, boot history and application hangs..."
    $startup  = @(Get-TorporStartupItem)
    $boot     = Get-TorporBootHistory
    $hangs    = @(Get-TorporAppHang)

    # ── CPU ─────────────────────────────────────────────────────────────────
    Write-Section "CPU"
    if ($null -ne $sample.CpuAvg) {
        Write-Info ("Average {0:N0}%  |  peak {1:N0}%  |  queue {2:N1}" -f $sample.CpuAvg, $sample.CpuPeak, $sample.QueueAvg)
        switch (Get-TorporLevel -Value $sample.CpuAvg -Warning $CpuWarningPercent -ErrorAt $CpuErrorPercent) {
            'Error'   { Add-TorporFinding -Code 'CpuSaturated' -Detail ("Average {0:N0}%, peak {1:N0}% over {2:N0} s" -f $sample.CpuAvg, $sample.CpuPeak, $sample.Elapsed) }
            'Warning' { Add-TorporFinding -Code 'CpuHigh'      -Detail ("Average {0:N0}%, peak {1:N0}% over {2:N0} s" -f $sample.CpuAvg, $sample.CpuPeak, $sample.Elapsed) }
        }
    }
    if ($null -ne $sample.QueueAvg -and $sample.QueueAvg -gt ($QueuePerCoreWarning * $ctx.Logical)) {
        Add-TorporFinding -Code 'ProcessorQueue' -Detail ("{0:N1} waiting threads on {1} logical processors" -f $sample.QueueAvg, $ctx.Logical)
    }
    if ($null -ne $sample.LimitAvg) {
        Write-Info ("Performance limit {0:N0}% (100 = unrestricted)" -f $sample.LimitAvg)
        if ($sample.LimitAvg -lt $CpuLimitWarningPercent) {
            Add-TorporFinding -Code 'CpuLimited' -Detail ("Average performance limit {0:N0}%{1}" -f $sample.LimitAvg, $(if ($ctx.OnBattery) { '; running on battery' } else { '' }))
        }
    }

    # ── Memory ──────────────────────────────────────────────────────────────
    Write-Section "MEMORY"
    $totalMB = $ctx.TotalRamBytes / 1MB
    $availPct = $null
    if ($totalMB -gt 0 -and $null -ne $sample.AvailMinMB) { $availPct = [math]::Round(100.0 * $sample.AvailMinMB / $totalMB, 1) }
    if ($null -ne $availPct) { Write-Info ("Available: lowest {0:N0} MB ({1}% of RAM)  |  commit avg {2:N0}%, peak {3:N0}%" -f $sample.AvailMinMB, $availPct, $sample.CommitAvg, $sample.CommitPeak) }
    $availLevel  = Get-TorporLevel -Value $availPct -Warning $MemAvailWarningPercent -ErrorAt $MemAvailErrorPercent -LowerIsWorse
    $commitLevel = Get-TorporLevel -Value $sample.CommitPeak -Warning $CommitWarningPercent -ErrorAt $CommitErrorPercent
    $memDetail = ("Lowest available {0:N0} MB ({1}% of RAM); commit peaked at {2:N0}% of the limit" -f $sample.AvailMinMB, $availPct, $sample.CommitPeak)
    if ($availLevel -eq 'Error' -or $commitLevel -eq 'Error')         { Add-TorporFinding -Code 'MemoryExhausted' -Detail $memDetail }
    elseif ($availLevel -eq 'Warning' -or $commitLevel -eq 'Warning') { Add-TorporFinding -Code 'MemoryPressure'  -Detail $memDetail }

    $ramGB = [math]::Round($ctx.TotalRamBytes / 1GB, 1)
    if ($ramGB -gt 0 -and $ramGB -le $LowMemoryWarningGB)   { Add-TorporFinding -Code 'LowInstalledMemory'    -Detail "$ramGB GB installed" }
    elseif ($ramGB -gt 0 -and $ramGB -lt $LowMemoryInfoGB)  { Add-TorporFinding -Code 'ModestInstalledMemory' -Detail "$ramGB GB installed" }
    if ($pagefile.Count -eq 0 -and -not $ctx.AutoPagefile) { Add-TorporFinding -Code 'PagefileDisabled' -Detail 'No page file configured and automatic management is off.' }

    # ── Disk ────────────────────────────────────────────────────────────────
    Write-Section "DISK"
    $sysLetter = "$env:SystemDrive"
    $disks = foreach ($d in $sample.Disks) {
        $info = $diskInfo[$d.Number]
        $isSystem = $d.Instance -match [regex]::Escape($sysLetter)
        $row = [PSCustomObject]@{
            Instance    = $d.Instance
            Model       = $(if ($info) { $info.Model } else { '' })
            MediaType   = $(if ($info) { $info.MediaType } else { 'Unknown' })
            BusType     = $(if ($info) { $info.BusType } else { '' })
            IsSystem    = $isSystem
            BusyPercent = $d.BusyPercent
            LatencyMs   = $d.LatencyMs
            Iops        = $d.Iops
            BytesPerSec = $d.BytesPerSec
        }
        Write-Info ("{0,-12} {1,-8} busy {2,5}%  {3,6} ms  {4,7} IOPS  {5}/s" -f $row.Instance, $row.MediaType, $row.BusyPercent, $row.LatencyMs, $row.Iops, (Format-Bytes $row.BytesPerSec))
        $level = 'Ok'
        foreach ($l in @((Get-TorporLevel -Value $row.BusyPercent -Warning $DiskBusyWarningPercent -ErrorAt $DiskBusyErrorPercent),
                         (Get-TorporLevel -Value $row.LatencyMs -Warning $DiskLatencyWarningMs -ErrorAt $DiskLatencyErrorMs))) {
            if ($l -eq 'Error') { $level = 'Error' } elseif ($l -eq 'Warning' -and $level -ne 'Error') { $level = 'Warning' }
        }
        $detail = ("Disk {0}: busy {1}%, {2} ms per request" -f $row.Instance, $row.BusyPercent, $row.LatencyMs)
        if ($level -eq 'Error')   { Add-TorporFinding -Code 'DiskSaturated' -Detail $detail }
        if ($level -eq 'Warning') { Add-TorporFinding -Code 'DiskBusy'      -Detail $detail }
        if ($isSystem -and $row.MediaType -eq 'HDD') { Add-TorporFinding -Code 'SpinningSystemDisk' -Detail "$($row.Model) ($($row.BusType))" }
        $row
    }
    $disks = @($disks)
    if ($sysDrive) {
        Write-Info ("{0} free: {1} of {2} ({3}%)" -f $sysDrive.Drive, (Format-Bytes $sysDrive.FreeBytes), (Format-Bytes $sysDrive.SizeBytes), $sysDrive.FreePercent)
        switch (Get-TorporLevel -Value $sysDrive.FreePercent -Warning $SystemDriveWarningPercent -ErrorAt $SystemDriveErrorPercent -LowerIsWorse) {
            'Error'   { Add-TorporFinding -Code 'SystemDriveCritical' -Detail ("{0} free on {1} ({2}%)" -f (Format-Bytes $sysDrive.FreeBytes), $sysDrive.Drive, $sysDrive.FreePercent) }
            'Warning' { Add-TorporFinding -Code 'SystemDriveLow'      -Detail ("{0} free on {1} ({2}%)" -f (Format-Bytes $sysDrive.FreeBytes), $sysDrive.Drive, $sysDrive.FreePercent) }
        }
    }

    # ── Processes ───────────────────────────────────────────────────────────
    Write-Section "TOP PROCESSES"
    $groups = @($sample.Groups)
    $topCpu = @($groups | Sort-Object CpuPercent -Descending | Select-Object -First 10)
    $topMem = @($groups | Sort-Object Private -Descending | Select-Object -First 10)
    $topIo  = @($groups | Where-Object { $_.IoBytesPerSec -gt 0 } | Sort-Object IoBytesPerSec -Descending | Select-Object -First 10)

    $svchostIds = @($groups | Where-Object { $_.Name -eq 'svchost' } | ForEach-Object { $_.Ids } | ForEach-Object { [int]$_ })
    $busySvchost = @($sample.Processes | Where-Object { $_.Name -eq 'svchost' -and $_.CpuPercent -ge 1 } | ForEach-Object { [int]$_.Id })
    $svcMap = Get-TorporSvchostService -ProcessIds $(if ($busySvchost.Count) { $busySvchost } else { $svchostIds })
    $busySvchostRows = @($sample.Processes | Where-Object { $_.Name -eq 'svchost' -and $_.CpuPercent -ge 1 } | Sort-Object CpuPercent -Descending | ForEach-Object {
        [PSCustomObject]@{ Id = $_.Id; CpuPercent = $_.CpuPercent; Services = $(if ($svcMap.ContainsKey([int]$_.Id)) { ($svcMap[[int]$_.Id] -join ', ') } else { '' }) }
    })

    foreach ($g in ($topCpu | Select-Object -First 5)) {
        Write-Info ("{0,-28} CPU {1,5}%  RAM {2,10}  x{3}" -f $g.Name, $g.CpuPercent, (Format-Bytes $g.Private), $g.Count)
    }
    $totalRam = [math]::Max(1.0, $ctx.TotalRamBytes)
    foreach ($g in $groups) {
        $memShare = [math]::Round(100.0 * $g.Private / $totalRam, 1)
        if ($g.CpuPercent -ge $HeavyProcessCpuPercent -or $memShare -ge $HeavyProcessMemPercent) {
            Add-TorporFinding -Code 'HeavyProcess' -Detail ("{0} (x{1}): {2}% CPU, {3} private memory ({4}% of RAM)" -f $g.Name, $g.Count, $g.CpuPercent, (Format-Bytes $g.Private), $memShare)
        }
    }

    # ── Context ─────────────────────────────────────────────────────────────
    Write-Section "CONTEXT"
    if ($ctx.Uptime -and $ctx.Uptime.TotalDays -gt $UptimeWarningDays) {
        Add-TorporFinding -Code 'LongUptime' -Detail ("Up {0:N0} days since {1}{2}" -f $ctx.Uptime.TotalDays, $ctx.LastBoot, $(if ($ctx.FastStartup) { '; Fast Startup is on' } else { '' }))
    }
    Write-Info ("Power plan: {0}{1}" -f $(if ($power.SchemeName) { $power.SchemeName } else { $power.SchemeGuid }), $(if ($power.PowerMode) { "  |  power mode: $($power.PowerMode)" } else { '' }))
    if ($power.SchemeGuid -eq $PowerSaverSchemeGuid -or $power.PowerMode -eq 'Best power efficiency') {
        Add-TorporFinding -Code 'PowerSaverPlan' -Detail ("Plan: {0}; power mode: {1}" -f $(if ($power.SchemeName) { $power.SchemeName } else { $power.SchemeGuid }), $(if ($power.PowerMode) { $power.PowerMode } else { 'n/a' }))
    }

    $enabledStartup = @($startup | Where-Object { $_.Enabled })
    Write-Info ("Startup programs: {0} enabled, {1} disabled" -f $enabledStartup.Count, ($startup.Count - $enabledStartup.Count))
    if ($enabledStartup.Count -gt $StartupItemsInfo) { Add-TorporFinding -Code 'ManyStartupItems' -Detail "$($enabledStartup.Count) enabled startup entries" }

    if ($boot.Boots.Count -gt 0) {
        $recent = @($boot.Boots | Select-Object -First 5)
        $avgBoot = Get-TorporAverage -Values @($recent | ForEach-Object { $_.TotalSeconds })
        Write-Info ("Boot time (last {0}): average {1:N0} s" -f $recent.Count, $avgBoot)
        if ($avgBoot -gt $SlowBootWarningSeconds) { Add-TorporFinding -Code 'SlowBoot' -Detail ("Average {0:N0} s over the last {1} boot(s)" -f $avgBoot, $recent.Count) }
    }
    if ($boot.Blamed.Count -gt 0) {
        Add-TorporFinding -Code 'BootDegradation' -Detail (($boot.Blamed | Select-Object -First 3 | ForEach-Object { "$($_.Kind): $($_.Name) (x$($_.Count))" }) -join '; ')
    }

    $hangers = @($hangs | Where-Object { $_.Count -ge $AppHangThreshold })
    if ($hangers.Count -gt 0) {
        Add-TorporFinding -Code 'AppHangs' -Detail (($hangers | Select-Object -First 3 | ForEach-Object { "$($_.App) x$($_.Count)" }) -join '; ')
    }

    return [PSCustomObject]@{
        Context      = $ctx
        Sample       = $sample
        AvailPct     = $availPct
        Disks        = $disks
        SystemDrive  = $sysDrive
        Power        = $power
        Pagefile     = $pagefile
        Startup      = $startup
        Boot         = $boot
        Hangs        = $hangs
        TopCpu       = $topCpu
        TopMem       = $topMem
        TopIo        = $topIo
        Svchost      = $busySvchostRows
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Get-LevelBadge {
    param([string]$Level, [string]$Text)
    $cls = switch ($Level) { 'Error' { 'err' } 'Warning' { 'warn' } default { 'ok' } }
    return "<span class='tk-badge-$cls'>$(EscHtml $Text)</span>"
}

function Build-TorporProcessRow {
    param([object[]]$Groups, [double]$TotalRam)
    $sb = [System.Text.StringBuilder]::new()
    if (@($Groups).Count -eq 0) { [void]$sb.Append("<tr><td colspan='6'>No process data.</td></tr>") }
    foreach ($g in @($Groups)) {
        $hint = Get-ProcessHint -Name $g.Name
        $memShare = if ($TotalRam -gt 0) { [math]::Round(100.0 * $g.Private / $TotalRam, 1) } else { 0 }
        [void]$sb.Append(
            "<tr><td><strong>$(EscHtml $g.Name)</strong>$(if ($g.Count -gt 1) { " <span class='tk-mono'>x$($g.Count)</span>" })</td>" +
            "<td>$($g.CpuPercent)%</td><td>$(EscHtml (Format-Bytes $g.Private)) <span class='tk-mono'>($memShare%)</span></td>" +
            "<td>$(EscHtml (Format-Bytes $g.WorkingSet))</td><td>$(EscHtml (Format-Bytes $g.IoBytesPerSec))/s</td>" +
            "<td>$(EscHtml $hint)</td></tr>")
    }
    return $sb.ToString()
}

function Build-TorporReport {
    param([object]$Result, [object]$Verdict)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $machine    = $env:COMPUTERNAME
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $ctx        = $Result.Context
    $s          = $Result.Sample

    $fRows = [System.Text.StringBuilder]::new()
    if ($Findings.Count -eq 0) { [void]$fRows.Append("<tr><td colspan='4'>Nothing slowing this machine was found during the sample.</td></tr>") }
    foreach ($f in ($Findings | Sort-Object @{ Expression = { switch ($_.Severity) { 'Error' { 0 } 'Warning' { 1 } default { 2 } } } })) {
        [void]$fRows.Append(
            "<tr><td><span class='tk-badge-$(Get-SeverityClass $f.Severity)'>$(EscHtml $f.Severity)</span></td>" +
            "<td class='tk-mono'>$(EscHtml $f.Code)</td>" +
            "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)" +
            $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
            "<td>$(EscHtml $f.Remedy)</td></tr>")
    }

    $cpuLevel    = Get-TorporLevel -Value $s.CpuAvg -Warning $CpuWarningPercent -ErrorAt $CpuErrorPercent
    $availLevel  = Get-TorporLevel -Value $Result.AvailPct -Warning $MemAvailWarningPercent -ErrorAt $MemAvailErrorPercent -LowerIsWorse
    $commitLevel = Get-TorporLevel -Value $s.CommitPeak -Warning $CommitWarningPercent -ErrorAt $CommitErrorPercent
    $limitLevel  = if ($null -ne $s.LimitAvg -and $s.LimitAvg -lt $CpuLimitWarningPercent) { 'Warning' } else { 'Ok' }

    $fmt = { param($v, $suffix) if ($null -eq $v) { 'n/a' } else { ('{0:N0}{1}' -f $v, $suffix) } }

    $loadRows = @(
        "<tr><td>CPU (average / peak)</td><td>$(& $fmt $s.CpuAvg '%') / $(& $fmt $s.CpuPeak '%')</td><td>$(Get-LevelBadge $cpuLevel $cpuLevel)</td><td>Warning at $CpuWarningPercent%, error at $CpuErrorPercent% average</td></tr>"
        "<tr><td>Processor queue (average)</td><td>$(if ($null -ne $s.QueueAvg) { '{0:N1}' -f $s.QueueAvg } else { 'n/a' })</td><td>$(Get-LevelBadge $(if ($null -ne $s.QueueAvg -and $s.QueueAvg -gt ($QueuePerCoreWarning * $ctx.Logical)) { 'Warning' } else { 'Ok' }) $(if ($null -ne $s.QueueAvg -and $s.QueueAvg -gt ($QueuePerCoreWarning * $ctx.Logical)) { 'Warning' } else { 'Ok' }))</td><td>Warning above $QueuePerCoreWarning per logical processor ($($ctx.Logical))</td></tr>"
        "<tr><td>CPU performance limit</td><td>$(& $fmt $s.LimitAvg '%')</td><td>$(Get-LevelBadge $limitLevel $limitLevel)</td><td>100% = unrestricted; below $CpuLimitWarningPercent% = power or thermal limit</td></tr>"
        "<tr><td>Available memory (lowest)</td><td>$(& $fmt $s.AvailMinMB ' MB') ($(& $fmt $Result.AvailPct '%'))</td><td>$(Get-LevelBadge $availLevel $availLevel)</td><td>Warning below $MemAvailWarningPercent%, error below $MemAvailErrorPercent% of RAM</td></tr>"
        "<tr><td>Commit charge (average / peak)</td><td>$(& $fmt $s.CommitAvg '%') / $(& $fmt $s.CommitPeak '%')</td><td>$(Get-LevelBadge $commitLevel $commitLevel)</td><td>Share of the commit limit (RAM + page file)</td></tr>"
        "<tr><td>Hard page faults</td><td>$(& $fmt $s.FaultsAvg ' pages/s')</td><td><span class='tk-badge-info'>Info</span></td><td>Pages read from disk to satisfy memory; sustained hundreds+ alongside low memory = paging</td></tr>"
    ) -join ''

    $dRows = [System.Text.StringBuilder]::new()
    if ($Result.Disks.Count -eq 0) { [void]$dRows.Append("<tr><td colspan='7'>No physical disk counters read.</td></tr>") }
    foreach ($d in $Result.Disks) {
        $busyLevel = Get-TorporLevel -Value $d.BusyPercent -Warning $DiskBusyWarningPercent -ErrorAt $DiskBusyErrorPercent
        $latLevel  = Get-TorporLevel -Value $d.LatencyMs -Warning $DiskLatencyWarningMs -ErrorAt $DiskLatencyErrorMs
        $media = if ($d.MediaType -eq 'HDD') { "<span class='tk-badge-warn'>HDD</span>" } elseif ($d.MediaType -eq 'SSD') { "<span class='tk-badge-ok'>SSD</span>" } else { "<span class='tk-badge-info'>$(EscHtml $d.MediaType)</span>" }
        [void]$dRows.Append(
            "<tr><td class='tk-mono'>$(EscHtml $d.Instance)$(if ($d.IsSystem) { ' <span class=''tk-badge-blue''>System</span>' })</td>" +
            "<td>$(EscHtml $d.Model)</td><td>$media</td>" +
            "<td>$(Get-LevelBadge $busyLevel ("$($d.BusyPercent)%"))</td><td>$(Get-LevelBadge $latLevel ("$($d.LatencyMs) ms"))</td>" +
            "<td>$($d.Iops)</td><td>$(EscHtml (Format-Bytes $d.BytesPerSec))/s</td></tr>")
    }

    $svcRows = [System.Text.StringBuilder]::new()
    foreach ($r in $Result.Svchost) {
        [void]$svcRows.Append("<tr><td class='tk-mono'>$($r.Id)</td><td>$($r.CpuPercent)%</td><td class='tk-mono'>$(EscHtml $r.Services)</td></tr>")
    }
    $svcBlock = ''
    if ($Result.Svchost.Count -gt 0) {
        $svcBlock = "<div class='tk-card'><div class='tk-card-header'><span class='tk-card-label'>Busy service hosts (svchost)</span></div><table class='tk-table'><thead><tr><th>PID</th><th>CPU</th><th>Services inside</th></tr></thead><tbody>$($svcRows.ToString())</tbody></table></div>"
    }

    $sRows = [System.Text.StringBuilder]::new()
    if ($Result.Startup.Count -eq 0) { [void]$sRows.Append("<tr><td colspan='5'>No startup entries found.</td></tr>") }
    foreach ($it in ($Result.Startup | Sort-Object @{ Expression = { -not $_.Enabled } }, Name)) {
        $badge = if ($it.Enabled) { "<span class='tk-badge-ok'>Enabled</span>" } else { "<span class='tk-badge-info'>Disabled</span>" }
        [void]$sRows.Append("<tr><td>$(EscHtml $it.Name)</td><td>$badge</td><td>$(EscHtml $it.Scope)</td><td>$(EscHtml $it.Source)</td><td class='tk-mono'>$(EscHtml $it.Command)</td></tr>")
    }

    $bRows = [System.Text.StringBuilder]::new()
    if ($Result.Boot.Boots.Count -eq 0) { [void]$bRows.Append("<tr><td colspan='3'>No boot performance events in the last $BootLookbackDays days (the log may be disabled).</td></tr>") }
    foreach ($b in ($Result.Boot.Boots | Select-Object -First 10)) {
        $lvl = if ($b.TotalSeconds -gt $SlowBootWarningSeconds) { 'Warning' } else { 'Ok' }
        [void]$bRows.Append("<tr><td class='tk-mono'>$(EscHtml ($b.Time.ToString('yyyy-MM-dd HH:mm')))</td><td>$(Get-LevelBadge $lvl ("$($b.TotalSeconds) s"))</td><td>$($b.MainPathSeconds) s</td></tr>")
    }

    $blRows = [System.Text.StringBuilder]::new()
    if ($Result.Boot.Blamed.Count -eq 0) { [void]$blRows.Append("<tr><td colspan='4'>Windows blamed nothing for a slow boot in the last $BootLookbackDays days.</td></tr>") }
    foreach ($b in ($Result.Boot.Blamed | Select-Object -First 20)) {
        [void]$blRows.Append("<tr><td>$(EscHtml $b.Kind)</td><td>$(EscHtml $b.Name)</td><td>$($b.Count)</td><td>$($b.WorstSec) s</td></tr>")
    }

    $hRows = [System.Text.StringBuilder]::new()
    if ($Result.Hangs.Count -eq 0) { [void]$hRows.Append("<tr><td colspan='3'>No application hangs in the last $AppHangDays days.</td></tr>") }
    foreach ($h in $Result.Hangs) {
        $lvl = if ($h.Count -ge $AppHangThreshold) { 'Warning' } else { 'Ok' }
        [void]$hRows.Append("<tr><td>$(EscHtml $h.App)</td><td>$(Get-LevelBadge $lvl ("$($h.Count)"))</td><td class='tk-mono'>$(EscHtml ($h.Last.ToString('yyyy-MM-dd HH:mm')))</td></tr>")
    }

    $pfText = if ($Result.Pagefile.Count -gt 0) {
        ($Result.Pagefile | ForEach-Object { "$($_.Path): $($_.AllocatedMB) MB allocated, $($_.CurrentMB) MB in use, peak $($_.PeakMB) MB" }) -join '; '
    } elseif ($ctx.AutoPagefile) { 'System managed' } else { 'None' }

    $uptimeText = if ($ctx.Uptime) { '{0:N1} days' -f $ctx.Uptime.TotalDays } else { 'unknown' }
    $sysDriveText = if ($Result.SystemDrive) { "{0} free of {1} ({2}%)" -f (Format-Bytes $Result.SystemDrive.FreeBytes), (Format-Bytes $Result.SystemDrive.SizeBytes), $Result.SystemDrive.FreePercent } else { 'unknown' }
    $enabledStartup = @($Result.Startup | Where-Object { $_.Enabled }).Count

    $htmlHead = Get-TKHtmlHead `
        -Title      'T.O.R.P.O.R. Slow-Machine Triage' `
        -ScriptName 'T.O.R.P.O.R.' `
        -Subtitle   "${orgPrefix}Slow-Machine Triage -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Model'     = $ctx.Model
            'Generated' = $reportDate
            'Sample'    = ('{0:N0} s' -f $s.Elapsed)
            'Verdict'   = $Verdict.Verdict
        }) `
        -NavItems   @('Findings', 'Load', 'Top Processes', 'Disks', 'Startup & Boot', 'Hangs', 'Machine')

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Verdict</div></div>
    <div class="tk-summary-card $(switch ($cpuLevel) { 'Error' { 'err' } 'Warning' { 'warn' } default { 'ok' } })"><div class="tk-summary-num">$(& $fmt $s.CpuAvg '%')</div><div class="tk-summary-lbl">CPU Average</div></div>
    <div class="tk-summary-card $(switch ($availLevel) { 'Error' { 'err' } 'Warning' { 'warn' } default { 'ok' } })"><div class="tk-summary-num">$(& $fmt $Result.AvailPct '%')</div><div class="tk-summary-lbl">Lowest Free RAM</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(EscHtml (Format-Bytes $ctx.TotalRamBytes))</div><div class="tk-summary-lbl">Installed RAM</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(EscHtml $uptimeText)</div><div class="tk-summary-lbl">Uptime</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$enabledStartup</div><div class="tk-summary-lbl">Startup Programs</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Severity</th><th>Code</th><th>Finding</th><th>Remedy</th></tr></thead>
      <tbody>$($fRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Load</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Measure</th><th>Value</th><th>State</th><th>Threshold</th></tr></thead>
      <tbody>$loadRows</tbody></table></div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Top Processes</div>
    <div class="tk-info-box"><span class="tk-info-label">How to read this</span> Processes are grouped by name. CPU is the share of the whole machine over the sample, so the rows add up to the CPU average. Private memory is what the program has committed for itself.</div>
    <div class="tk-card"><div class="tk-card-header"><span class="tk-card-label">By CPU</span></div><table class="tk-table">
      <thead><tr><th>Program</th><th>CPU</th><th>Private memory</th><th>Working set</th><th>Disk I/O</th><th>Usually means</th></tr></thead>
      <tbody>$(Build-TorporProcessRow -Groups $Result.TopCpu -TotalRam $ctx.TotalRamBytes)</tbody></table></div>
    <div class="tk-card"><div class="tk-card-header"><span class="tk-card-label">By memory</span></div><table class="tk-table">
      <thead><tr><th>Program</th><th>CPU</th><th>Private memory</th><th>Working set</th><th>Disk I/O</th><th>Usually means</th></tr></thead>
      <tbody>$(Build-TorporProcessRow -Groups $Result.TopMem -TotalRam $ctx.TotalRamBytes)</tbody></table></div>
    <div class="tk-card"><div class="tk-card-header"><span class="tk-card-label">By disk I/O</span></div><table class="tk-table">
      <thead><tr><th>Program</th><th>CPU</th><th>Private memory</th><th>Working set</th><th>Disk I/O</th><th>Usually means</th></tr></thead>
      <tbody>$(Build-TorporProcessRow -Groups $Result.TopIo -TotalRam $ctx.TotalRamBytes)</tbody></table></div>
    $svcBlock
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> Disks</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Disk</th><th>Model</th><th>Media</th><th>Busy</th><th>Response time</th><th>IOPS</th><th>Throughput</th></tr></thead>
      <tbody>$($dRows.ToString())</tbody></table></div>
    <div class="tk-info-box"><span class="tk-info-label">System drive</span> $(EscHtml $sysDriveText)</div>
  </div>

  <div class="tk-section" id="s05">
    <div class="tk-section-title"><span class="tk-section-num">05</span> Startup &amp; Boot</div>
    <div class="tk-card"><div class="tk-card-header"><span class="tk-card-label">Recent boots</span></div><table class="tk-table">
      <thead><tr><th>Boot</th><th>To idle desktop</th><th>Main path</th></tr></thead>
      <tbody>$($bRows.ToString())</tbody></table></div>
    <div class="tk-card"><div class="tk-card-header"><span class="tk-card-label">What Windows blamed for slow boots</span></div><table class="tk-table">
      <thead><tr><th>Kind</th><th>Component</th><th>Times</th><th>Worst delay</th></tr></thead>
      <tbody>$($blRows.ToString())</tbody></table></div>
    <div class="tk-card"><div class="tk-card-header"><span class="tk-card-label">Startup programs</span></div><table class="tk-table">
      <thead><tr><th>Name</th><th>State</th><th>Scope</th><th>Source</th><th>Command</th></tr></thead>
      <tbody>$($sRows.ToString())</tbody></table></div>
    <div class="tk-info-box"><span class="tk-info-label">Scope</span> "Current user" entries are those of the account T.O.R.P.O.R. ran as. Startup entries are listed for their cost; T.A.L.O.N. audits them for persistence.</div>
  </div>

  <div class="tk-section" id="s06">
    <div class="tk-section-title"><span class="tk-section-num">06</span> Hangs</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Application</th><th>Hangs ($AppHangDays days)</th><th>Last</th></tr></thead>
      <tbody>$($hRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s07">
    <div class="tk-section-title"><span class="tk-section-num">07</span> Machine</div>
    <div class="tk-card"><div class="tk-info-box">
      <span class="tk-info-label">OS</span> $(EscHtml $ctx.OS)<br/>
      <span class="tk-info-label">CPU</span> $(EscHtml $ctx.Cpu) ($($ctx.Cores) cores / $($ctx.Logical) logical)<br/>
      <span class="tk-info-label">RAM</span> $(EscHtml (Format-Bytes $ctx.TotalRamBytes))<br/>
      <span class="tk-info-label">Page file</span> $(EscHtml $pfText)<br/>
      <span class="tk-info-label">Power plan</span> $(EscHtml $(if ($Result.Power.SchemeName) { $Result.Power.SchemeName } else { $Result.Power.SchemeGuid }))$(if ($Result.Power.PowerMode) { " / power mode $(EscHtml $Result.Power.PowerMode)" })$(if ($ctx.HasBattery) { $(if ($ctx.OnBattery) { ' -- on battery' } else { ' -- on AC power' }) })<br/>
      <span class="tk-info-label">Last boot</span> $(EscHtml "$($ctx.LastBoot)") ($(EscHtml $uptimeText))$(if ($ctx.FastStartup) { ' -- Fast Startup on: Shut down does not reset uptime' })
    </div></div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'T.O.R.P.O.R. v5.1')
    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# ORCHESTRATION
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-TorporRun {
    param([int]$Seconds)
    $Findings.Clear()

    $result = Invoke-TorporTriage -Seconds $Seconds

    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) { Write-Ok "Nothing slowing this machine showed up during the sample." }
    foreach ($f in $Findings) {
        $line = "$($f.Title)" + $(if ($f.Detail) { " -- $($f.Detail)" } else { '' })
        switch ($f.Severity) { 'Error' { Write-Fail $line } 'Warning' { Write-Warn $line } default { Write-Info $line } }
    }

    $verdict = Get-TorporVerdict -FindingList $Findings.ToArray()
    Write-Section "VERDICT"
    switch ($verdict.Class) { 'err' { Write-Fail $verdict.Verdict } 'warn' { Write-Warn $verdict.Verdict } default { Write-Ok $verdict.Verdict } }
    $top = @($result.TopCpu | Select-Object -First 1)
    Add-TKNote -Text ("TORPOR on {0}: {1} ({2} finding(s)); CPU avg {3}, top program {4}." -f $env:COMPUTERNAME, $verdict.Verdict, $Findings.Count,
        $(if ($null -ne $result.Sample.CpuAvg) { '{0:N0}%' -f $result.Sample.CpuAvg } else { 'n/a' }),
        $(if ($top.Count) { "$($top[0].Name) $($top[0].CpuPercent)%" } else { 'n/a' })) -Category 'Info' -ScriptName 'torpor'

    Write-Step "Generating HTML report..."
    $html    = Build-TorporReport -Result $result -Verdict $verdict
    $outPath = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) ("TORPOR_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'torpor' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-TorporBanner
    Invoke-TorporRun -Seconds $SampleSeconds
} else {
    $longSample = [math]::Max(120, $SampleSeconds)
    $choice = ''
    do {
        Show-TorporBanner
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] Triage  -  sample for $SampleSeconds seconds while the machine is slow" -ForegroundColor $C.Info
        Write-Host "  [2] Long triage  -  sample for $longSample seconds, for slowness that comes and goes" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' { Invoke-TorporRun -Seconds $SampleSeconds }
            '2' { Invoke-TorporRun -Seconds $longSample }
            'Q' { Write-Host ""; Write-Host "  Closing T.O.R.P.O.R." -ForegroundColor $C.Header; Write-Host "" }
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

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath -and -not (Test-Path (Join-Path $PSScriptRoot '.git'))) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
