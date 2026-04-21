<#
.SYNOPSIS
    S.E.N.T.I.N.E.L. — Scans & Evaluates services, Networks, Tasks, Infrastructure, Node Events & Logs
    Service, Task & Event Log Monitor for PowerShell 5.1+

.DESCRIPTION
    Audits critical Windows services, scheduled tasks, and recent event log errors on the local
    machine or optionally a remote machine via WinRM. Generates a dark-themed HTML health report,
    displays interactive console summaries, and can restart stopped critical services with
    per-service confirmation prompts.

.USAGE
    PS C:\> .\sentinel.ps1                              # Interactive menu (local machine)
    PS C:\> .\sentinel.ps1 -Unattended                  # Export health report silently
    PS C:\> .\sentinel.ps1 -Unattended -Target HOSTNAME  # Remote machine report

.NOTES
    Version : 1.0

    Tools Available
    ─────────────────────────────────────────────────────────────────
    G.R.I.M.O.I.R.E.       — Technician Toolkit hub and central launcher
    R.U.N.E.P.R.E.S.S.     — Printer driver installation & configuration
    R.E.S.T.O.R.A.T.I.O.N. — Windows Update management
    C.O.N.J.U.R.E.         — Software deployment via winget / Chocolatey
    O.R.A.C.L.E.           — System diagnostics & HTML report generation
    C.O.V.E.N.A.N.T.       — Machine onboarding & Entra ID domain join
    P.H.A.N.T.O.M.         — Profile migration & data transfer
    C.I.P.H.E.R.           — BitLocker drive encryption management
    W.A.R.D.               — User account & local security audit
    A.R.C.H.I.V.E.         — Pre-reimaging profile backup
    S.I.G.I.L.             — Security baseline & policy enforcement
    S.P.E.C.T.E.R.         — Remote machine execution via WinRM
    L.E.Y.L.I.N.E.         — Network diagnostics & remediation
    F.O.R.G.E.             — Driver update detection & installation
    A.E.G.I.S.             — Azure environment assessment & reporting
    B.A.S.T.I.O.N.         — Active Directory & identity management
    L.A.N.T.E.R.N.         — Network discovery & asset inventory
    T.H.R.E.S.H.O.L.D.     — Disk & storage health monitoring
    V.A.U.L.T.             — M365 license & mailbox auditing
    S.E.N.T.I.N.E.L.       — Service & scheduled task monitoring
    R.E.L.I.C.             — Certificate health & SSL expiry monitoring
    H.E.A.R.T.H.           — Toolkit setup & configuration wizard

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Healthy / running
    Yellow   Warnings / degraded
    Red      Critical errors / stopped
    Gray     Information and details
#>

param(
    [switch]$Unattended,
    [string]$Target = '',
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

Import-Module "$PSScriptRoot\TechnicianToolkit.psm1" -Force
Assert-AdminPrivilege

# ─────────────────────────────────────────────────────────────────────────────
# SCRIPT PATH RESOLUTION
# ─────────────────────────────────────────────────────────────────────────────

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

$C = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
    Accent   = 'Blue'
}

# ─────────────────────────────────────────────────────────────────────────────
# REMOTE TARGET STATE
# ─────────────────────────────────────────────────────────────────────────────

$script:RemoteTarget = $Target.Trim()

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-SentinelBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ███████╗███████╗███╗   ██╗████████╗██╗███╗   ██╗███████╗██╗
  ██╔════╝██╔════╝████╗  ██║╚══██╔══╝██║████╗  ██║██╔════╝██║
  ███████╗█████╗  ██╔██╗ ██║   ██║   ██║██╔██╗ ██║█████╗  ██║
  ╚════██║██╔══╝  ██║╚██╗██║   ██║   ██║██║╚██╗██║██╔══╝  ██║
  ███████║███████╗██║ ╚████║   ██║   ██║██║ ╚████║███████╗███████╗
  ╚══════╝╚══════╝╚═╝  ╚═══╝   ╚═╝   ╚═╝╚═╝  ╚═══╝╚══════╝╚══════╝

"@ -ForegroundColor Cyan
    Write-Host "    S.E.N.T.I.N.E.L. — Scans & Evaluates services, Networks, Tasks, Infrastructure, Node Events & Logs" -ForegroundColor Cyan
    Write-Host "    Service, Task & Event Log Monitor" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function HtmlEncode {
    param([string]$s)
    $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

function Show-Spinner {
    param([string]$Message, [scriptblock]$Work)
    $frames = @('|','/','-','\')
    $i = 0
    $job = Start-Job -ScriptBlock $Work
    Write-Host -NoNewline "  $Message " -ForegroundColor $C.Progress
    while ($job.State -eq 'Running') {
        Write-Host -NoNewline "`b$($frames[$i % 4])" -ForegroundColor $C.Progress
        $i++
        Start-Sleep -Milliseconds 120
    }
    Write-Host -NoNewline "`b" -ForegroundColor $C.Progress
    $result = Receive-Job $job
    Remove-Job $job
    return $result
}

# ─────────────────────────────────────────────────────────────────────────────
# CRITICAL SERVICES LIST
# ─────────────────────────────────────────────────────────────────────────────

$CriticalServices = @(
    @{ Name='wuauserv';          Display='Windows Update'                   },
    @{ Name='WinDefend';         Display='Windows Defender Antivirus'       },
    @{ Name='EventLog';          Display='Windows Event Log'                },
    @{ Name='Schedule';          Display='Task Scheduler'                   },
    @{ Name='Dnscache';          Display='DNS Client'                       },
    @{ Name='LanmanWorkstation'; Display='Workstation (SMB Client)'         },
    @{ Name='W32Time';           Display='Windows Time'                     },
    @{ Name='SamSs';             Display='Security Accounts Manager'        },
    @{ Name='RpcSs';             Display='Remote Procedure Call'            },
    @{ Name='BITS';              Display='Background Intelligent Transfer'  },
    @{ Name='cryptsvc';          Display='Cryptographic Services'           },
    @{ Name='MpsSvc';            Display='Windows Firewall'                 },
    @{ Name='Spooler';           Display='Print Spooler'                    },
    @{ Name='lmhosts';           Display='TCP/IP NetBIOS Helper'            },
    @{ Name='WerSvc';            Display='Windows Error Reporting'          }
)

# ─────────────────────────────────────────────────────────────────────────────
# GET-SERVICEHEALTH
# ─────────────────────────────────────────────────────────────────────────────

function Get-ServiceHealth {
    $serviceNames = $CriticalServices | ForEach-Object { $_.Name }
    $displayMap   = @{}
    foreach ($svc in $CriticalServices) { $displayMap[$svc.Name] = $svc.Display }

    $scriptBlock = {
        param($names)
        $results = @()
        foreach ($name in $names) {
            try {
                $svc = Get-Service -Name $name -ErrorAction Stop
                $results += [PSCustomObject]@{
                    Name      = $svc.Name
                    Status    = $svc.Status.ToString()
                    StartType = $svc.StartType.ToString()
                }
            }
            catch {
                $results += [PSCustomObject]@{
                    Name      = $name
                    Status    = 'NotFound'
                    StartType = 'Unknown'
                }
            }
        }
        return $results
    }

    if ($script:RemoteTarget) {
        try {
            $raw = Invoke-Command -ComputerName $script:RemoteTarget -ScriptBlock $scriptBlock -ArgumentList (,$serviceNames) -ErrorAction Stop
        }
        catch {
            Write-Host "  [-] Remote service query failed: $_" -ForegroundColor $C.Error
            $raw = @()
        }
    } else {
        $raw = & $scriptBlock -names $serviceNames
    }

    $output = @()
    foreach ($r in $raw) {
        $concern = ($r.Status -eq 'Stopped' -and $r.StartType -eq 'Automatic')
        $output += [PSCustomObject]@{
            Name        = $r.Name
            DisplayName = if ($displayMap.ContainsKey($r.Name)) { $displayMap[$r.Name] } else { $r.Name }
            Status      = $r.Status
            StartType   = $r.StartType
            Concern     = $concern
        }
    }
    return $output
}

# ─────────────────────────────────────────────────────────────────────────────
# GET-TASKAUDIT
# ─────────────────────────────────────────────────────────────────────────────

function Get-TaskAudit {
    $scriptBlock = {
        $tasks = Get-ScheduledTask -ErrorAction SilentlyContinue
        $results = @()
        foreach ($task in $tasks) {
            try {
                $info = $task | Get-ScheduledTaskInfo -ErrorAction Stop
                $results += [PSCustomObject]@{
                    TaskName       = $task.TaskName
                    TaskPath       = $task.TaskPath
                    State          = $task.State.ToString()
                    LastRunTime    = $info.LastRunTime
                    LastTaskResult = $info.LastTaskResult
                    NextRunTime    = $info.NextRunTime
                }
            }
            catch {
                $results += [PSCustomObject]@{
                    TaskName       = $task.TaskName
                    TaskPath       = $task.TaskPath
                    State          = $task.State.ToString()
                    LastRunTime    = $null
                    LastTaskResult = -1
                    NextRunTime    = $null
                }
            }
        }
        return $results
    }

    if ($script:RemoteTarget) {
        try {
            $raw = Invoke-Command -ComputerName $script:RemoteTarget -ScriptBlock $scriptBlock -ErrorAction Stop
        }
        catch {
            Write-Host "  [-] Remote task query failed: $_" -ForegroundColor $C.Error
            return @()
        }
    } else {
        $raw = & $scriptBlock
    }

    $cutoff = (Get-Date).AddDays(-7)
    # LastTaskResult codes: 0=success, 267009=running, 267011=never run
    $successCodes = @(0, 267009, 267011)

    $output = @()
    foreach ($t in $raw) {
        $isMicrosoft  = $t.TaskPath -like '\Microsoft\*'
        $lastRunStale = ($t.LastRunTime -ne $null -and $t.LastRunTime -lt $cutoff -and $t.LastRunTime -gt [datetime]'1900-01-01')
        $hasFailed    = ($t.LastTaskResult -notin $successCodes)
        $isDisabled   = ($t.State -eq 'Disabled')

        # Flag logic
        $flagReason = ''
        if ($hasFailed -and $lastRunStale -and -not $isMicrosoft) {
            $flagReason = 'Failed+Stale'
        } elseif ($hasFailed -and -not $isMicrosoft) {
            $flagReason = 'Failed'
        } elseif ($isDisabled -and -not $isMicrosoft) {
            $flagReason = 'Disabled'
        } elseif ($hasFailed -and $isMicrosoft) {
            $flagReason = 'MSFailed'
        }

        $output += [PSCustomObject]@{
            TaskName       = $t.TaskName
            TaskPath       = $t.TaskPath
            State          = $t.State
            LastRunTime    = $t.LastRunTime
            LastTaskResult = $t.LastTaskResult
            NextRunTime    = $t.NextRunTime
            IsMicrosoft    = $isMicrosoft
            FlagReason     = $flagReason
        }
    }
    return $output
}

# ─────────────────────────────────────────────────────────────────────────────
# GET-EVENTERRORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-EventErrors {
    $since = (Get-Date).AddHours(-24)

    if ($script:RemoteTarget) {
        $target = $script:RemoteTarget
        $sysErrors = try {
            Get-EventLog -LogName System -EntryType Error -After $since -Newest 50 -ComputerName $target -ErrorAction Stop
        } catch { @() }

        $appErrors = try {
            Get-EventLog -LogName Application -EntryType Error -After $since -Newest 50 -ComputerName $target -ErrorAction Stop
        } catch { @() }
    } else {
        $sysErrors = try {
            Get-EventLog -LogName System -EntryType Error -After $since -Newest 50 -ErrorAction Stop
        } catch { @() }

        $appErrors = try {
            Get-EventLog -LogName Application -EntryType Error -After $since -Newest 50 -ErrorAction Stop
        } catch { @() }
    }

    $allErrors = @()
    foreach ($e in $sysErrors) {
        $allErrors += [PSCustomObject]@{
            Log           = 'System'
            TimeGenerated = $e.TimeGenerated
            EntryType     = $e.EntryType.ToString()
            Source        = $e.Source
            EventID       = $e.EventID
            Message       = if ($e.Message.Length -gt 120) { $e.Message.Substring(0,120) + '...' } else { $e.Message }
        }
    }
    foreach ($e in $appErrors) {
        $allErrors += [PSCustomObject]@{
            Log           = 'Application'
            TimeGenerated = $e.TimeGenerated
            EntryType     = $e.EntryType.ToString()
            Source        = $e.Source
            EventID       = $e.EventID
            Message       = if ($e.Message.Length -gt 120) { $e.Message.Substring(0,120) + '...' } else { $e.Message }
        }
    }

    # Build source summary (top 10)
    $sourceSummary = $allErrors |
        Group-Object Source |
        Sort-Object Count -Descending |
        Select-Object -First 10 |
        ForEach-Object { [PSCustomObject]@{ Source = $_.Name; Count = $_.Count } }

    return [PSCustomObject]@{
        Events        = ($allErrors | Sort-Object TimeGenerated -Descending)
        SourceSummary = $sourceSummary
        TotalCount    = $allErrors.Count
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# INVOKE-RESTARTSERVICE
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-RestartService {
    param([array]$ServiceData)

    $stopped = $ServiceData | Where-Object { $_.Status -eq 'Stopped' -and $_.StartType -eq 'Automatic' }

    if (-not $stopped -or $stopped.Count -eq 0) {
        Write-Host "  [+] No stopped Automatic-start services to restart." -ForegroundColor $C.Success
        return
    }

    Write-Host "  Found $($stopped.Count) stopped Automatic-start service(s):" -ForegroundColor $C.Warning
    Write-Host ""

    $restartAll = $false
    $skipAll    = $false

    foreach ($svc in $stopped) {
        if ($skipAll) { break }

        $label = "$($svc.DisplayName) ($($svc.Name))"

        if ($restartAll) {
            Write-Host "  [*] Restarting: $label" -ForegroundColor $C.Progress
        } else {
            Write-Host -NoNewline "  Restart $label ? (Y/N/A=All/S=Skip all): " -ForegroundColor $C.Warning
            $answer = (Read-Host).Trim().ToUpper()
            switch ($answer) {
                'A' { $restartAll = $true }
                'S' { $skipAll    = $true; Write-Host "  [*] Skipping remaining." -ForegroundColor $C.Info; break }
                'N' { Write-Host "  [*] Skipped." -ForegroundColor $C.Info; continue }
            }
        }

        if (-not $skipAll) {
            try {
                if ($script:RemoteTarget) {
                    Invoke-Command -ComputerName $script:RemoteTarget -ScriptBlock {
                        param($n) Start-Service -Name $n -ErrorAction Stop
                    } -ArgumentList $svc.Name -ErrorAction Stop
                } else {
                    Start-Service -Name $svc.Name -ErrorAction Stop
                }
                Write-Host "  [+] Started: $label" -ForegroundColor $C.Success
            }
            catch {
                Write-Host "  [-] Failed to start $label`: $_" -ForegroundColor $C.Error
            }
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SHOW-SERVICEHEALTH (console)
# ─────────────────────────────────────────────────────────────────────────────

function Show-ServiceHealth {
    param([array]$Data)

    Write-Section "SERVICE HEALTH — CRITICAL SERVICES"

    if (-not $Data -or $Data.Count -eq 0) {
        Write-Host "  No service data available." -ForegroundColor $C.Warning
        return
    }

    $colW = @(26, 34, 10, 12, 8)
    $hdr  = "  {0,-$($colW[0])}{1,-$($colW[1])}{2,-$($colW[2])}{3,-$($colW[3])}{4}" -f 'Service Name','Display Name','Status','Start Type','Concern'
    Write-Host $hdr -ForegroundColor $C.Header

    foreach ($svc in $Data) {
        $statusColor = switch ($svc.Status) {
            'Running' { $C.Success }
            'Stopped' { $C.Error   }
            default   { $C.Warning }
        }
        $concernStr  = if ($svc.Concern) { '!! CONCERN' } else { '' }
        $concernColor = if ($svc.Concern) { $C.Error } else { $C.Info }

        $line = "  {0,-$($colW[0])}{1,-$($colW[1])}{2,-$($colW[2])}{3,-$($colW[3])}" -f `
            $svc.Name, $svc.DisplayName, $svc.Status, $svc.StartType
        Write-Host -NoNewline $line
        Write-Host $concernStr -ForegroundColor $concernColor
    }

    $ok      = ($Data | Where-Object { -not $_.Concern }).Count
    $concern = ($Data | Where-Object {    $_.Concern }).Count
    Write-Host ""
    Write-Host "  Services OK: $ok   Concerns: $concern" -ForegroundColor $(if ($concern -gt 0) { $C.Warning } else { $C.Success })
}

# ─────────────────────────────────────────────────────────────────────────────
# SHOW-TASKAUDIT (console)
# ─────────────────────────────────────────────────────────────────────────────

function Show-TaskAudit {
    param([array]$Data)

    Write-Section "SCHEDULED TASK AUDIT"

    if (-not $Data -or $Data.Count -eq 0) {
        Write-Host "  No task data available." -ForegroundColor $C.Warning
        return
    }

    $flagged = $Data | Where-Object { $_.FlagReason -and $_.FlagReason -ne 'MSFailed' } | Sort-Object FlagReason
    $msFailed = $Data | Where-Object { $_.FlagReason -eq 'MSFailed' }

    if ($flagged.Count -eq 0) {
        Write-Host "  [+] No problematic non-Microsoft tasks found." -ForegroundColor $C.Success
    } else {
        Write-Host "  Non-Microsoft flagged tasks ($($flagged.Count)):" -ForegroundColor $C.Warning
        Write-Host ""
        $hdr = "  {0,-32}{1,-28}{2,-10}{3,-12}{4}" -f 'Task Name','Path','State','Last Result','Flag'
        Write-Host $hdr -ForegroundColor $C.Header

        foreach ($t in $flagged) {
            $color = switch -Wildcard ($t.FlagReason) {
                'Failed*' { $C.Error   }
                'Disabled' { $C.Warning }
                default    { $C.Info   }
            }
            $shortPath = if ($t.TaskPath.Length -gt 27) { $t.TaskPath.Substring(0,24) + '...' } else { $t.TaskPath }
            $shortName = if ($t.TaskName.Length -gt 31) { $t.TaskName.Substring(0,28) + '...' } else { $t.TaskName }
            $line = "  {0,-32}{1,-28}{2,-10}{3,-12}{4}" -f $shortName, $shortPath, $t.State, $t.LastTaskResult, $t.FlagReason
            Write-Host $line -ForegroundColor $color
        }
    }

    if ($msFailed.Count -gt 0) {
        Write-Host ""
        Write-Host "  Microsoft tasks with errors ($($msFailed.Count)):" -ForegroundColor $C.Warning
        foreach ($t in $msFailed | Select-Object -First 10) {
            Write-Host "    $($t.TaskPath)$($t.TaskName)  [Result: $($t.LastTaskResult)]" -ForegroundColor $C.Info
        }
    }

    Write-Host ""
    $disabledNonMs = ($Data | Where-Object { $_.FlagReason -eq 'Disabled' }).Count
    $failedNonMs   = ($Data | Where-Object { $_.FlagReason -like 'Failed*' }).Count
    Write-Host "  Failed (non-MS): $failedNonMs   Disabled (non-MS): $disabledNonMs   MS errors: $($msFailed.Count)" -ForegroundColor $C.Info
}

# ─────────────────────────────────────────────────────────────────────────────
# SHOW-EVENTERRORS (console)
# ─────────────────────────────────────────────────────────────────────────────

function Show-EventErrors {
    param([object]$Data)

    Write-Section "RECENT EVENT LOG ERRORS (Last 24 Hours)"

    if (-not $Data -or $Data.TotalCount -eq 0) {
        Write-Host "  [+] No errors found in System or Application logs." -ForegroundColor $C.Success
        return
    }

    Write-Host "  Total errors: $($Data.TotalCount)" -ForegroundColor $C.Warning
    Write-Host ""

    # Source summary
    Write-Host "  Top Error Sources:" -ForegroundColor $C.Header
    $hdr = "  {0,-40}{1}" -f 'Source','Count'
    Write-Host $hdr -ForegroundColor $C.Header
    foreach ($src in $Data.SourceSummary) {
        $color = if ($src.Count -ge 10) { $C.Error } elseif ($src.Count -ge 3) { $C.Warning } else { $C.Info }
        Write-Host ("  {0,-40}{1}" -f $src.Source, $src.Count) -ForegroundColor $color
    }

    Write-Host ""
    Write-Host "  Recent Events (newest first):" -ForegroundColor $C.Header
    $hdr2 = "  {0,-22}{1,-6}{2,-26}{3,-8}{4}" -f 'Time','Log','Source','EventID','Message'
    Write-Host $hdr2 -ForegroundColor $C.Header

    foreach ($evt in ($Data.Events | Select-Object -First 20)) {
        $timeStr   = $evt.TimeGenerated.ToString('MM/dd HH:mm:ss')
        $shortSrc  = if ($evt.Source.Length -gt 25) { $evt.Source.Substring(0,22) + '...' } else { $evt.Source }
        $shortMsg  = if ($evt.Message.Length -gt 50) { $evt.Message.Substring(0,47) + '...' } else { $evt.Message }
        $logAbbr   = if ($evt.Log -eq 'System') { 'SYS' } else { 'APP' }
        $line = "  {0,-22}{1,-6}{2,-26}{3,-8}{4}" -f $timeStr, $logAbbr, $shortSrc, $evt.EventID, $shortMsg
        Write-Host $line -ForegroundColor $C.Error
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# BUILD-HTMLREPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param(
        [array]$Services,
        [array]$Tasks,
        [object]$Events,
        [string]$MachineName,
        [string]$ReportTimestamp
    )

    # Summary counts
    $svcOk       = ($Services | Where-Object { -not $_.Concern }).Count
    $svcConcern  = ($Services | Where-Object {    $_.Concern }).Count
    $taskFailed  = ($Tasks    | Where-Object { $_.FlagReason -like 'Failed*' }).Count
    $taskDisabled = ($Tasks   | Where-Object { $_.FlagReason -eq 'Disabled' }).Count
    $eventTotal  = if ($Events) { $Events.TotalCount } else { 0 }

    # Service rows
    $svcRows = ''
    foreach ($svc in $Services) {
        $statusBadge = if ($svc.Status -eq 'Running') {
            "<span class='tk-badge-ok'>Running</span>"
        } elseif ($svc.Status -eq 'Stopped') {
            "<span class='tk-badge-err'>Stopped</span>"
        } else {
            "<span class='tk-badge-warn'>$(HtmlEncode($svc.Status))</span>"
        }
        $concernCell = if ($svc.Concern) {
            "<span class='tk-badge-err'>!! Concern</span>"
        } else {
            "<span style='color:#555'>-</span>"
        }
        $svcRows += "<tr>
            <td>$(HtmlEncode($svc.Name))</td>
            <td>$(HtmlEncode($svc.DisplayName))</td>
            <td>$statusBadge</td>
            <td>$(HtmlEncode($svc.StartType))</td>
            <td>$concernCell</td>
        </tr>`n"
    }

    # Task rows — non-Microsoft flagged
    $taskRowsFlagged = ''
    $flaggedTasks = $Tasks | Where-Object { $_.FlagReason -and $_.FlagReason -ne 'MSFailed' } | Sort-Object FlagReason
    foreach ($t in $flaggedTasks) {
        $resultBadge = if ($t.LastTaskResult -in @(0, 267009, 267011)) {
            "<span class='tk-badge-ok'>$($t.LastTaskResult)</span>"
        } else {
            "<span class='tk-badge-err'>$($t.LastTaskResult)</span>"
        }
        $stateBadge = if ($t.State -eq 'Ready') {
            "<span class='tk-badge-ok'>Ready</span>"
        } elseif ($t.State -eq 'Disabled') {
            "<span class='tk-badge-warn'>Disabled</span>"
        } elseif ($t.State -eq 'Running') {
            "<span class='tk-badge-ok'>Running</span>"
        } else {
            "<span class='tk-badge-info'>$(HtmlEncode($t.State))</span>"
        }
        $lastRun  = if ($t.LastRunTime -and $t.LastRunTime -gt [datetime]'1900-01-01') { $t.LastRunTime.ToString('yyyy-MM-dd HH:mm') } else { 'Never' }
        $nextRun  = if ($t.NextRunTime -and $t.NextRunTime -gt [datetime]'1900-01-01') { $t.NextRunTime.ToString('yyyy-MM-dd HH:mm') } else { 'N/A' }
        $taskRowsFlagged += "<tr>
            <td>$(HtmlEncode($t.TaskName))</td>
            <td class='tk-mono' style='font-size:11px'>$(HtmlEncode($t.TaskPath))</td>
            <td>$stateBadge</td>
            <td>$lastRun</td>
            <td>$resultBadge</td>
            <td>$nextRun</td>
        </tr>`n"
    }

    # Task rows — MS failed
    $taskRowsMs = ''
    $msFailed = $Tasks | Where-Object { $_.FlagReason -eq 'MSFailed' }
    foreach ($t in $msFailed) {
        $lastRun = if ($t.LastRunTime -and $t.LastRunTime -gt [datetime]'1900-01-01') { $t.LastRunTime.ToString('yyyy-MM-dd HH:mm') } else { 'Never' }
        $taskRowsMs += "<tr>
            <td>$(HtmlEncode($t.TaskName))</td>
            <td class='tk-mono' style='font-size:11px'>$(HtmlEncode($t.TaskPath))</td>
            <td><span class='tk-badge-info'>$(HtmlEncode($t.State))</span></td>
            <td>$lastRun</td>
            <td><span class='tk-badge-warn'>$($t.LastTaskResult)</span></td>
            <td>N/A</td>
        </tr>`n"
    }

    # Event source summary rows
    $srcRows = ''
    if ($Events -and $Events.SourceSummary) {
        foreach ($src in $Events.SourceSummary) {
            $cntBadge = if ($src.Count -ge 10) { 'tk-badge-err' } elseif ($src.Count -ge 3) { 'tk-badge-warn' } else { 'tk-badge-info' }
            $srcRows += "<tr><td>$(HtmlEncode($src.Source))</td><td><span class='$cntBadge'>$($src.Count)</span></td></tr>`n"
        }
    }

    # Event detail rows
    $evtRows = ''
    if ($Events -and $Events.Events) {
        foreach ($evt in ($Events.Events | Select-Object -First 50)) {
            $timeStr  = $evt.TimeGenerated.ToString('yyyy-MM-dd HH:mm:ss')
            $logBadge = if ($evt.Log -eq 'System') {
                "<span class='tk-badge-warn'>SYS</span>"
            } else {
                "<span class='tk-badge-info'>APP</span>"
            }
            $evtRows += "<tr>
                <td class='tk-mono'>$timeStr</td>
                <td>$logBadge</td>
                <td>$(HtmlEncode($evt.Source))</td>
                <td class='tk-mono'>$($evt.EventID)</td>
                <td style='font-size:12px'>$(HtmlEncode($evt.Message))</td>
            </tr>`n"
        }
    }

    $noFlagged  = if ($flaggedTasks.Count -eq 0) { "<div class='tk-info-box'><span class='tk-info-label'>INFO</span> No non-Microsoft tasks flagged.</div>" } else { '' }
    $noMsFailed = if ($msFailed.Count -eq 0)     { "<div class='tk-info-box'><span class='tk-info-label'>INFO</span> No Microsoft task failures.</div>"      } else { '' }
    $noEvents   = if ($eventTotal -eq 0)          { "<div class='tk-info-box'><span class='tk-info-label tk-badge-ok'>OK</span> No errors in System or Application logs in the last 24 hours.</div>" } else { '' }

    $svcOkClass      = if ($svcConcern -gt 0) { 'warn' } else { 'ok' }
    $svcConcernClass = if ($svcConcern -eq 0) { 'ok' }   else { 'err' }
    $taskFailClass   = if ($taskFailed -eq 0)  { 'ok' }   else { 'err' }
    $taskDisClass    = if ($taskDisabled -eq 0) { 'ok' }  else { 'warn' }
    $evtClass        = if ($eventTotal -eq 0)  { 'ok' }   elseif ($eventTotal -lt 10) { 'warn' } else { 'err' }

    $showFlagTable = if ($flaggedTasks.Count -eq 0) { ' style="display:none"' } else { '' }
    $showMsTable   = if ($msFailed.Count -eq 0)     { ' style="display:none"' } else { '' }

    $html  = Get-TKHtmlHead `
        -Title      'Service & Task Health Report' `
        -ScriptName 'S.E.N.T.I.N.E.L.' `
        -Subtitle   $MachineName `
        -MetaItems  ([ordered]@{
            'Machine'   = $MachineName
            'Generated' = $ReportTimestamp
            'Tool'      = 'Service, Task & Event Log Monitor v1.0'
        }) `
        -NavItems   @('Critical Services', 'Scheduled Tasks', 'Event Log Errors')

    $html += @"

<div class="tk-summary-row">
    <div class="tk-summary-card $svcOkClass">
        <div class="tk-summary-num">$svcOk</div>
        <div class="tk-summary-lbl">Services OK</div>
    </div>
    <div class="tk-summary-card $svcConcernClass">
        <div class="tk-summary-num">$svcConcern</div>
        <div class="tk-summary-lbl">Services Concern</div>
    </div>
    <div class="tk-summary-card $taskFailClass">
        <div class="tk-summary-num">$taskFailed</div>
        <div class="tk-summary-lbl">Tasks Failed</div>
    </div>
    <div class="tk-summary-card $taskDisClass">
        <div class="tk-summary-num">$taskDisabled</div>
        <div class="tk-summary-lbl">Tasks Disabled</div>
    </div>
    <div class="tk-summary-card $evtClass">
        <div class="tk-summary-num">$eventTotal</div>
        <div class="tk-summary-lbl">Event Errors (24h)</div>
    </div>
</div>

<div class="tk-section" id="critical-services">
    <div class="tk-card">
        <div class="tk-card-header">
            <span class="tk-card-label">Critical Services</span>
        </div>
        <table class="tk-table">
            <thead><tr>
                <th>Service Name</th>
                <th>Display Name</th>
                <th>Status</th>
                <th>Start Type</th>
                <th>Concern</th>
            </tr></thead>
            <tbody>
                $svcRows
            </tbody>
        </table>
    </div>
</div>

<div class="tk-section" id="scheduled-tasks">
    <div class="tk-card">
        <div class="tk-card-header">
            <span class="tk-card-label">Scheduled Tasks</span>
        </div>
        <div class="tk-section-subtitle" style="padding:0 0 10px 0">Non-Microsoft Flagged Tasks</div>
        $noFlagged
        <table class="tk-table"$showFlagTable>
            <thead><tr>
                <th>Task Name</th>
                <th>Path</th>
                <th>State</th>
                <th>Last Run</th>
                <th>Last Result</th>
                <th>Next Run</th>
            </tr></thead>
            <tbody>
                $taskRowsFlagged
            </tbody>
        </table>
        <div class="tk-divider"></div>
        <div class="tk-section-subtitle" style="padding:10px 0">Microsoft Tasks with Errors</div>
        $noMsFailed
        <table class="tk-table"$showMsTable>
            <thead><tr>
                <th>Task Name</th>
                <th>Path</th>
                <th>State</th>
                <th>Last Run</th>
                <th>Last Result</th>
                <th>Next Run</th>
            </tr></thead>
            <tbody>
                $taskRowsMs
            </tbody>
        </table>
    </div>
</div>

<div class="tk-section" id="event-log-errors">
    <div class="tk-card">
        <div class="tk-card-header">
            <span class="tk-card-label">Event Log Errors (Last 24 Hours)</span>
        </div>
        $noEvents
        <div class="tk-section-subtitle" style="padding:0 0 10px 0">Top Error Sources</div>
        <table class="tk-table" style="max-width:480px">
            <thead><tr><th>Source</th><th>Error Count</th></tr></thead>
            <tbody>$srcRows</tbody>
        </table>
        <div class="tk-divider"></div>
        <div class="tk-section-subtitle" style="padding:10px 0">Recent Error Events (newest first, up to 50)</div>
        <table class="tk-table">
            <thead><tr>
                <th>Time</th>
                <th>Log</th>
                <th>Source</th>
                <th>Event ID</th>
                <th>Message</th>
            </tr></thead>
            <tbody>
                $evtRows
            </tbody>
        </table>
    </div>
</div>

"@

    $html += Get-TKHtmlFoot -ScriptName 'S.E.N.T.I.N.E.L. v1.0'
    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# REMOTE TARGET CONNECT
# ─────────────────────────────────────────────────────────────────────────────

function Connect-RemoteTarget {
    Write-Host ""
    Write-Host -NoNewline "  Enter hostname or IP (leave blank to revert to LOCAL): " -ForegroundColor $C.Header
    $newTarget = (Read-Host).Trim()

    if ([string]::IsNullOrWhiteSpace($newTarget)) {
        $script:RemoteTarget = ''
        Write-Host "  [*] Reverted to LOCAL machine." -ForegroundColor $C.Success
        return
    }

    Write-Host "  [*] Testing WinRM connectivity to $newTarget..." -ForegroundColor $C.Progress
    try {
        Test-WSMan -ComputerName $newTarget -ErrorAction Stop | Out-Null
        $script:RemoteTarget = $newTarget
        Write-Host "  [+] Connected. Target is now: $newTarget" -ForegroundColor $C.Success
    }
    catch {
        Write-Host "  [-] WinRM connection failed: $_" -ForegroundColor $C.Error
        Write-Host "      Ensure WinRM is enabled on target: Enable-PSRemoting -Force" -ForegroundColor $C.Warning
        Write-Host "  [*] Keeping current target: $(if ($script:RemoteTarget) { $script:RemoteTarget } else { 'LOCAL' })" -ForegroundColor $C.Info
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# EXPORT HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Export-HtmlReport {
    param(
        [array]$Services,
        [array]$Tasks,
        [object]$Events
    )

    $machineName = if ($script:RemoteTarget) { $script:RemoteTarget } else { $env:COMPUTERNAME }
    $timestamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reportName  = "SENTINEL_${timestamp}.html"
    $reportPath  = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) $reportName

    Write-Host "  [*] Building HTML report..." -ForegroundColor $C.Progress

    $html = Build-HtmlReport `
        -Services  $Services `
        -Tasks     $Tasks `
        -Events    $Events `
        -MachineName      $machineName `
        -ReportTimestamp  (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

    try {
        [System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
        Write-Host "  [+] Report saved: $reportPath" -ForegroundColor $C.Success
    }
    catch {
        Write-Host "  [-] Failed to save report: $_" -ForegroundColor $C.Error
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# UNATTENDED MODE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-SentinelBanner
    $machineName = if ($script:RemoteTarget) { $script:RemoteTarget } else { $env:COMPUTERNAME }
    Write-Host "  [*] Unattended mode  -  Target: $machineName" -ForegroundColor $C.Progress
    Write-Host ""

    Write-Host "  [*] Collecting service health..." -ForegroundColor $C.Progress
    $services = Get-ServiceHealth

    Write-Host "  [*] Collecting scheduled task audit..." -ForegroundColor $C.Progress
    $tasks = Get-TaskAudit

    Write-Host "  [*] Collecting event log errors (this may take a moment)..." -ForegroundColor $C.Progress
    $events = Get-EventErrors

    Export-HtmlReport -Services $services -Tasks $tasks -Events $events
    Write-Host ""
    Write-Host "  [+] S.E.N.T.I.N.E.L. unattended run complete." -ForegroundColor $C.Success
    Write-Host ""
    if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
    exit 0
}

# ─────────────────────────────────────────────────────────────────────────────
# INTERACTIVE MENU
# ─────────────────────────────────────────────────────────────────────────────

# Pre-load data
$cachedServices = $null
$cachedTasks    = $null
$cachedEvents   = $null

function Invoke-RefreshAll {
    $machineName = if ($script:RemoteTarget) { $script:RemoteTarget } else { $env:COMPUTERNAME }
    Write-Host "  [*] Refreshing all data from: $machineName" -ForegroundColor $C.Progress

    Write-Host "  [*] Querying services..." -ForegroundColor $C.Progress
    $script:cachedServices = Get-ServiceHealth

    Write-Host "  [*] Querying scheduled tasks..." -ForegroundColor $C.Progress
    $script:cachedTasks = Get-TaskAudit

    Write-Host "  [*] Querying event logs (last 24h, up to 50 each log)..." -ForegroundColor $C.Progress
    $script:cachedEvents = Get-EventErrors

    Write-Host "  [+] Refresh complete." -ForegroundColor $C.Success
}

# Initial data load
Show-SentinelBanner
Invoke-RefreshAll

$choice = ''

do {
    Show-SentinelBanner

    $targetLabel = if ($script:RemoteTarget) { $script:RemoteTarget } else { 'LOCAL' }
    $svcConcern  = if ($cachedServices) { ($cachedServices | Where-Object { $_.Concern }).Count } else { '?' }
    $taskFailed  = if ($cachedTasks)    { ($cachedTasks | Where-Object { $_.FlagReason -like 'Failed*' }).Count } else { '?' }
    $evtCount    = if ($cachedEvents)   { $cachedEvents.TotalCount } else { '?' }

    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host "  S.E.N.T.I.N.E.L. MENU   -   Target: $targetLabel" -ForegroundColor $C.Header
    Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
    Write-Host ""

    $svcColor = if ($svcConcern -gt 0) { $C.Error } else { $C.Success }
    $tskColor = if ($taskFailed -gt 0) { $C.Error } else { $C.Success }
    $evtColor = if ($evtCount -gt 0)   { $C.Warning } else { $C.Success }

    Write-Host -NoNewline "  [1] " -ForegroundColor $C.Header
    Write-Host -NoNewline "Show service health          " -ForegroundColor $C.Info
    Write-Host "Concerns: $svcConcern" -ForegroundColor $svcColor

    Write-Host -NoNewline "  [2] " -ForegroundColor $C.Header
    Write-Host -NoNewline "Show scheduled task audit    " -ForegroundColor $C.Info
    Write-Host "Failed:   $taskFailed" -ForegroundColor $tskColor

    Write-Host -NoNewline "  [3] " -ForegroundColor $C.Header
    Write-Host -NoNewline "Show recent event log errors " -ForegroundColor $C.Info
    Write-Host "Errors:   $evtCount" -ForegroundColor $evtColor

    Write-Host "  [4] Restart stopped critical services (interactive)" -ForegroundColor $C.Warning
    Write-Host "  [5] Export HTML report" -ForegroundColor $C.Info
    Write-Host "  [6] Connect to remote machine" -ForegroundColor $C.Info
    Write-Host "  [R] Refresh all" -ForegroundColor $C.Progress
    Write-Host "  [Q] Quit" -ForegroundColor $C.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
    $choice = (Read-Host).Trim().ToUpper()

    switch ($choice) {
        '1' {
            Show-SentinelBanner
            if (-not $cachedServices) {
                Write-Host "  [*] Collecting service health..." -ForegroundColor $C.Progress
                $cachedServices = Get-ServiceHealth
            }
            Show-ServiceHealth -Data $cachedServices
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
        '2' {
            Show-SentinelBanner
            if (-not $cachedTasks) {
                Write-Host "  [*] Collecting task audit..." -ForegroundColor $C.Progress
                $cachedTasks = Get-TaskAudit
            }
            Show-TaskAudit -Data $cachedTasks
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
        '3' {
            Show-SentinelBanner
            if (-not $cachedEvents) {
                Write-Host "  [*] Collecting event log errors (this may take a moment)..." -ForegroundColor $C.Progress
                $cachedEvents = Get-EventErrors
            }
            Show-EventErrors -Data $cachedEvents
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
        '4' {
            Show-SentinelBanner
            Write-Section "RESTART STOPPED CRITICAL SERVICES"
            if (-not $cachedServices) {
                Write-Host "  [*] Collecting service health..." -ForegroundColor $C.Progress
                $cachedServices = Get-ServiceHealth
            }
            Invoke-RestartService -ServiceData $cachedServices
            Write-Host ""
            Write-Host "  [*] Refreshing service data after restart..." -ForegroundColor $C.Progress
            $cachedServices = Get-ServiceHealth
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
        '5' {
            Show-SentinelBanner
            Write-Section "EXPORT HTML REPORT"
            if (-not $cachedServices) {
                Write-Host "  [*] Collecting service health..." -ForegroundColor $C.Progress
                $cachedServices = Get-ServiceHealth
            }
            if (-not $cachedTasks) {
                Write-Host "  [*] Collecting task audit..." -ForegroundColor $C.Progress
                $cachedTasks = Get-TaskAudit
            }
            if (-not $cachedEvents) {
                Write-Host "  [*] Collecting event log errors (this may take a moment)..." -ForegroundColor $C.Progress
                $cachedEvents = Get-EventErrors
            }
            Export-HtmlReport -Services $cachedServices -Tasks $cachedTasks -Events $cachedEvents
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
        '6' {
            Show-SentinelBanner
            Connect-RemoteTarget
            # Clear cached data so next action re-fetches from new target
            $cachedServices = $null
            $cachedTasks    = $null
            $cachedEvents   = $null
            if ($script:RemoteTarget) {
                Write-Host ""
                Write-Host "  [*] Loading data from $($script:RemoteTarget)..." -ForegroundColor $C.Progress
                Invoke-RefreshAll
            }
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
        'R' {
            Show-SentinelBanner
            Write-Section "REFRESH ALL DATA"
            Invoke-RefreshAll
            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }
        'Q' {
            Write-Host ""
            Write-Host "  Closing S.E.N.T.I.N.E.L." -ForegroundColor $C.Header
            Write-Host ""
        }
        default {
            Write-Host ""
            Write-Host "  [!!] Invalid selection. Enter 1-6, R, or Q." -ForegroundColor $C.Warning
            Start-Sleep -Seconds 1
        }
    }

} while ($choice -ne 'Q')

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
