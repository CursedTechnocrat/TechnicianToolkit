<#
.SYNOPSIS
    O.R.A.C.L.E. — Observes, Reports & Audits Computer Logs & Environments
    System Diagnostic & Health Assessment Tool for PowerShell 5.1+

.DESCRIPTION
    Audits and reports on the current state of a Windows machine. Collects
    hardware specs, OS info, network config, uptime, pending Windows Updates,
    installed software, and recent event log errors — then exports a
    dark-themed HTML report with color-coded indicators to the Desktop.

.USAGE
    PS C:\> .\oracle.ps1                    # Must be run as Administrator
    PS C:\> .\oracle.ps1 -Unattended        # Silent mode — no prompts, no banner

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

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success messages
    Yellow   Warnings and cautions
    Red      Critical errors
    Gray     Information and details
#>

param([switch]$Unattended)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script must be run as Administrator!" -ForegroundColor Red
    exit 1
}

# Set console to UTF-8 so Unicode block characters render correctly
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ─────────────────────────────────────────────────────────────────────────────
# BANNER DISPLAY
# ─────────────────────────────────────────────────────────────────────────────

function Show-OracleBanner {
    Write-Host @"

   ██████╗ ██████╗  █████╗  ██████╗██╗     ███████╗
  ██╔═══██╗██╔══██╗██╔══██╗██╔════╝██║     ██╔════╝
  ██║   ██║██████╔╝███████║██║     ██║     █████╗
  ██║   ██║██╔══██╗██╔══██║██║     ██║     ██╔══╝
  ╚██████╔╝██║  ██║██║  ██║╚██████╗███████╗███████╗
   ╚═════╝ ╚═╝  ╚═╝╚═╝  ╚═╝ ╚═════╝╚══════╝╚══════╝

"@ -ForegroundColor Cyan
    Write-Host "    O.R.A.C.L.E. — Observes, Reports & Audits Computer Logs & Environments" -ForegroundColor Cyan
    Write-Host "    System Diagnostic & Health Assessment Tool" -ForegroundColor Cyan
    Write-Host ""
}

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

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

# Folder where the HTML report will be saved.
# Change this to any valid path, e.g. "C:\Reports" or "\\server\share\Reports"
$ReportOutputPath = $ScriptPath

# ─────────────────────────────────────────────────────────────────────────────
# COLOR SCHEMA DEFINITION
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
# DISPLAY BANNER
# ─────────────────────────────────────────────────────────────────────────────

if (-not $Unattended) { Show-OracleBanner }

$reportTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$reportFilename  = "ORACLE_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$reportPath      = "$ReportOutputPath\$reportFilename"

Write-Host "════════════════════════════════════════════════" -ForegroundColor $ColorSchema.Header
Write-Host "     SYSTEM DIAGNOSTIC REPORT" -ForegroundColor $ColorSchema.Header
Write-Host "════════════════════════════════════════════════" -ForegroundColor $ColorSchema.Header
Write-Host "  Machine   : $env:COMPUTERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  Run As    : $env:USERDOMAIN\$env:USERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  Timestamp : $reportTimestamp" -ForegroundColor $ColorSchema.Info
Write-Host "  Report    : $reportPath" -ForegroundColor $ColorSchema.Info
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# DATA COLLECTION STORAGE
# ─────────────────────────────────────────────────────────────────────────────

$reportData = [ordered]@{}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1: HARDWARE
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[1/8] Collecting Hardware Info..." -ForegroundColor $ColorSchema.Progress

try {
    $cs        = Get-CimInstance -ClassName Win32_ComputerSystem
    $bios      = Get-CimInstance -ClassName Win32_BIOS
    $cpu       = Get-CimInstance -ClassName Win32_Processor
    $ramBytes  = ($cs.TotalPhysicalMemory)
    $ramGB     = [math]::Round($ramBytes / 1GB, 2)
    $disks     = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"

    Write-Host "    Manufacturer : $($cs.Manufacturer)" -ForegroundColor $ColorSchema.Info
    Write-Host "    Model        : $($cs.Model)" -ForegroundColor $ColorSchema.Info
    Write-Host "    Serial       : $($bios.SerialNumber)" -ForegroundColor $ColorSchema.Info
    Write-Host "    CPU          : $($cpu.Name)" -ForegroundColor $ColorSchema.Info
    Write-Host "    Cores/Threads: $($cpu.NumberOfCores) cores / $($cpu.NumberOfLogicalProcessors) threads" -ForegroundColor $ColorSchema.Info
    Write-Host "    RAM          : $ramGB GB" -ForegroundColor $ColorSchema.Info

    $diskSummary = @()
    foreach ($disk in $disks) {
        $totalGB = [math]::Round($disk.Size / 1GB, 1)
        $freeGB  = [math]::Round($disk.FreeSpace / 1GB, 1)
        $usedGB  = [math]::Round(($disk.Size - $disk.FreeSpace) / 1GB, 1)
        $pct     = if ($disk.Size -gt 0) { [math]::Round(($disk.Size - $disk.FreeSpace) / $disk.Size * 100, 1) } else { 0 }

        $color = if ($pct -ge 90) { $ColorSchema.Error } elseif ($pct -ge 75) { $ColorSchema.Warning } else { $ColorSchema.Info }
        Write-Host "    Disk $($disk.DeviceID)      : $usedGB GB used / $totalGB GB total ($pct% full)" -ForegroundColor $color

        $diskSummary += [PSCustomObject]@{
            Drive   = $disk.DeviceID
            Label   = $disk.VolumeName
            TotalGB = $totalGB
            UsedGB  = $usedGB
            FreeGB  = $freeGB
            PctUsed = $pct
        }
    }

    $reportData['Hardware'] = [PSCustomObject]@{
        Manufacturer = $cs.Manufacturer
        Model        = $cs.Model
        Serial       = $bios.SerialNumber
        CPU          = $cpu.Name
        Cores        = $cpu.NumberOfCores
        Threads      = $cpu.NumberOfLogicalProcessors
        RAMGB        = $ramGB
        Disks        = $diskSummary
    }

    Write-Host "[+] Hardware info collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting hardware info: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2: OPERATING SYSTEM
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[2/8] Collecting OS Info..." -ForegroundColor $ColorSchema.Progress

try {
    $os          = Get-CimInstance -ClassName Win32_OperatingSystem
    $osVersion   = $os.Caption
    $osBuild     = $os.BuildNumber
    $osArch      = $os.OSArchitecture
    $osInstall   = $os.InstallDate

    # Activation status via slmgr
    $licenseStatus = "Unknown"
    try {
        $slmgr = & cscript //nologo "$env:SystemRoot\System32\slmgr.vbs" /dli 2>&1
        $licenseLine = $slmgr | Where-Object { $_ -match "License Status" }
        if ($licenseLine) {
            $licenseStatus = ($licenseLine -replace "License Status:\s*", "").Trim()
        }
    }
    catch {
        $licenseStatus = "Could not query"
    }

    Write-Host "    OS       : $osVersion" -ForegroundColor $ColorSchema.Info
    Write-Host "    Build    : $osBuild  ($osArch)" -ForegroundColor $ColorSchema.Info
    Write-Host "    Installed: $(Get-Date $osInstall -Format 'yyyy-MM-dd')" -ForegroundColor $ColorSchema.Info

    $activationColor = if ($licenseStatus -match "Licensed") { $ColorSchema.Success } else { $ColorSchema.Warning }
    Write-Host "    Activation: $licenseStatus" -ForegroundColor $activationColor

    $reportData['OS'] = [PSCustomObject]@{
        Caption      = $osVersion
        Build        = $osBuild
        Architecture = $osArch
        InstallDate  = Get-Date $osInstall -Format 'yyyy-MM-dd'
        Activation   = $licenseStatus
    }

    Write-Host "[+] OS info collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting OS info: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3: NETWORK CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[3/8] Collecting Network Configuration..." -ForegroundColor $ColorSchema.Progress

try {
    $adapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True"
    $netSummary = @()

    foreach ($adapter in $adapters) {
        $ip      = ($adapter.IPAddress      | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' }) -join ', '
        $dns     = ($adapter.DNSServerSearchOrder) -join ', '
        $gateway = ($adapter.DefaultIPGateway) -join ', '

        Write-Host "    Adapter  : $($adapter.Description)" -ForegroundColor $ColorSchema.Info
        Write-Host "    IP       : $ip" -ForegroundColor $ColorSchema.Info
        Write-Host "    MAC      : $($adapter.MACAddress)" -ForegroundColor $ColorSchema.Info
        Write-Host "    Gateway  : $gateway" -ForegroundColor $ColorSchema.Info
        Write-Host "    DNS      : $dns" -ForegroundColor $ColorSchema.Info
        Write-Host ""

        $netSummary += [PSCustomObject]@{
            Adapter = $adapter.Description
            IP      = $ip
            MAC     = $adapter.MACAddress
            Gateway = $gateway
            DNS     = $dns
        }
    }

    $reportData['Network'] = $netSummary
    Write-Host "[+] Network config collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting network config: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4: SYSTEM HEALTH & UPTIME
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[4/8] Collecting System Health..." -ForegroundColor $ColorSchema.Progress

try {
    $os        = Get-CimInstance -ClassName Win32_OperatingSystem
    $lastBoot  = $os.LastBootUpTime
    $uptime    = (Get-Date) - $lastBoot
    $uptimeStr = "{0}d {1}h {2}m" -f $uptime.Days, $uptime.Hours, $uptime.Minutes

    Write-Host "    Last Boot : $(Get-Date $lastBoot -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor $ColorSchema.Info
    Write-Host "    Uptime    : $uptimeStr" -ForegroundColor $ColorSchema.Info

    # Battery (laptops only)
    $batteryInfo = $null
    $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
    if ($battery) {
        $charge = $battery.EstimatedChargeRemaining
        $status = switch ($battery.BatteryStatus) {
            1 { "Discharging" } 2 { "AC - Fully Charged" } 3 { "Fully Charged" }
            4 { "Low" } 5 { "Critical" } 6 { "Charging" } 7 { "Charging/High" }
            8 { "Charging/Low" } 9 { "Charging/Critical" } default { "Unknown" }
        }
        $battColor = if ($charge -lt 20) { $ColorSchema.Error } elseif ($charge -lt 40) { $ColorSchema.Warning } else { $ColorSchema.Success }
        Write-Host "    Battery   : $charge% ($status)" -ForegroundColor $battColor
        $batteryInfo = [PSCustomObject]@{ Charge = $charge; Status = $status }
    }
    else {
        Write-Host "    Battery   : N/A (desktop or not detected)" -ForegroundColor $ColorSchema.Info
    }

    $reportData['Health'] = [PSCustomObject]@{
        LastBoot = Get-Date $lastBoot -Format 'yyyy-MM-dd HH:mm:ss'
        Uptime   = $uptimeStr
        Battery  = $batteryInfo
    }

    Write-Host "[+] System health collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting system health: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 5: STORAGE & RAID HEALTH
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[5/8] Collecting Storage & RAID Health..." -ForegroundColor $ColorSchema.Progress

$physicalDiskSummary = @()
$virtualDiskSummary  = @()

try {
    $physicalDisks = Get-PhysicalDisk -ErrorAction Stop

    foreach ($pd in $physicalDisks) {
        $sizeGB = if ($pd.Size -gt 0) { [math]::Round($pd.Size / 1GB, 1) } else { 0 }
        $healthColor = switch ($pd.HealthStatus) {
            'Healthy' { $ColorSchema.Success }
            'Warning' { $ColorSchema.Warning }
            default   { $ColorSchema.Error   }
        }
        Write-Host "    [$($pd.DeviceId)] $($pd.FriendlyName) | $($pd.MediaType) | $sizeGB GB | Health: $($pd.HealthStatus) | Status: $($pd.OperationalStatus)" -ForegroundColor $healthColor

        $physicalDiskSummary += [PSCustomObject]@{
            ID                = $pd.DeviceId
            Name              = $pd.FriendlyName
            MediaType         = $pd.MediaType
            BusType           = $pd.BusType
            SizeGB            = $sizeGB
            HealthStatus      = $pd.HealthStatus
            OperationalStatus = $pd.OperationalStatus
        }
    }

    # Storage Spaces virtual disks (software RAID)
    $virtualDisks = Get-VirtualDisk -ErrorAction SilentlyContinue
    if ($virtualDisks) {
        Write-Host "    Storage Spaces (virtual disks):" -ForegroundColor $ColorSchema.Info
        foreach ($vd in $virtualDisks) {
            $vHealthColor = switch ($vd.HealthStatus) {
                'Healthy' { $ColorSchema.Success }
                'Warning' { $ColorSchema.Warning }
                default   { $ColorSchema.Error   }
            }
            $vSizeGB = if ($vd.Size -gt 0) { [math]::Round($vd.Size / 1GB, 1) } else { 0 }
            Write-Host "      VDisk: $($vd.FriendlyName) | $($vd.ResiliencySettingName) | $vSizeGB GB | Health: $($vd.HealthStatus) / Op: $($vd.OperationalStatus)" -ForegroundColor $vHealthColor

            $virtualDiskSummary += [PSCustomObject]@{
                Name              = $vd.FriendlyName
                ResiliencyType    = $vd.ResiliencySettingName
                SizeGB            = $vSizeGB
                HealthStatus      = $vd.HealthStatus
                OperationalStatus = $vd.OperationalStatus
            }
        }
    }
    else {
        Write-Host "    No Storage Spaces virtual disks detected" -ForegroundColor $ColorSchema.Info
    }

    $reportData['Storage'] = [PSCustomObject]@{
        PhysicalDisks = $physicalDiskSummary
        VirtualDisks  = $virtualDiskSummary
    }

    Write-Host "[+] Storage health collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting storage health: $_" -ForegroundColor $ColorSchema.Error
    $reportData['Storage'] = [PSCustomObject]@{
        PhysicalDisks = @()
        VirtualDisks  = @()
    }
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6: PENDING WINDOWS UPDATES
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[6/8] Scanning for Pending Updates..." -ForegroundColor $ColorSchema.Progress
Write-Host "    This may take a moment..." -ForegroundColor $ColorSchema.Info

$pendingUpdates = @()
try {
    $updateSession  = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    $searchResult   = $updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

    if ($searchResult.Updates.Count -eq 0) {
        Write-Host "    [+] No pending updates found" -ForegroundColor $ColorSchema.Success
    }
    else {
        Write-Host "    [!!] $($searchResult.Updates.Count) pending update(s) found:" -ForegroundColor $ColorSchema.Warning
        foreach ($update in $searchResult.Updates) {
            Write-Host "      * $($update.Title)" -ForegroundColor $ColorSchema.Warning
            $pendingUpdates += [PSCustomObject]@{
                Title    = $update.Title
                Severity = if ($update.MsrcSeverity) { $update.MsrcSeverity } else { "N/A" }
                KB       = ($update.KBArticleIDs -join ', ')
            }
        }
    }

    $reportData['Updates'] = $pendingUpdates
    Write-Host "[+] Update scan complete" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error scanning for updates: $_" -ForegroundColor $ColorSchema.Error
    $reportData['Updates'] = @()
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6: INSTALLED SOFTWARE
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[7/8] Collecting Installed Software..." -ForegroundColor $ColorSchema.Progress

$installedApps = @()
try {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $regPaths) {
        Get-ItemProperty $path -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName -ne '' } |
            ForEach-Object {
                $installedApps += [PSCustomObject]@{
                    Name      = $_.DisplayName
                    Version   = if ($_.DisplayVersion) { $_.DisplayVersion } else { "N/A" }
                    Publisher = if ($_.Publisher)       { $_.Publisher }       else { "N/A" }
                    InstallDate = if ($_.InstallDate)   { $_.InstallDate }     else { "N/A" }
                }
            }
    }

    # Deduplicate by name and sort
    $installedApps = $installedApps | Sort-Object Name -Unique

    Write-Host "    Found $($installedApps.Count) installed application(s)" -ForegroundColor $ColorSchema.Info
    $reportData['Software'] = $installedApps
    Write-Host "[+] Software list collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting software list: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 7: RECENT EVENT LOG ERRORS
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[8/8] Scanning Event Logs (last 24 hours)..." -ForegroundColor $ColorSchema.Progress

$eventSummary = @()
try {
    $since  = (Get-Date).AddHours(-24)
    $levels = @(1, 2)  # 1 = Critical, 2 = Error

    $events = Get-WinEvent -FilterHashtable @{
        LogName   = 'System', 'Application'
        Level     = $levels
        StartTime = $since
    } -ErrorAction SilentlyContinue -MaxEvents 50

    if ($null -eq $events -or $events.Count -eq 0) {
        Write-Host "    [+] No critical/error events in the last 24 hours" -ForegroundColor $ColorSchema.Success
    }
    else {
        Write-Host "    [!!] $($events.Count) error/critical event(s) found in the last 24 hours" -ForegroundColor $ColorSchema.Warning
        $events | Select-Object -First 10 | ForEach-Object {
            $lvl = if ($_.Level -eq 1) { "CRITICAL" } else { "ERROR" }
            Write-Host "      [$lvl] $(Get-Date $_.TimeCreated -Format 'HH:mm:ss') | $($_.ProviderName) — $($_.Message.Split([Environment]::NewLine)[0])" -ForegroundColor $ColorSchema.Warning
        }
        if ($events.Count -gt 10) {
            Write-Host "      ... and $($events.Count - 10) more (see full report)" -ForegroundColor $ColorSchema.Info
        }

        $eventSummary = $events | ForEach-Object {
            [PSCustomObject]@{
                Time     = Get-Date $_.TimeCreated -Format 'yyyy-MM-dd HH:mm:ss'
                Level    = if ($_.Level -eq 1) { "Critical" } else { "Error" }
                Source   = $_.ProviderName
                Log      = $_.LogName
                Message  = $_.Message.Split([Environment]::NewLine)[0]
                EventID  = $_.Id
            }
        }
    }

    $reportData['Events'] = $eventSummary
    Write-Host "[+] Event log scan complete" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error reading event logs: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# GENERATE HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "Generating HTML report..." -ForegroundColor $ColorSchema.Progress

function ConvertTo-HtmlTable {
    param([array]$Objects, [string]$EmptyMessage = "No data available.")
    if (-not $Objects -or $Objects.Count -eq 0) {
        return "<p class='empty'>$EmptyMessage</p>"
    }
    $headers = $Objects[0].PSObject.Properties.Name
    $html  = "<table><thead><tr>"
    $html += ($headers | ForEach-Object { "<th>$_</th>" }) -join ""
    $html += "</tr></thead><tbody>"
    foreach ($row in $Objects) {
        $html += "<tr>"
        foreach ($h in $headers) {
            $val = $row.$h
            $escaped = "$val" -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
            $html += "<td>$escaped</td>"
        }
        $html += "</tr>"
    }
    $html += "</tbody></table>"
    return $html
}

# Storage health badge
$storageUnhealthy = ($reportData['Storage'].PhysicalDisks | Where-Object { $_.HealthStatus -ne 'Healthy' }).Count
$storageBadge = if ($storageUnhealthy -eq 0) {
    "<span class='badge badge-ok'>Healthy</span>"
} else {
    "<span class='badge badge-err'>$storageUnhealthy degraded</span>"
}

# Physical disk rows
$physDiskRows = ""
foreach ($pd in $reportData['Storage'].PhysicalDisks) {
    $hColor = switch ($pd.HealthStatus) {
        'Healthy' { '#2ecc71' } 'Warning' { '#f39c12' } default { '#e74c3c' }
    }
    $physDiskRows += "<tr><td>$($pd.ID)</td><td>$($pd.Name)</td><td>$($pd.MediaType)</td><td>$($pd.BusType)</td><td>$($pd.SizeGB) GB</td>"
    $physDiskRows += "<td style='color:$hColor;font-weight:600;'>$($pd.HealthStatus)</td><td>$($pd.OperationalStatus)</td></tr>"
}

# Virtual disk rows
$virtDiskRows = ""
foreach ($vd in $reportData['Storage'].VirtualDisks) {
    $hColor = switch ($vd.HealthStatus) {
        'Healthy' { '#2ecc71' } 'Warning' { '#f39c12' } default { '#e74c3c' }
    }
    $virtDiskRows += "<tr><td>$($vd.Name)</td><td>$($vd.ResiliencyType)</td><td>$($vd.SizeGB) GB</td>"
    $virtDiskRows += "<td style='color:$hColor;font-weight:600;'>$($vd.HealthStatus)</td><td>$($vd.OperationalStatus)</td></tr>"
}

# Disk rows for hardware section
$diskRows = ""
foreach ($d in $reportData['Hardware'].Disks) {
    $barColor = if ($d.PctUsed -ge 90) { "#e74c3c" } elseif ($d.PctUsed -ge 75) { "#f39c12" } else { "#2ecc71" }
    $diskRows += @"
        <tr>
            <td>$($d.Drive)</td>
            <td>$($d.Label)</td>
            <td>$($d.TotalGB) GB</td>
            <td>$($d.UsedGB) GB</td>
            <td>$($d.FreeGB) GB</td>
            <td>
                <div style='background:#444;border-radius:4px;height:14px;width:100%;'>
                    <div style='background:$barColor;width:$($d.PctUsed)%;height:14px;border-radius:4px;'></div>
                </div>
                $($d.PctUsed)%
            </td>
        </tr>
"@
}

# Network rows
$netRows = ""
foreach ($n in $reportData['Network']) {
    $netRows += "<tr><td>$($n.Adapter)</td><td>$($n.IP)</td><td>$($n.MAC)</td><td>$($n.Gateway)</td><td>$($n.DNS)</td></tr>"
}

# Battery
$hw = $reportData['Hardware']
$os = $reportData['OS']
$health = $reportData['Health']
$batteryHtml = if ($health.Battery) {
    "<p><strong>Battery:</strong> $($health.Battery.Charge)% — $($health.Battery.Status)</p>"
} else {
    "<p><strong>Battery:</strong> N/A (desktop or not detected)</p>"
}

# Updates badge
$updateCount = $reportData['Updates'].Count
$updateBadge = if ($updateCount -eq 0) {
    "<span class='badge badge-ok'>Up to date</span>"
} else {
    "<span class='badge badge-warn'>$updateCount pending</span>"
}
$updatesTable = if ($updateCount -gt 0) {
    ConvertTo-HtmlTable -Objects $reportData['Updates'] -EmptyMessage "No pending updates."
} else {
    "<p class='empty'>System is fully up to date.</p>"
}

# Events badge
$eventCount = $reportData['Events'].Count
$eventBadge = if ($eventCount -eq 0) {
    "<span class='badge badge-ok'>Clean</span>"
} else {
    "<span class='badge badge-warn'>$eventCount events</span>"
}

$softwareTable = ConvertTo-HtmlTable -Objects $reportData['Software'] -EmptyMessage "No software found."
$eventsTable   = ConvertTo-HtmlTable -Objects $reportData['Events']   -EmptyMessage "No critical/error events in the last 24 hours."

$htmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="dark">
<title>O.R.A.C.L.E. — $env:COMPUTERNAME — $reportTimestamp</title>
<style>
  :root { color-scheme: dark; }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', sans-serif; background: #1a1a2e; color: #e0e0e0; font-size: 14px; }
  header { background: linear-gradient(135deg, #0f3460, #16213e); padding: 28px 40px; border-bottom: 3px solid #00d4ff; }
  header h1 { color: #00d4ff; font-size: 2em; letter-spacing: 4px; font-weight: 700; }
  header p  { color: #aaa; margin-top: 6px; font-size: 0.9em; }
  header .meta { display: flex; gap: 30px; margin-top: 14px; flex-wrap: wrap; }
  header .meta span { color: #ccc; font-size: 0.85em; }
  header .meta strong { color: #00d4ff; }
  main { padding: 30px 40px; max-width: 1400px; margin: 0 auto; }
  section { background: #16213e; border-radius: 8px; margin-bottom: 24px; overflow: hidden;
            border: 1px solid #0f3460; }
  section h2 { background: #0f3460; color: #00d4ff; padding: 14px 20px; font-size: 1em;
               letter-spacing: 2px; text-transform: uppercase; display: flex; align-items: center; gap: 10px; }
  section .content { padding: 20px; }
  .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
  .kv { display: flex; gap: 10px; padding: 6px 0; border-bottom: 1px solid #0f3460; }
  .kv:last-child { border-bottom: none; }
  .kv .key   { color: #888; min-width: 140px; font-size: 0.85em; }
  .kv .value { color: #e0e0e0; font-weight: 500; }
  table { width: 100%; border-collapse: collapse; font-size: 0.85em; }
  th { background: #0f3460; color: #00d4ff; padding: 10px 12px; text-align: left;
       font-weight: 600; letter-spacing: 1px; text-transform: uppercase; font-size: 0.78em; }
  td { padding: 8px 12px; border-bottom: 1px solid #1e3a5f; color: #ccc; }
  tr:hover td { background: #1e3a5f; }
  .empty { color: #666; padding: 12px 0; font-style: italic; }
  .badge { display: inline-block; padding: 3px 10px; border-radius: 20px; font-size: 0.78em;
           font-weight: 700; letter-spacing: 1px; text-transform: uppercase; }
  .badge-ok   { background: #1a4a2e; color: #2ecc71; border: 1px solid #2ecc71; }
  .badge-warn { background: #4a3000; color: #f39c12; border: 1px solid #f39c12; }
  .badge-err  { background: #4a0000; color: #e74c3c; border: 1px solid #e74c3c; }
  footer { text-align: center; padding: 20px; color: #444; font-size: 0.8em; border-top: 1px solid #0f3460; }
</style>
</head>
<body>
<header>
  <h1>O.R.A.C.L.E.</h1>
  <p>Observes, Reports &amp; Audits Computer Logs &amp; Environments</p>
  <div class="meta">
    <span><strong>Machine:</strong> $env:COMPUTERNAME</span>
    <span><strong>Run As:</strong> $env:USERDOMAIN\$env:USERNAME</span>
    <span><strong>Generated:</strong> $reportTimestamp</span>
    <span><strong>Storage:</strong> $storageBadge</span>
    <span><strong>Updates:</strong> $updateBadge</span>
    <span><strong>Events (24h):</strong> $eventBadge</span>
  </div>
</header>
<main>

  <!-- HARDWARE -->
  <section>
    <h2>Hardware</h2>
    <div class="content">
      <div class="grid-2">
        <div>
          <div class="kv"><span class="key">Manufacturer</span><span class="value">$($hw.Manufacturer)</span></div>
          <div class="kv"><span class="key">Model</span><span class="value">$($hw.Model)</span></div>
          <div class="kv"><span class="key">Serial Number</span><span class="value">$($hw.Serial)</span></div>
        </div>
        <div>
          <div class="kv"><span class="key">CPU</span><span class="value">$($hw.CPU)</span></div>
          <div class="kv"><span class="key">Cores / Threads</span><span class="value">$($hw.Cores) / $($hw.Threads)</span></div>
          <div class="kv"><span class="key">RAM</span><span class="value">$($hw.RAMGB) GB</span></div>
        </div>
      </div>
      <br>
      <table>
        <thead><tr><th>Drive</th><th>Label</th><th>Total</th><th>Used</th><th>Free</th><th>Usage</th></tr></thead>
        <tbody>$diskRows</tbody>
      </table>
    </div>
  </section>

  <!-- OPERATING SYSTEM -->
  <section>
    <h2>Operating System</h2>
    <div class="content">
      <div class="kv"><span class="key">OS</span><span class="value">$($os.Caption)</span></div>
      <div class="kv"><span class="key">Build</span><span class="value">$($os.Build)</span></div>
      <div class="kv"><span class="key">Architecture</span><span class="value">$($os.Architecture)</span></div>
      <div class="kv"><span class="key">Install Date</span><span class="value">$($os.InstallDate)</span></div>
      <div class="kv"><span class="key">Activation</span><span class="value">$($os.Activation)</span></div>
    </div>
  </section>

  <!-- NETWORK -->
  <section>
    <h2>Network Configuration</h2>
    <div class="content">
      <table>
        <thead><tr><th>Adapter</th><th>IP Address</th><th>MAC</th><th>Gateway</th><th>DNS</th></tr></thead>
        <tbody>$netRows</tbody>
      </table>
    </div>
  </section>

  <!-- HEALTH -->
  <section>
    <h2>System Health</h2>
    <div class="content">
      <div class="kv"><span class="key">Last Boot</span><span class="value">$($health.LastBoot)</span></div>
      <div class="kv"><span class="key">Uptime</span><span class="value">$($health.Uptime)</span></div>
      <div class="kv"><span class="key">Battery</span><span class="value">$(if ($health.Battery) { "$($health.Battery.Charge)% — $($health.Battery.Status)" } else { "N/A" })</span></div>
    </div>
  </section>

  <!-- STORAGE & RAID HEALTH -->
  <section>
    <h2>Storage &amp; RAID Health $storageBadge</h2>
    <div class="content">
      <h3 style="color:#00d4ff;font-size:0.85em;letter-spacing:1px;text-transform:uppercase;margin-bottom:10px;">Physical Disks</h3>
      $(if ($physDiskRows) {
        "<table><thead><tr><th>ID</th><th>Name</th><th>Type</th><th>Bus</th><th>Size</th><th>Health</th><th>Status</th></tr></thead><tbody>$physDiskRows</tbody></table>"
      } else {
        "<p class='empty'>No physical disk data available.</p>"
      })
      $(if ($virtDiskRows) {
        "<h3 style='color:#00d4ff;font-size:0.85em;letter-spacing:1px;text-transform:uppercase;margin:16px 0 10px;'>Storage Spaces (Virtual Disks)</h3><table><thead><tr><th>Name</th><th>Resiliency</th><th>Size</th><th>Health</th><th>Status</th></tr></thead><tbody>$virtDiskRows</tbody></table>"
      } else {
        "<p style='color:#666;font-size:0.85em;margin-top:12px;'>No Storage Spaces virtual disks detected.</p>"
      })
    </div>
  </section>

  <!-- PENDING UPDATES -->
  <section>
    <h2>Pending Windows Updates $updateBadge</h2>
    <div class="content">$updatesTable</div>
  </section>

  <!-- INSTALLED SOFTWARE -->
  <section>
    <h2>Installed Software ($($reportData['Software'].Count) apps)</h2>
    <div class="content">$softwareTable</div>
  </section>

  <!-- EVENT LOG -->
  <section>
    <h2>Event Log — Errors &amp; Critical (Last 24h) $eventBadge</h2>
    <div class="content">$eventsTable</div>
  </section>

</main>
<footer>
  Generated by O.R.A.C.L.E. — Part of the Technician Toolkit &nbsp;|&nbsp; $reportTimestamp
</footer>
</body>
</html>
"@

try {
    $htmlReport | Out-File -FilePath $reportPath -Encoding UTF8 -Force
    Write-Host "[+] Report saved: $reportPath" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error saving report: $_" -ForegroundColor $ColorSchema.Error
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "════════════════════════════════════════════════" -ForegroundColor $ColorSchema.Header
Write-Host "     PROBE COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host "════════════════════════════════════════════════" -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host "  Machine    : $env:COMPUTERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  OS         : $($reportData['OS'].Caption)" -ForegroundColor $ColorSchema.Info
Write-Host "  RAM        : $($reportData['Hardware'].RAMGB) GB" -ForegroundColor $ColorSchema.Info
Write-Host "  Uptime     : $($reportData['Health'].Uptime)" -ForegroundColor $ColorSchema.Info

$physDiskCount = $reportData['Storage'].PhysicalDisks.Count
if ($storageUnhealthy -gt 0) {
    Write-Host "  Storage    : $storageUnhealthy degraded disk(s) detected!" -ForegroundColor $ColorSchema.Error
} else {
    Write-Host "  Storage    : $physDiskCount disk(s) — all healthy" -ForegroundColor $ColorSchema.Success
}

if ($updateCount -gt 0) {
    Write-Host "  Updates    : $updateCount pending" -ForegroundColor $ColorSchema.Warning
}
else {
    Write-Host "  Updates    : Up to date" -ForegroundColor $ColorSchema.Success
}

if ($eventCount -gt 0) {
    Write-Host "  Events     : $eventCount error(s) in last 24h" -ForegroundColor $ColorSchema.Warning
}
else {
    Write-Host "  Events     : Clean" -ForegroundColor $ColorSchema.Success
}

Write-Host ""
Write-Host "  Report     : $reportPath" -ForegroundColor $ColorSchema.Accent
Write-Host ""

if (-not $Unattended) {
    $openReport = Read-Host "Open the HTML report now? (Y/N)"
    if ($openReport -eq 'Y' -or $openReport -eq 'y') {
        try {
            Start-Process $reportPath
            Write-Host "[+] Opening report in default browser..." -ForegroundColor $ColorSchema.Success
        }
        catch {
            Write-Host "[-] Could not open report automatically. Navigate to: $reportPath" -ForegroundColor $ColorSchema.Warning
        }
    }
}

Write-Host ""
Write-Host "════════════════════════════════════════════════" -ForegroundColor $ColorSchema.Header
Write-Host "  SCRIPT EXECUTION COMPLETED" -ForegroundColor $ColorSchema.Header
Write-Host "════════════════════════════════════════════════" -ForegroundColor $ColorSchema.Header
Write-Host ""
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
