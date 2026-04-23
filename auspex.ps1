<#
.SYNOPSIS
    A.U.S.P.E.X. — Audits, Uncovers, Surveys Performance, Events & eXceptions
    System Diagnostic & Health Assessment Tool for PowerShell 5.1+

.DESCRIPTION
    Audits and reports on the current state of a Windows machine. Collects
    hardware specs, OS info, network config, uptime, pending Windows Updates,
    installed software, and recent event log errors — then exports a
    dark-themed HTML report with color-coded indicators to the Desktop.

.USAGE
    PS C:\> .\auspex.ps1                    # Must be run as Administrator
    PS C:\> .\auspex.ps1 -Unattended        # Silent mode — no prompts, no banner

.NOTES
    Version : 3.0

#>

param(
    [switch]$Unattended,
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

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
Assert-AdminPrivilege

# ─────────────────────────────────────────────────────────────────────────────
# BANNER DISPLAY
# ─────────────────────────────────────────────────────────────────────────────

function Show-AuspexBanner {
    Write-Host @"

   █████╗ ██╗   ██╗███████╗██████╗ ███████╗██╗  ██╗
  ██╔══██╗██║   ██║██╔════╝██╔══██╗██╔════╝╚██╗██╔╝
  ███████║██║   ██║███████╗██████╔╝█████╗   ╚███╔╝
  ██╔══██║██║   ██║╚════██║██╔═══╝ ██╔══╝   ██╔██╗
  ██║  ██║╚██████╔╝███████║██║     ███████╗██╔╝ ██╗
  ╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═╝     ╚══════╝╚═╝  ╚═╝

"@ -ForegroundColor Cyan
    Write-Host "    A.U.S.P.E.X. — Audits, Uncovers, Surveys Performance, Events & eXceptions" -ForegroundColor Cyan
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

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $ScriptPath) }

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

# Folder where the HTML report will be saved.
# Override via config.json LogDirectory, or set $ReportOutputPath manually below.
$ReportOutputPath = Resolve-LogDirectory -FallbackPath $ScriptPath

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

if (-not $Unattended) { Show-AuspexBanner }

$reportTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$reportFilename  = "AUSPEX_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
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

Write-Host "[1/10] Collecting Hardware Info..." -ForegroundColor $ColorSchema.Progress

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

Write-Host "[2/10] Collecting OS Info..." -ForegroundColor $ColorSchema.Progress

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

Write-Host "[3/10] Collecting Network Configuration..." -ForegroundColor $ColorSchema.Progress

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

Write-Host "[4/10] Collecting System Health..." -ForegroundColor $ColorSchema.Progress

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

Write-Host "[5/10] Collecting Storage & RAID Health..." -ForegroundColor $ColorSchema.Progress

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

    # Hardware RAID controller detection via WMI
    $raidControllers = @()
    try {
        $scsiControllers = Get-CimInstance -ClassName Win32_SCSIController -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match 'RAID|MegaRAID|PERC|Smart Array|LSI|Adaptec|Avago|Broadcom|RST|Intel.*RAID|Areca|3ware' }

        foreach ($ctrl in $scsiControllers) {
            $ctrlColor = if ($ctrl.Status -eq 'OK') { $ColorSchema.Success } else { $ColorSchema.Warning }
            Write-Host "    RAID Controller : $($ctrl.Name) [$($ctrl.Status)]" -ForegroundColor $ctrlColor
            $raidControllers += [PSCustomObject]@{
                Name         = $ctrl.Name
                Manufacturer = if ($ctrl.Manufacturer) { $ctrl.Manufacturer } else { 'N/A' }
                Status       = $ctrl.Status
                DriverName   = if ($ctrl.DriverName) { $ctrl.DriverName } else { 'N/A' }
            }
        }

        if ($raidControllers.Count -eq 0) {
            Write-Host "    No dedicated hardware RAID controllers detected" -ForegroundColor $ColorSchema.Info
        }
    }
    catch {
        Write-Host "    Could not query RAID controllers: $_" -ForegroundColor $ColorSchema.Warning
    }

    # Vendor CLI tools — capture detailed RAID config if available
    $raidVendorOutput = @()
    $cliCandidates = @(
        @{ Name = 'StorCLI';   Exe = 'StorCLI64.exe'; Dirs = @('C:\Windows\System32','C:\Program Files\MegaRAID\StorCLI','C:\Program Files (x86)\MegaRAID\StorCLI');         Args = @('/call','show') },
        @{ Name = 'PERCCLI';   Exe = 'perccli64.exe'; Dirs = @('C:\Windows\System32','C:\Program Files\PERCCLI','C:\Program Files (x86)\PERCCLI');                            Args = @('/call','show') },
        @{ Name = 'HP SSACLI'; Exe = 'ssacli.exe';    Dirs = @('C:\Program Files\Smart Storage Administrator\ssacli\bin','C:\Program Files (x86)\Compaq\Hpacucli\Bin');       Args = @('ctrl','all','show','config') }
    )

    foreach ($cli in $cliCandidates) {
        $exePath = $cli.Dirs | ForEach-Object { Join-Path $_ $cli.Exe } | Where-Object { Test-Path $_ } | Select-Object -First 1
        if ($exePath) {
            try {
                Write-Host "    Found $($cli.Name): $exePath" -ForegroundColor $ColorSchema.Info
                $rawOutput = & "$exePath" @($cli.Args) 2>&1 | Out-String
                $raidVendorOutput += [PSCustomObject]@{ Tool = $cli.Name; Output = $rawOutput.Trim() }
                Write-Host "    [+] $($cli.Name) output captured" -ForegroundColor $ColorSchema.Success
            }
            catch {
                Write-Host "    $($cli.Name) found but failed to run: $_" -ForegroundColor $ColorSchema.Warning
            }
        }
    }

    $reportData['Storage'] = [PSCustomObject]@{
        PhysicalDisks    = $physicalDiskSummary
        VirtualDisks     = $virtualDiskSummary
        RaidControllers  = $raidControllers
        RaidVendorOutput = $raidVendorOutput
    }

    Write-Host "[+] Storage health collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting storage health: $_" -ForegroundColor $ColorSchema.Error
    $reportData['Storage'] = [PSCustomObject]@{
        PhysicalDisks    = @()
        VirtualDisks     = @()
        RaidControllers  = @()
        RaidVendorOutput = @()
    }
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6: PENDING WINDOWS UPDATES
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[6/10] Scanning for Pending Updates..." -ForegroundColor $ColorSchema.Progress
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

Write-Host "[7/10] Collecting Installed Software..." -ForegroundColor $ColorSchema.Progress

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

Write-Host "[8/10] Scanning Event Logs (last 24 hours)..." -ForegroundColor $ColorSchema.Progress

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

# ─────────────────────────────────────────────────────────────────────────────
# STEP 9: SCHEDULED TASKS (NON-MICROSOFT)
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[9/10] Collecting Scheduled Tasks..." -ForegroundColor $ColorSchema.Progress

$scheduledTaskSummary = @()
try {
    $tasks = Get-ScheduledTask -ErrorAction Stop |
        Where-Object { $_.TaskPath -notmatch '^\\Microsoft\\' }

    foreach ($task in $tasks) {
        $info = Get-ScheduledTaskInfo -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction SilentlyContinue
        $actions = ($task.Actions | ForEach-Object {
            if ($_.Execute) { "$($_.Execute) $($_.Arguments)".Trim() }
        }) -join '; '

        $lastRun = if ($info -and $info.LastRunTime -and $info.LastRunTime -gt [DateTime]::MinValue) {
            $info.LastRunTime.ToString('yyyy-MM-dd HH:mm')
        } else { 'Never' }

        $nextRun = if ($info -and $info.NextRunTime -and $info.NextRunTime -gt [DateTime]::MinValue) {
            $info.NextRunTime.ToString('yyyy-MM-dd HH:mm')
        } else { 'N/A' }

        $scheduledTaskSummary += [PSCustomObject]@{
            Name       = $task.TaskName
            Path       = $task.TaskPath
            State      = $task.State
            LastRun    = $lastRun
            NextRun    = $nextRun
            LastResult = if ($info) { "0x{0:X8}" -f $info.LastTaskResult } else { 'N/A' }
            Actions    = $actions
        }
    }

    Write-Host "    Found $($scheduledTaskSummary.Count) non-Microsoft scheduled task(s)" -ForegroundColor $ColorSchema.Info
    $reportData['ScheduledTasks'] = $scheduledTaskSummary
    Write-Host "[+] Scheduled tasks collected" -ForegroundColor $ColorSchema.Success
}
catch {
    Write-Host "[-] Error collecting scheduled tasks: $_" -ForegroundColor $ColorSchema.Error
    $reportData['ScheduledTasks'] = @()
}

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# STEP 10: SECURITY / AV STATUS
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "[10/10] Collecting Security & AV Status..." -ForegroundColor $ColorSchema.Progress

$avProducts = @()
try {
    $defender = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if ($defender) {
        $defAgeDays = if ($defender.AntivirusSignatureLastUpdated) {
            [int](New-TimeSpan -Start $defender.AntivirusSignatureLastUpdated -End (Get-Date)).TotalDays
        } else { $null }

        $lastScan = if ($defender.QuickScanEndTime -and $defender.QuickScanEndTime -ne [datetime]::MinValue) {
            Get-Date $defender.QuickScanEndTime -Format 'yyyy-MM-dd HH:mm'
        } else { 'Never' }

        $avProducts += [PSCustomObject]@{
            Product            = 'Windows Defender'
            RealTimeProtection = if ($defender.RealTimeProtectionEnabled) { 'On' } else { 'Off' }
            DefinitionAge      = if ($null -ne $defAgeDays) { "$defAgeDays day(s)" } else { 'Unknown' }
            LastQuickScan      = $lastScan
            ServiceEnabled     = if ($defender.AMServiceEnabled) { 'Yes' } else { 'No' }
        }
        Write-Host "[+] Windows Defender status collected" -ForegroundColor $ColorSchema.Success
    }
}
catch {
    Write-Host "[-] Could not read Windows Defender status: $_" -ForegroundColor $ColorSchema.Warning
}

try {
    $thirdParty = Get-WmiObject -Namespace 'root\SecurityCenter2' -Class AntiVirusProduct -ErrorAction SilentlyContinue |
                  Where-Object { $_.displayName -notmatch 'Windows Defender|Microsoft Defender' }
    foreach ($av in $thirdParty) {
        $avProducts += [PSCustomObject]@{
            Product            = $av.displayName
            RealTimeProtection = 'N/A'
            DefinitionAge      = 'N/A'
            LastQuickScan      = 'N/A'
            ServiceEnabled     = 'Registered'
        }
    }
}
catch {
    # SecurityCenter2 namespace is absent on Server SKUs and some Core editions.
}

$reportData['Security'] = $avProducts

Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# GENERATE HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "Generating HTML report..." -ForegroundColor $ColorSchema.Progress

function ConvertTo-HtmlTable {
    param([array]$Objects, [string]$EmptyMessage = "No data available.")
    if (-not $Objects -or $Objects.Count -eq 0) {
        return "<p class='tk-info-box'>$EmptyMessage</p>"
    }
    $headers = $Objects[0].PSObject.Properties.Name
    $html  = "<table class='tk-table'><thead><tr>"
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
$storageUnhealthy  = ($reportData['Storage'].PhysicalDisks   | Where-Object { $_.HealthStatus -ne 'Healthy' }).Count
$raidCtrlUnhealthy = ($reportData['Storage'].RaidControllers  | Where-Object { $_.Status       -ne 'OK'      }).Count
$storageBadge = if ($storageUnhealthy -eq 0 -and $raidCtrlUnhealthy -eq 0) {
    "<span class='tk-badge-ok'>Healthy</span>"
} elseif ($raidCtrlUnhealthy -gt 0) {
    "<span class='tk-badge-err'>$raidCtrlUnhealthy controller issue(s)</span>"
} else {
    "<span class='tk-badge-err'>$storageUnhealthy degraded</span>"
}

# Physical disk rows
$physDiskRows = ""
foreach ($pd in $reportData['Storage'].PhysicalDisks) {
    $hBadge = switch ($pd.HealthStatus) {
        'Healthy' { "tk-badge-ok" } 'Warning' { "tk-badge-warn" } default { "tk-badge-err" }
    }
    $physDiskRows += "<tr><td>$($pd.ID)</td><td>$($pd.Name)</td><td>$($pd.MediaType)</td><td>$($pd.BusType)</td><td>$($pd.SizeGB) GB</td>"
    $physDiskRows += "<td><span class='$hBadge'>$($pd.HealthStatus)</span></td><td>$($pd.OperationalStatus)</td></tr>"
}

# Virtual disk rows
$virtDiskRows = ""
foreach ($vd in $reportData['Storage'].VirtualDisks) {
    $hBadge = switch ($vd.HealthStatus) {
        'Healthy' { "tk-badge-ok" } 'Warning' { "tk-badge-warn" } default { "tk-badge-err" }
    }
    $virtDiskRows += "<tr><td>$($vd.Name)</td><td>$($vd.ResiliencyType)</td><td>$($vd.SizeGB) GB</td>"
    $virtDiskRows += "<td><span class='$hBadge'>$($vd.HealthStatus)</span></td><td>$($vd.OperationalStatus)</td></tr>"
}

# RAID controller rows
$raidCtrlRows = ""
foreach ($rc in $reportData['Storage'].RaidControllers) {
    $rcBadge = if ($rc.Status -eq 'OK') { "tk-badge-ok" } else { "tk-badge-warn" }
    $raidCtrlRows += "<tr><td>$($rc.Name)</td><td>$($rc.Manufacturer)</td>"
    $raidCtrlRows += "<td><span class='$rcBadge'>$($rc.Status)</span></td><td>$($rc.DriverName)</td></tr>"
}

# Vendor CLI output blocks
$raidVendorHtml = ""
foreach ($rv in $reportData['Storage'].RaidVendorOutput) {
    $escaped = $rv.Output -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
    $raidVendorHtml += "<div class='tk-info-box' style='margin:14px 0 6px;'><span class='tk-info-label'>$($rv.Tool) Output</span></div>"
    $raidVendorHtml += "<pre class='tk-mono' style='padding:12px;font-size:0.78em;overflow-x:auto;white-space:pre-wrap;'>$escaped</pre>"
}

# Disk rows for hardware section
$diskRows = ""
foreach ($d in $reportData['Hardware'].Disks) {
    $barClass = if ($d.PctUsed -ge 90) { "err" } elseif ($d.PctUsed -ge 75) { "warn" } else { "ok" }
    $diskRows += @"
        <tr>
            <td>$($d.Drive)</td>
            <td>$($d.Label)</td>
            <td>$($d.TotalGB) GB</td>
            <td>$($d.UsedGB) GB</td>
            <td>$($d.FreeGB) GB</td>
            <td>
                <div class='tk-progress-wrap'>
                    <div class='tk-progress-bar $barClass' style='width:$($d.PctUsed)%;'></div>
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
    "<p><strong>Battery:</strong> $($health.Battery.Charge)%  -  $($health.Battery.Status)</p>"
} else {
    "<p><strong>Battery:</strong> N/A (desktop or not detected)</p>"
}

# Updates badge
$updateCount = $reportData['Updates'].Count
$updateBadge = if ($updateCount -eq 0) {
    "<span class='tk-badge-ok'>Up to date</span>"
} else {
    "<span class='tk-badge-warn'>$updateCount pending</span>"
}
$updatesTable = if ($updateCount -gt 0) {
    ConvertTo-HtmlTable -Objects $reportData['Updates'] -EmptyMessage "No pending updates."
} else {
    "<p class='empty'>System is fully up to date.</p>"
}

# Events badge
$eventCount = $reportData['Events'].Count
$eventBadge = if ($eventCount -eq 0) {
    "<span class='tk-badge-ok'>Clean</span>"
} else {
    "<span class='tk-badge-warn'>$eventCount events</span>"
}

$softwareTable = ConvertTo-HtmlTable -Objects $reportData['Software']        -EmptyMessage "No software found."
$eventsTable   = ConvertTo-HtmlTable -Objects $reportData['Events']          -EmptyMessage "No critical/error events in the last 24 hours."
$tasksTable    = ConvertTo-HtmlTable -Objects $reportData['ScheduledTasks']  -EmptyMessage "No non-Microsoft scheduled tasks found."

$taskCount = $reportData['ScheduledTasks'].Count
$taskBadge = if ($taskCount -eq 0) {
    "<span class='tk-badge-ok'>None</span>"
} else {
    "<span class='tk-badge-info'>$taskCount task(s)</span>"
}

# AV / Security badge
$avUnprotected = ($reportData['Security'] | Where-Object { $_.RealTimeProtection -eq 'Off' }).Count
$avBadge = if ($reportData['Security'].Count -eq 0) {
    "<span class='tk-badge-warn'>No AV detected</span>"
} elseif ($avUnprotected -gt 0) {
    "<span class='tk-badge-err'>$avUnprotected unprotected</span>"
} else {
    "<span class='tk-badge-ok'>Protected</span>"
}
$securityTable = ConvertTo-HtmlTable -Objects $reportData['Security'] -EmptyMessage "No AV products detected."

$tkConfig   = Get-TKConfig
$tkOrgName  = if (-not [string]::IsNullOrWhiteSpace($tkConfig.OrgName)) { EscHtml $tkConfig.OrgName } else { $null }
$tkSubtitle = if ($tkOrgName) { "$tkOrgName  -  $env:COMPUTERNAME" } else { $env:COMPUTERNAME }

$tkMetaItems = [ordered]@{
    'Machine'    = $env:COMPUTERNAME
    'Run As'     = "$env:USERDOMAIN\$env:USERNAME"
    'Generated'  = $reportTimestamp
    'Storage'    = $storageBadge
    'Updates'    = $updateBadge
}

$tkNavItems = @(
    'Hardware',
    'Operating System',
    'Network Configuration',
    'System Health',
    'Storage & RAID Health',
    'Pending Updates',
    'Installed Software',
    'Event Log',
    'Scheduled Tasks',
    'Security & AV'
)

$htmlReport = (Get-TKHtmlHead `
    -Title      'System Diagnostic Report' `
    -ScriptName 'A.U.S.P.E.X.' `
    -Subtitle   $tkSubtitle `
    -MetaItems  $tkMetaItems `
    -NavItems   $tkNavItems) + @"

  <!-- HARDWARE -->
  <div class="tk-section" id="hardware">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 1</span>
        <h2 class="tk-section-title">Hardware</h2>
      </div>
      <div style="padding:20px;">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:20px;">
          <div class="tk-info-box">
            <div><span class="tk-info-label">Manufacturer</span> $($hw.Manufacturer)</div>
            <div><span class="tk-info-label">Model</span> $($hw.Model)</div>
            <div><span class="tk-info-label">Serial Number</span> $($hw.Serial)</div>
          </div>
          <div class="tk-info-box">
            <div><span class="tk-info-label">CPU</span> $($hw.CPU)</div>
            <div><span class="tk-info-label">Cores / Threads</span> $($hw.Cores) / $($hw.Threads)</div>
            <div><span class="tk-info-label">RAM</span> $($hw.RAMGB) GB</div>
          </div>
        </div>
        <table class="tk-table">
          <thead><tr><th>Drive</th><th>Label</th><th>Total</th><th>Used</th><th>Free</th><th>Usage</th></tr></thead>
          <tbody>$diskRows</tbody>
        </table>
      </div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- OPERATING SYSTEM -->
  <div class="tk-section" id="operating-system">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 2</span>
        <h2 class="tk-section-title">Operating System</h2>
      </div>
      <div style="padding:20px;">
        <div class="tk-info-box">
          <div><span class="tk-info-label">OS</span> $($os.Caption)</div>
          <div><span class="tk-info-label">Build</span> $($os.Build)</div>
          <div><span class="tk-info-label">Architecture</span> $($os.Architecture)</div>
          <div><span class="tk-info-label">Install Date</span> $($os.InstallDate)</div>
          <div><span class="tk-info-label">Activation</span> $($os.Activation)</div>
        </div>
      </div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- NETWORK -->
  <div class="tk-section" id="network-configuration">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 3</span>
        <h2 class="tk-section-title">Network Configuration</h2>
      </div>
      <div style="padding:20px;">
        <table class="tk-table">
          <thead><tr><th>Adapter</th><th>IP Address</th><th>MAC</th><th>Gateway</th><th>DNS</th></tr></thead>
          <tbody>$netRows</tbody>
        </table>
      </div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- HEALTH -->
  <div class="tk-section" id="system-health">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 4</span>
        <h2 class="tk-section-title">System Health</h2>
      </div>
      <div style="padding:20px;">
        <div class="tk-info-box">
          <div><span class="tk-info-label">Last Boot</span> $($health.LastBoot)</div>
          <div><span class="tk-info-label">Uptime</span> $($health.Uptime)</div>
          <div><span class="tk-info-label">Battery</span> $(if ($health.Battery) { "$($health.Battery.Charge)% - $($health.Battery.Status)" } else { "N/A" })</div>
        </div>
      </div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- STORAGE & RAID HEALTH -->
  <div class="tk-section" id="storage-raid-health">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 5</span>
        <h2 class="tk-section-title">Storage &amp; RAID Health</h2>
        <span style="margin-left:auto;">$storageBadge</span>
      </div>
      <div style="padding:20px;">
        <p class="tk-section-subtitle" style="margin-bottom:10px;">Physical Disks</p>
        $(if ($physDiskRows) {
          "<table class='tk-table'><thead><tr><th>ID</th><th>Name</th><th>Type</th><th>Bus</th><th>Size</th><th>Health</th><th>Status</th></tr></thead><tbody>$physDiskRows</tbody></table>"
        } else {
          "<div class='tk-info-box'>No physical disk data available.</div>"
        })
        $(if ($virtDiskRows) {
          "<p class='tk-section-subtitle' style='margin:16px 0 10px;'>Storage Spaces (Virtual Disks)</p><table class='tk-table'><thead><tr><th>Name</th><th>Resiliency</th><th>Size</th><th>Health</th><th>Status</th></tr></thead><tbody>$virtDiskRows</tbody></table>"
        } else {
          "<div class='tk-info-box' style='margin-top:12px;'>No Storage Spaces virtual disks detected.</div>"
        })
        <p class="tk-section-subtitle" style="margin:16px 0 10px;">Hardware RAID Controllers</p>
        $(if ($raidCtrlRows) {
          "<table class='tk-table'><thead><tr><th>Name</th><th>Manufacturer</th><th>Status</th><th>Driver</th></tr></thead><tbody>$raidCtrlRows</tbody></table>"
        } else {
          "<div class='tk-info-box'>No dedicated hardware RAID controllers detected.</div>"
        })
        $(if ($raidVendorHtml) {
          "<p class='tk-section-subtitle' style='margin:16px 0 6px;'>Vendor CLI Detail</p>$raidVendorHtml"
        })
      </div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- PENDING UPDATES -->
  <div class="tk-section" id="pending-updates">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 6</span>
        <h2 class="tk-section-title">Pending Windows Updates</h2>
        <span style="margin-left:auto;">$updateBadge</span>
      </div>
      <div style="padding:20px;">$updatesTable</div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- INSTALLED SOFTWARE -->
  <div class="tk-section" id="installed-software">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 7</span>
        <h2 class="tk-section-title">Installed Software</h2>
        <span class="tk-badge-info" style="margin-left:auto;">$($reportData['Software'].Count) apps</span>
      </div>
      <div style="padding:20px;">$softwareTable</div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- EVENT LOG -->
  <div class="tk-section" id="event-log">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 8</span>
        <h2 class="tk-section-title">Event Log  -  Errors &amp; Critical (Last 24h)</h2>
        <span style="margin-left:auto;">$eventBadge</span>
      </div>
      <div style="padding:20px;">$eventsTable</div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- SCHEDULED TASKS -->
  <div class="tk-section" id="scheduled-tasks">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 9</span>
        <h2 class="tk-section-title">Non-Microsoft Scheduled Tasks</h2>
        <span style="margin-left:auto;">$taskBadge</span>
      </div>
      <div style="padding:20px;">$tasksTable</div>
    </div>
  </div>

  <hr class="tk-divider">

  <!-- SECURITY / AV STATUS -->
  <div class="tk-section" id="security-av">
    <div class="tk-card">
      <div class="tk-card-header">
        <span class="tk-section-tag">PART 10</span>
        <h2 class="tk-section-title">Security &amp; Antivirus Status</h2>
        <span style="margin-left:auto;">$avBadge</span>
      </div>
      <div style="padding:20px;">$securityTable</div>
    </div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'A.U.S.P.E.X. v3.0')

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

Write-Host "================================================" -ForegroundColor $ColorSchema.Header
Write-Host "     PROBE COMPLETE" -ForegroundColor $ColorSchema.Header
Write-Host "================================================" -ForegroundColor $ColorSchema.Header
Write-Host ""
Write-Host "  Machine    : $env:COMPUTERNAME" -ForegroundColor $ColorSchema.Info
Write-Host "  OS         : $($reportData['OS'].Caption)" -ForegroundColor $ColorSchema.Info
Write-Host "  RAM        : $($reportData['Hardware'].RAMGB) GB" -ForegroundColor $ColorSchema.Info
Write-Host "  Uptime     : $($reportData['Health'].Uptime)" -ForegroundColor $ColorSchema.Info

$physDiskCount  = $reportData['Storage'].PhysicalDisks.Count
$raidCtrlCount  = $reportData['Storage'].RaidControllers.Count
if ($storageUnhealthy -gt 0) {
    Write-Host "  Storage    : $storageUnhealthy degraded disk(s) detected!" -ForegroundColor $ColorSchema.Error
} else {
    Write-Host "  Storage    : $physDiskCount disk(s)  -  all healthy" -ForegroundColor $ColorSchema.Success
}
if ($raidCtrlCount -gt 0) {
    if ($raidCtrlUnhealthy -gt 0) {
        Write-Host "  RAID       : $raidCtrlUnhealthy controller(s) reporting issues!" -ForegroundColor $ColorSchema.Error
    } else {
        Write-Host "  RAID       : $raidCtrlCount controller(s)  -  all OK" -ForegroundColor $ColorSchema.Success
    }
} else {
    Write-Host "  RAID       : No hardware RAID controllers detected" -ForegroundColor $ColorSchema.Info
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

Write-Host "  Ext. Tasks : $taskCount non-Microsoft scheduled task(s)" -ForegroundColor $ColorSchema.Info

if ($avUnprotected -gt 0) {
    Write-Host "  Security   : $avUnprotected AV product(s) with protection OFF" -ForegroundColor $ColorSchema.Error
} elseif ($reportData['Security'].Count -eq 0) {
    Write-Host "  Security   : No AV detected" -ForegroundColor $ColorSchema.Warning
} else {
    Write-Host "  Security   : Protected" -ForegroundColor $ColorSchema.Success
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
Write-Host "================================================" -ForegroundColor $ColorSchema.Header
Write-Host "  SCRIPT EXECUTION COMPLETED" -ForegroundColor $ColorSchema.Header
Write-Host "================================================" -ForegroundColor $ColorSchema.Header
Write-Host ""
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
