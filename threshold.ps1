<#
.SYNOPSIS
    T.H.R.E.S.H.O.L.D. — Tests Hardware Reliability, Evaluates Storage Health, & Optimizes/Logs Disk data
    Disk & Storage Health Monitor for PowerShell 5.1+

.DESCRIPTION
    T.H.R.E.S.H.O.L.D. is a comprehensive disk and storage health monitoring tool that provides
    physical disk health checks via CIM/WMI, volume space summaries, disk cleanup operations,
    old profile detection, and HTML report generation. It uses Get-PhysicalDisk, Get-Disk,
    Get-Volume, and Win32_DiskDrive to surface health status, operational state, and space usage.
    Volumes below 15% free are flagged as Warning; below 5% as Critical. Full SMART attribute
    reading requires third-party tools such as CrystalDiskInfo — this tool uses HealthStatus
    as a proxy.

.USAGE
    PS C:\> .\threshold.ps1                         # Interactive menu
    PS C:\> .\threshold.ps1 -Unattended             # Run health check and export HTML report silently

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

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Healthy / sufficient space
    Yellow   Warnings / low space
    Red      Critical errors / very low space
    Gray     Information and details
#>

param(
    [switch]$Unattended
)

#region ── Bootstrap ────────────────────────────────────────────────────────────

Import-Module "$PSScriptRoot\TechnicianToolkit.psm1" -Force
Assert-AdminPrivilege

# Script path resolution
$ScriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }

# Color hashtable
$C = @{
    Cyan    = 'Cyan'
    Magenta = 'Magenta'
    Green   = 'Green'
    Yellow  = 'Yellow'
    Red     = 'Red'
    Gray    = 'Gray'
    White   = 'White'
}

#endregion

#region ── Banner ───────────────────────────────────────────────────────────────

function Show-Banner {
    Clear-Host
    $banner = @"

  ████████╗██╗  ██╗██████╗ ███████╗███████╗██╗  ██╗ ██████╗ ██╗     ██████╗
  ╚══██╔══╝██║  ██║██╔══██╗██╔════╝██╔════╝██║  ██║██╔═══██╗██║     ██╔══██╗
     ██║   ███████║██████╔╝█████╗  ███████╗███████║██║   ██║██║     ██║  ██║
     ██║   ██╔══██║██╔══██╗██╔══╝  ╚════██║██╔══██║██║   ██║██║     ██║  ██║
     ██║   ██║  ██║██║  ██║███████╗███████║██║  ██║╚██████╔╝███████╗██████╔╝
     ╚═╝   ╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝╚══════╝╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═════╝
"@
    Write-Host $banner -ForegroundColor $C.Cyan
    Write-Host "  Tests Hardware Reliability, Evaluates Storage Health, & Optimizes/Logs Disk data" -ForegroundColor $C.Gray
    Write-Host "  Disk & Storage Health Monitor  |  v1.0" -ForegroundColor $C.Gray
    Write-Host ""
}

#endregion

#region ── Helpers ──────────────────────────────────────────────────────────────

function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -lt 1KB)  { return "$Bytes B" }
    if ($Bytes -lt 1MB)  { return "{0:N1} KB" -f ($Bytes / 1KB) }
    if ($Bytes -lt 1GB)  { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -lt 1TB)  { return "{0:N2} GB" -f ($Bytes / 1GB) }
    return "{0:N2} TB" -f ($Bytes / 1TB)
}

function Get-HealthColor {
    param([string]$Status)
    switch ($Status) {
        'Healthy'   { return $C.Green }
        'Warning'   { return $C.Yellow }
        'Unhealthy' { return $C.Red }
        default     { return $C.Gray }
    }
}

#endregion

#region ── Data Collection ──────────────────────────────────────────────────────

function Get-PhysicalDiskInfo {
    $physDisks   = Get-PhysicalDisk -ErrorAction SilentlyContinue
    $diskObjects = Get-Disk        -ErrorAction SilentlyContinue
    $cimDisks    = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue

    $results = @()
    foreach ($pd in $physDisks) {
        $diskNum = $null
        # Try to match by serial or DeviceId
        $matchedDisk = $diskObjects | Where-Object { $_.Number -eq ($pd.DeviceId -replace '\D','') } | Select-Object -First 1
        if (-not $matchedDisk) {
            # Fallback: match by size
            $matchedDisk = $diskObjects | Where-Object { $_.Size -eq $pd.Size } | Select-Object -First 1
        }

        $cimMatch = $cimDisks | Where-Object { $_.Size -eq $pd.Size } | Select-Object -First 1

        $results += [PSCustomObject]@{
            FriendlyName      = $pd.FriendlyName
            MediaType         = $pd.MediaType
            Size              = $pd.Size
            SizeFormatted     = Format-Bytes -Bytes ([long]$pd.Size)
            HealthStatus      = if ($pd.HealthStatus) { $pd.HealthStatus } else { 'Unknown' }
            OperationalStatus = if ($pd.OperationalStatus) { $pd.OperationalStatus } else {
                                    if ($cimMatch) { $cimMatch.Status } else { 'Unknown' }
                                }
            BusType           = $pd.BusType
            DiskNumber        = if ($matchedDisk) { $matchedDisk.Number } else { 'N/A' }
            PartitionStyle    = if ($matchedDisk) { $matchedDisk.PartitionStyle } else { 'N/A' }
            IsSystem          = if ($matchedDisk) { $matchedDisk.IsSystem } else { $false }
            IsBoot            = if ($matchedDisk) { $matchedDisk.IsBoot } else { $false }
        }
    }
    return $results
}

function Get-VolumeInfo {
    $volumes = Get-Volume -ErrorAction SilentlyContinue | Where-Object {
        $_.DriveType -ne 'CD-ROM' -and $_.DriveLetter -ne $null -and $_.Size -gt 0
    }

    $results = @()
    foreach ($vol in $volumes) {
        $pctFree = if ($vol.Size -gt 0) { [Math]::Round(($vol.SizeRemaining / $vol.Size) * 100, 1) } else { 0 }
        $spaceStatus = if ($pctFree -lt 5) { 'Critical' } elseif ($pctFree -lt 15) { 'Warning' } else { 'OK' }

        $results += [PSCustomObject]@{
            DriveLetter     = $vol.DriveLetter
            Label           = if ($vol.FileSystemLabel) { $vol.FileSystemLabel } else { '(No Label)' }
            FileSystem      = $vol.FileSystem
            Size            = $vol.Size
            SizeFormatted   = Format-Bytes -Bytes ([long]$vol.Size)
            SizeRemaining   = $vol.SizeRemaining
            FreeFormatted   = Format-Bytes -Bytes ([long]$vol.SizeRemaining)
            PercentFree     = $pctFree
            SpaceStatus     = $spaceStatus
            DriveType       = $vol.DriveType
            HealthStatus    = if ($vol.HealthStatus) { $vol.HealthStatus } else { 'Unknown' }
        }
    }
    return $results
}

#endregion

#region ── Display Functions ────────────────────────────────────────────────────

function Show-DiskSummary {
    Write-Section "PHYSICAL DISK SUMMARY"

    $diskInfo = Get-PhysicalDiskInfo
    if (-not $diskInfo -or $diskInfo.Count -eq 0) {
        Write-Host "  No physical disks detected." -ForegroundColor $C.Yellow
    } else {
        Write-Host ("  {0,-30} {1,-6} {2,-10} {3,-12} {4,-12} {5}" -f `
            "Name", "Num", "Type", "Size", "Health", "OpStatus") -ForegroundColor $C.Gray
        Write-Host ("  " + ("─" * 80)) -ForegroundColor $C.Gray

        foreach ($d in $diskInfo) {
            $hColor = Get-HealthColor -Status $d.HealthStatus
            $flag = if ($d.IsBoot) { "[BOOT]" } elseif ($d.IsSystem) { "[SYS]" } else { "" }
            Write-Host ("  {0,-30} {1,-6} {2,-10} {3,-12}" -f `
                ($d.FriendlyName.Substring(0, [Math]::Min(29, $d.FriendlyName.Length))),
                $d.DiskNumber, $d.MediaType, $d.SizeFormatted) -NoNewline -ForegroundColor $C.White
            Write-Host (" {0,-12}" -f $d.HealthStatus) -NoNewline -ForegroundColor $hColor
            Write-Host (" {0} {1}" -f $d.OperationalStatus, $flag) -ForegroundColor $C.Gray
        }
    }

    Write-Section "VOLUME SPACE SUMMARY"

    $volInfo = Get-VolumeInfo
    if (-not $volInfo -or $volInfo.Count -eq 0) {
        Write-Host "  No volumes detected." -ForegroundColor $C.Yellow
    } else {
        Write-Host ("  {0,-5} {1,-20} {2,-8} {3,-10} {4,-10} {5,-8} {6}" -f `
            "Drive", "Label", "FS", "Total", "Free", "% Free", "Status") -ForegroundColor $C.Gray
        Write-Host ("  " + ("─" * 75)) -ForegroundColor $C.Gray

        foreach ($v in $volInfo) {
            $sColor = switch ($v.SpaceStatus) {
                'Critical' { $C.Red }
                'Warning'  { $C.Yellow }
                default    { $C.Green }
            }
            $bar = if ($v.PercentFree -ge 100) { 100 } else { [int]$v.PercentFree }
            $barStr = "[" + ("#" * [Math]::Round($bar / 5)) + ("." * (20 - [Math]::Round($bar / 5))) + "]"

            Write-Host ("  {0,-5} {1,-20} {2,-8} {3,-10} {4,-10}" -f `
                "$($v.DriveLetter):", $v.Label.Substring(0, [Math]::Min(19, $v.Label.Length)),
                $v.FileSystem, $v.SizeFormatted, $v.FreeFormatted) -NoNewline -ForegroundColor $C.White
            Write-Host (" {0,-6}% {1}" -f $v.PercentFree, $barStr) -ForegroundColor $sColor
        }
    }

    # Totals
    if ($volInfo) {
        $totalSize = ($volInfo | Measure-Object -Property Size -Sum).Sum
        $totalFree = ($volInfo | Measure-Object -Property SizeRemaining -Sum).Sum
        Write-Host ""
        Write-Host ("  Total Storage : {0}" -f (Format-Bytes -Bytes ([long]$totalSize))) -ForegroundColor $C.Cyan
        Write-Host ("  Total Free    : {0}" -f (Format-Bytes -Bytes ([long]$totalFree))) -ForegroundColor $C.Cyan
    }
}

function Show-DiskHealth {
    Write-Section "DISK HEALTH STATUS"

    $diskInfo = Get-PhysicalDiskInfo
    $cimDisks = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue

    Write-Host "  Physical Disk Health (via Get-PhysicalDisk):" -ForegroundColor $C.Cyan
    Write-Host ""

    foreach ($d in $diskInfo) {
        $hColor = Get-HealthColor -Status $d.HealthStatus
        Write-Host "  Disk $($d.DiskNumber): $($d.FriendlyName)" -ForegroundColor $C.White
        Write-Host ("    Media Type        : {0}" -f $d.MediaType) -ForegroundColor $C.Gray
        Write-Host ("    Size              : {0}" -f $d.SizeFormatted) -ForegroundColor $C.Gray
        Write-Host ("    Bus Type          : {0}" -f $d.BusType) -ForegroundColor $C.Gray
        Write-Host ("    Partition Style   : {0}" -f $d.PartitionStyle) -ForegroundColor $C.Gray
        Write-Host "    Health Status     : " -NoNewline -ForegroundColor $C.Gray
        Write-Host $d.HealthStatus -ForegroundColor $hColor
        Write-Host ("    Operational Status: {0}" -f $d.OperationalStatus) -ForegroundColor $C.Gray
        Write-Host ""
    }

    Write-Host "  WMI Win32_DiskDrive Operational Status:" -ForegroundColor $C.Cyan
    Write-Host ""
    foreach ($cd in $cimDisks) {
        $statusColor = if ($cd.Status -eq 'OK') { $C.Green } else { $C.Yellow }
        Write-Host ("  {0}" -f $cd.Caption) -ForegroundColor $C.White
        Write-Host ("    Status       : {0}" -f $cd.Status) -NoNewline -ForegroundColor $C.Gray
        Write-Host "" -ForegroundColor $statusColor
        Write-Host ("    Availability : {0}" -f $cd.Availability) -ForegroundColor $C.Gray
        Write-Host ("    Interface    : {0}" -f $cd.InterfaceType) -ForegroundColor $C.Gray
        Write-Host ("    Sectors/Track: {0}" -f $cd.SectorsPerTrack) -ForegroundColor $C.Gray
        Write-Host ""
    }

    Write-Host "  NOTE: Full SMART attribute data requires a third-party tool such as" -ForegroundColor $C.Yellow
    Write-Host "        CrystalDiskInfo. HealthStatus above is a proxy indicator only." -ForegroundColor $C.Yellow
}

#endregion

#region ── Cleanup Functions ────────────────────────────────────────────────────

function Invoke-TempCleanup {
    $tempPaths = @($env:TEMP, "C:\Windows\Temp")
    $totalDeleted = 0
    $totalBytes   = 0L

    foreach ($path in $tempPaths) {
        if (-not (Test-Path $path)) { continue }

        Write-Host "  Scanning: $path" -ForegroundColor $C.Magenta
        $files = Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue |
                 Where-Object { -not $_.PSIsContainer }

        $beforeBytes = ($files | Measure-Object -Property Length -Sum).Sum
        Write-Host ("    Found {0} files ({1})" -f $files.Count, (Format-Bytes -Bytes ([long]$beforeBytes))) -ForegroundColor $C.Gray

        $deleted = 0
        $bytesFreed = 0L
        foreach ($f in $files) {
            try {
                $bytesFreed += $f.Length
                Remove-Item -Path $f.FullName -Force -ErrorAction Stop
                $deleted++
            } catch {
                # Skip locked/protected files
            }
        }

        # Also remove empty subdirectories
        Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.PSIsContainer } |
            Sort-Object FullName -Descending |
            ForEach-Object { Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue }

        Write-Host ("    Deleted {0} files, freed {1}" -f $deleted, (Format-Bytes -Bytes $bytesFreed)) -ForegroundColor $C.Green
        $totalDeleted += $deleted
        $totalBytes   += $bytesFreed
    }

    return [PSCustomObject]@{
        FilesDeleted = $totalDeleted
        BytesFreed   = $totalBytes
    }
}

function Invoke-WUCacheCleanup {
    $wuPath = "C:\Windows\SoftwareDistribution\Download"

    if (-not (Test-Path $wuPath)) {
        Write-Host "  WU cache path not found: $wuPath" -ForegroundColor $C.Yellow
        return
    }

    $beforeSize = (Get-ChildItem -Path $wuPath -Recurse -Force -ErrorAction SilentlyContinue |
                   Measure-Object -Property Length -Sum).Sum

    Write-Host ("  Before: {0}" -f (Format-Bytes -Bytes ([long]$beforeSize))) -ForegroundColor $C.Gray
    Write-Host "  Stopping Windows Update service (wuauserv)..." -ForegroundColor $C.Magenta

    Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    Write-Host "  Removing cached downloads..." -ForegroundColor $C.Magenta
    Get-ChildItem -Path $wuPath -Force -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

    $afterSize = (Get-ChildItem -Path $wuPath -Recurse -Force -ErrorAction SilentlyContinue |
                  Measure-Object -Property Length -Sum).Sum

    Write-Host "  Starting Windows Update service..." -ForegroundColor $C.Magenta
    Start-Service -Name wuauserv -ErrorAction SilentlyContinue

    $freed = [long]$beforeSize - [long]$afterSize
    Write-Host ("  After : {0}  |  Freed: {1}" -f (Format-Bytes -Bytes ([long]$afterSize)), (Format-Bytes -Bytes $freed)) -ForegroundColor $C.Green
}

function Invoke-RecycleBinCleanup {
    Write-Host "  Emptying Recycle Bin..." -ForegroundColor $C.Magenta
    try {
        Clear-RecycleBin -Force -ErrorAction Stop
        Write-Host "  Recycle Bin emptied successfully." -ForegroundColor $C.Green
    } catch {
        Write-Host ("  Warning: {0}" -f $_.Exception.Message) -ForegroundColor $C.Yellow
    }
}

function Show-CleanupMenu {
    Write-Section "DISK CLEANUP"

    Write-Host ""
    Write-Host "  Select cleanup operation:" -ForegroundColor $C.White
    Write-Host "    A.  Clear Temp Files (%TEMP% and C:\Windows\Temp)" -ForegroundColor $C.Gray
    Write-Host "    B.  Clear Windows Update Cache" -ForegroundColor $C.Gray
    Write-Host "    C.  Empty Recycle Bin" -ForegroundColor $C.Gray
    Write-Host "    D.  All of the above" -ForegroundColor $C.Cyan
    Write-Host "    X.  Back to main menu" -ForegroundColor $C.Gray
    Write-Host ""

    $sub = Read-Host "  Choice"
    switch ($sub.ToUpper()) {
        'A' {
            $confirm = Read-Host "  Delete temp files? This cannot be undone. [Y/N]"
            if ($confirm -match '^[Yy]') {
                $result = Invoke-TempCleanup
                Write-Host ""
                Write-Host ("  Total: {0} files deleted, {1} freed." -f $result.FilesDeleted, (Format-Bytes -Bytes $result.BytesFreed)) -ForegroundColor $C.Green
            } else { Write-Host "  Cancelled." -ForegroundColor $C.Gray }
        }
        'B' {
            $confirm = Read-Host "  Clear Windows Update cache? [Y/N]"
            if ($confirm -match '^[Yy]') {
                Invoke-WUCacheCleanup
            } else { Write-Host "  Cancelled." -ForegroundColor $C.Gray }
        }
        'C' {
            $confirm = Read-Host "  Empty Recycle Bin? [Y/N]"
            if ($confirm -match '^[Yy]') {
                Invoke-RecycleBinCleanup
            } else { Write-Host "  Cancelled." -ForegroundColor $C.Gray }
        }
        'D' {
            $confirm = Read-Host "  Run ALL cleanup operations? [Y/N]"
            if ($confirm -match '^[Yy]') {
                Write-Host ""
                Write-Host "  [1/3] Temp Files" -ForegroundColor $C.Cyan
                $result = Invoke-TempCleanup
                Write-Host ("  Total temp: {0} files, {1} freed." -f $result.FilesDeleted, (Format-Bytes -Bytes $result.BytesFreed)) -ForegroundColor $C.Green

                Write-Host ""
                Write-Host "  [2/3] Windows Update Cache" -ForegroundColor $C.Cyan
                Invoke-WUCacheCleanup

                Write-Host ""
                Write-Host "  [3/3] Recycle Bin" -ForegroundColor $C.Cyan
                Invoke-RecycleBinCleanup
            } else { Write-Host "  Cancelled." -ForegroundColor $C.Gray }
        }
        'X' { return }
        default { Write-Host "  Invalid selection." -ForegroundColor $C.Yellow }
    }
}

#endregion

#region ── Old Profile Detection ────────────────────────────────────────────────

function Show-OldProfiles {
    Write-Section "OLD USER PROFILES (> 90 DAYS)"

    $usersRoot = "C:\Users"
    $cutoff    = (Get-Date).AddDays(-90)

    if (-not (Test-Path $usersRoot)) {
        Write-Host "  C:\Users not found." -ForegroundColor $C.Yellow
        return
    }

    $excluded = @('Public', 'Default', 'Default User', 'All Users', 'defaultuser0')
    $profiles = Get-ChildItem -Path $usersRoot -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -notin $excluded -and $_.LastWriteTime -lt $cutoff }

    if (-not $profiles -or $profiles.Count -eq 0) {
        Write-Host "  No user profiles older than 90 days found." -ForegroundColor $C.Green
        return
    }

    Write-Host ("  {0,-20} {1,-22} {2,-12} {3}" -f "Profile", "Last Modified", "Est. Size", "Path") -ForegroundColor $C.Gray
    Write-Host ("  " + ("─" * 75)) -ForegroundColor $C.Gray

    foreach ($prof in $profiles) {
        $sizeStr = "Calculating..."
        try {
            $sizeBytes = (Get-ChildItem -Path $prof.FullName -Recurse -Force -ErrorAction SilentlyContinue |
                          Measure-Object -Property Length -Sum).Sum
            $sizeStr = Format-Bytes -Bytes ([long]$sizeBytes)
        } catch {
            $sizeStr = "Access Denied"
        }

        $age = [Math]::Round(((Get-Date) - $prof.LastWriteTime).TotalDays)
        $nameStr = $prof.Name.Substring(0, [Math]::Min(19, $prof.Name.Length))

        Write-Host ("  {0,-20} {1,-22} {2,-12} {3}" -f `
            $nameStr,
            ($prof.LastWriteTime.ToString("yyyy-MM-dd") + " ($age days ago)"),
            $sizeStr,
            $prof.FullName) -ForegroundColor $C.Yellow
    }

    Write-Host ""
    Write-Host ("  {0} profile(s) found older than 90 days." -f $profiles.Count) -ForegroundColor $C.Cyan
    Write-Host "  Review before deletion — use ARCHIVE or PHANTOM for profile migration." -ForegroundColor $C.Gray
}

#endregion

#region ── HTML Report ──────────────────────────────────────────────────────────

function Build-HtmlReport {
    param(
        [array]$DiskData,
        [array]$VolumeData
    )

    $timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $hostname   = $env:COMPUTERNAME
    $filename   = "THRESHOLD_{0}.html" -f (Get-Date -Format "yyyyMMdd_HHmmss")
    $outputPath = Join-Path $ScriptPath $filename

    $totalDrives   = $DiskData.Count
    $healthyCount  = ($DiskData | Where-Object { $_.HealthStatus -eq 'Healthy' }).Count
    $warnCount     = ($DiskData | Where-Object { $_.HealthStatus -notin @('Healthy', 'Unknown') }).Count
    $totalStorage  = ($VolumeData | Measure-Object -Property Size -Sum).Sum
    $totalFree     = ($VolumeData | Measure-Object -Property SizeRemaining -Sum).Sum

    # Build disk rows
    $diskRows = ""
    foreach ($d in $DiskData) {
        $badgeColor = switch ($d.HealthStatus) {
            'Healthy'   { '#00c853' }
            'Warning'   { '#ffd600' }
            'Unhealthy' { '#d50000' }
            default     { '#757575' }
        }
        $bootTag = if ($d.IsBoot) { ' <span style="font-size:10px;background:#1565c0;padding:1px 5px;border-radius:3px;">BOOT</span>' } else { '' }
        $diskRows += @"
            <tr>
                <td>$($d.FriendlyName)$bootTag</td>
                <td>$($d.MediaType)</td>
                <td>$($d.SizeFormatted)</td>
                <td><span style="background:$badgeColor;color:#000;padding:2px 10px;border-radius:12px;font-weight:bold;font-size:12px;">$($d.HealthStatus)</span></td>
                <td>$($d.OperationalStatus)</td>
                <td>$($d.BusType)</td>
                <td>$($d.PartitionStyle)</td>
            </tr>
"@
    }

    # Build volume rows
    $volRows = ""
    foreach ($v in $VolumeData) {
        $rowBg = switch ($v.SpaceStatus) {
            'Critical' { 'rgba(213,0,0,0.15)' }
            'Warning'  { 'rgba(255,214,0,0.10)' }
            default    { 'transparent' }
        }
        $barColor = switch ($v.SpaceStatus) {
            'Critical' { '#d50000' }
            'Warning'  { '#ffd600' }
            default    { '#00c853' }
        }
        $barWidth = [Math]::Min(100, [Math]::Max(1, [int]$v.PercentFree))
        $statusBadgeColor = switch ($v.SpaceStatus) {
            'Critical' { '#d50000' }
            'Warning'  { '#ffd600' }
            default    { '#00c853' }
        }

        $volRows += @"
            <tr style="background:$rowBg;">
                <td style="font-weight:bold;">$($v.DriveLetter):</td>
                <td>$($v.Label)</td>
                <td>$($v.FileSystem)</td>
                <td>$($v.SizeFormatted)</td>
                <td>$($v.FreeFormatted)</td>
                <td>
                    <div style="display:flex;align-items:center;gap:8px;">
                        <div style="flex:1;background:#2a2a4a;border-radius:4px;height:12px;min-width:80px;">
                            <div style="width:$($barWidth)%;background:$barColor;height:12px;border-radius:4px;"></div>
                        </div>
                        <span style="min-width:40px;font-size:12px;">$($v.PercentFree)%</span>
                    </div>
                </td>
                <td><span style="background:$statusBadgeColor;color:#000;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:bold;">$($v.SpaceStatus)</span></td>
                <td>$($v.HealthStatus)</td>
            </tr>
"@
    }

    # Recommendations
    $recommendations = ""
    $critVols = $VolumeData | Where-Object { $_.SpaceStatus -eq 'Critical' }
    $warnVols = $VolumeData | Where-Object { $_.SpaceStatus -eq 'Warning' }
    $badDisks = $DiskData   | Where-Object { $_.HealthStatus -notin @('Healthy', 'Unknown') }

    if ($critVols) {
        foreach ($v in $critVols) {
            $recommendations += "<li class='rec-critical'>Drive <strong>$($v.DriveLetter):</strong> ($($v.Label)) is critically low on space ($($v.PercentFree)% free). Immediate action required.</li>`n"
        }
    }
    if ($warnVols) {
        foreach ($v in $warnVols) {
            $recommendations += "<li class='rec-warning'>Drive <strong>$($v.DriveLetter):</strong> ($($v.Label)) has low free space ($($v.PercentFree)% free). Consider cleanup.</li>`n"
        }
    }
    if ($badDisks) {
        foreach ($d in $badDisks) {
            $recommendations += "<li class='rec-critical'>Physical disk <strong>$($d.FriendlyName)</strong> reports Health Status: <strong>$($d.HealthStatus)</strong>. Back up data immediately.</li>`n"
        }
    }
    if (-not $recommendations) {
        $recommendations = "<li class='rec-ok'>All monitored disks and volumes are within healthy thresholds.</li>"
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>T.H.R.E.S.H.O.L.D. Report — $hostname</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Segoe UI', Consolas, monospace;
            background: #1a1a2e;
            color: #e0e0e0;
            padding: 24px;
            min-height: 100vh;
        }
        .header {
            border-bottom: 2px solid #00d4ff;
            padding-bottom: 16px;
            margin-bottom: 24px;
        }
        .header h1 {
            font-size: 28px;
            color: #00d4ff;
            letter-spacing: 4px;
            font-weight: 700;
        }
        .header .subtitle {
            color: #9e9e9e;
            font-size: 13px;
            margin-top: 4px;
        }
        .header .meta {
            color: #757575;
            font-size: 12px;
            margin-top: 8px;
        }
        .cards {
            display: flex;
            flex-wrap: wrap;
            gap: 16px;
            margin-bottom: 28px;
        }
        .card {
            background: #16213e;
            border: 1px solid #0f3460;
            border-radius: 8px;
            padding: 16px 24px;
            min-width: 150px;
            flex: 1;
        }
        .card .card-label {
            font-size: 11px;
            color: #757575;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        .card .card-value {
            font-size: 26px;
            font-weight: 700;
            color: #00d4ff;
            margin-top: 4px;
        }
        .card .card-sub {
            font-size: 11px;
            color: #9e9e9e;
            margin-top: 2px;
        }
        .section {
            margin-bottom: 28px;
        }
        .section-title {
            font-size: 14px;
            font-weight: 700;
            color: #00d4ff;
            text-transform: uppercase;
            letter-spacing: 2px;
            border-left: 3px solid #00d4ff;
            padding-left: 10px;
            margin-bottom: 12px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }
        thead th {
            background: #0f3460;
            color: #00d4ff;
            padding: 10px 12px;
            text-align: left;
            font-weight: 600;
            letter-spacing: 0.5px;
        }
        tbody tr {
            border-bottom: 1px solid #1e2a4a;
        }
        tbody tr:hover {
            background: rgba(0, 212, 255, 0.05);
        }
        tbody td {
            padding: 10px 12px;
            color: #e0e0e0;
        }
        .rec-list {
            list-style: none;
            padding: 0;
        }
        .rec-list li {
            padding: 10px 14px;
            border-radius: 6px;
            margin-bottom: 8px;
            font-size: 13px;
            border-left: 4px solid;
        }
        .rec-critical {
            background: rgba(213,0,0,0.15);
            border-left-color: #d50000;
            color: #ff8a80;
        }
        .rec-warning {
            background: rgba(255,214,0,0.10);
            border-left-color: #ffd600;
            color: #fff176;
        }
        .rec-ok {
            background: rgba(0,200,83,0.10);
            border-left-color: #00c853;
            color: #b9f6ca;
        }
        .note-box {
            background: #16213e;
            border: 1px solid #0f3460;
            border-radius: 6px;
            padding: 12px 16px;
            font-size: 12px;
            color: #9e9e9e;
            margin-top: 16px;
        }
        .note-box strong { color: #ffd600; }
        .footer {
            margin-top: 32px;
            border-top: 1px solid #0f3460;
            padding-top: 12px;
            font-size: 11px;
            color: #424242;
            text-align: center;
        }
    </style>
</head>
<body>

<div class="header">
    <h1>T.H.R.E.S.H.O.L.D.</h1>
    <div class="subtitle">Tests Hardware Reliability, Evaluates Storage Health, &amp; Optimizes/Logs Disk data</div>
    <div class="meta">Host: $hostname &nbsp;|&nbsp; Generated: $timestamp &nbsp;|&nbsp; Disk &amp; Storage Health Monitor v1.0</div>
</div>

<div class="cards">
    <div class="card">
        <div class="card-label">Total Drives</div>
        <div class="card-value">$totalDrives</div>
        <div class="card-sub">Physical disks</div>
    </div>
    <div class="card">
        <div class="card-label">Healthy</div>
        <div class="card-value" style="color:#00c853;">$healthyCount</div>
        <div class="card-sub">of $totalDrives disks</div>
    </div>
    <div class="card">
        <div class="card-label">Warning / Critical</div>
        <div class="card-value" style="color:$(if($warnCount -gt 0){'#ffd600'}else{'#00c853'});">$warnCount</div>
        <div class="card-sub">disks need attention</div>
    </div>
    <div class="card">
        <div class="card-label">Total Storage</div>
        <div class="card-value">$(Format-Bytes -Bytes ([long]$totalStorage))</div>
        <div class="card-sub">across all volumes</div>
    </div>
    <div class="card">
        <div class="card-label">Free Storage</div>
        <div class="card-value">$(Format-Bytes -Bytes ([long]$totalFree))</div>
        <div class="card-sub">available across volumes</div>
    </div>
</div>

<div class="section">
    <div class="section-title">Physical Disks</div>
    <table>
        <thead>
            <tr>
                <th>Name</th>
                <th>Type</th>
                <th>Size</th>
                <th>Health Status</th>
                <th>Operational Status</th>
                <th>Bus Type</th>
                <th>Partition Style</th>
            </tr>
        </thead>
        <tbody>
$diskRows
        </tbody>
    </table>
    <div class="note-box">
        <strong>Note:</strong> Health Status is sourced from <code>Get-PhysicalDisk</code> and serves as a proxy for SMART data.
        For full SMART attribute analysis (reallocated sectors, spin retries, uncorrectable errors, etc.),
        use a dedicated tool such as <strong>CrystalDiskInfo</strong> or <strong>smartctl</strong>.
    </div>
</div>

<div class="section">
    <div class="section-title">Volume Space</div>
    <table>
        <thead>
            <tr>
                <th>Drive</th>
                <th>Label</th>
                <th>File System</th>
                <th>Total Size</th>
                <th>Free Space</th>
                <th>% Free</th>
                <th>Space Status</th>
                <th>Health</th>
            </tr>
        </thead>
        <tbody>
$volRows
        </tbody>
    </table>
</div>

<div class="section">
    <div class="section-title">Recommendations</div>
    <ul class="rec-list">
$recommendations
    </ul>
</div>

<div class="footer">
    T.H.R.E.S.H.O.L.D. — Technician Toolkit &nbsp;|&nbsp; Disk &amp; Storage Health Monitor &nbsp;|&nbsp; $timestamp
</div>

</body>
</html>
"@

    $html | Out-File -FilePath $outputPath -Encoding UTF8 -Force
    return $outputPath
}

#endregion

#region ── Main Loop ────────────────────────────────────────────────────────────

function Show-Menu {
    Show-Banner
    Write-Host "  MAIN MENU" -ForegroundColor $C.Cyan
    Write-Host ""
    Write-Host "    1.  Show Disk & Volume Summary" -ForegroundColor $C.White
    Write-Host "    2.  Check Disk Health Status" -ForegroundColor $C.White
    Write-Host "    3.  Run Disk Cleanup" -ForegroundColor $C.White
    Write-Host "    4.  Detect Large Old Profiles (> 90 days)" -ForegroundColor $C.White
    Write-Host "    5.  Export HTML Health Report" -ForegroundColor $C.White
    Write-Host ""
    Write-Host "    Q.  Quit" -ForegroundColor $C.Gray
    Write-Host ""
}

if ($Unattended) {
    Show-Banner
    Write-Host "  Running in Unattended mode — collecting data and generating report..." -ForegroundColor $C.Magenta
    Write-Host ""

    Write-Host "  Collecting physical disk data..." -ForegroundColor $C.Gray
    $diskData = Get-PhysicalDiskInfo

    Write-Host "  Collecting volume data..." -ForegroundColor $C.Gray
    $volData  = Get-VolumeInfo

    Write-Host "  Building HTML report..." -ForegroundColor $C.Gray
    $reportPath = Build-HtmlReport -DiskData $diskData -VolumeData $volData

    Write-Host ""
    Write-Host ("  Report saved: {0}" -f $reportPath) -ForegroundColor $C.Green
    Write-Host "  Unattended run complete." -ForegroundColor $C.Cyan
    exit 0
}

# Interactive loop
do {
    Show-Menu
    $choice = Read-Host "  Select an option"

    switch ($choice.ToUpper()) {
        '1' {
            Show-DiskSummary
            Write-Host ""
            Read-Host "  Press Enter to return to menu"
        }
        '2' {
            Show-DiskHealth
            Write-Host ""
            Read-Host "  Press Enter to return to menu"
        }
        '3' {
            Show-CleanupMenu
            Write-Host ""
            Read-Host "  Press Enter to return to menu"
        }
        '4' {
            Show-OldProfiles
            Write-Host ""
            Read-Host "  Press Enter to return to menu"
        }
        '5' {
            Write-Section "EXPORTING HTML REPORT"
            Write-Host "  Collecting disk data..." -ForegroundColor $C.Magenta
            $diskData = Get-PhysicalDiskInfo

            Write-Host "  Collecting volume data..." -ForegroundColor $C.Magenta
            $volData  = Get-VolumeInfo

            Write-Host "  Building report..." -ForegroundColor $C.Magenta
            $reportPath = Build-HtmlReport -DiskData $diskData -VolumeData $volData

            Write-Host ""
            Write-Host ("  Report saved to: {0}" -f $reportPath) -ForegroundColor $C.Green

            $open = Read-Host "  Open report in browser? [Y/N]"
            if ($open -match '^[Yy]') {
                Start-Process $reportPath
            }
            Write-Host ""
            Read-Host "  Press Enter to return to menu"
        }
        'Q' {
            Show-Banner
            Write-Host "  Exiting T.H.R.E.S.H.O.L.D. — Disk & Storage Health Monitor." -ForegroundColor $C.Cyan
            Write-Host ""
        }
        default {
            Write-Host ""
            Write-Host "  Invalid selection. Please choose 1–5 or Q." -ForegroundColor $C.Yellow
            Start-Sleep -Seconds 1
        }
    }
} while ($choice.ToUpper() -ne 'Q')

# Self-removal
Remove-Item $PSCommandPath -Force -ErrorAction SilentlyContinue

#endregion
