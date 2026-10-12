# leyline.ps1 - L.E.Y.L.I.N.E. — Locates, Examines & Yields Latency, Infrastructure, Network & Endpoints
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
    L.E.Y.L.I.N.E. — Locates, Examines & Yields Latency, Infrastructure, Network & Endpoints
    Network Diagnostics & Remediation Tool for PowerShell 5.1+

.DESCRIPTION
    Tests and diagnoses network connectivity at every layer — adapter status,
    gateway reachability, DNS resolution, internet connectivity, and port
    availability. Offers one-click remediation: flush DNS, release/renew IP,
    and reset the network stack.

.USAGE
    PS C:\> .\leyline.ps1                                          # Interactive menu
    PS C:\> .\leyline.ps1 -Unattended -Action Status               # Adapter + ping summary
    PS C:\> .\leyline.ps1 -Unattended -Action FlushDNS             # Flush DNS cache
    PS C:\> .\leyline.ps1 -Unattended -Action Renew                # Release & renew IP
    PS C:\> .\leyline.ps1 -Unattended -Action PortTest -Target "8.8.8.8:53"

.NOTES
    Version : 6.0

#>

param(
    [switch]$Unattended,
    [ValidateSet('Status','FlushDNS','Renew','ResetStack','PortTest','Trace','ARPScan')]
    [string]$Action = 'Status',
    [ValidatePattern('^$|^[A-Za-z0-9][A-Za-z0-9._-]{0,252}(:[0-9]{1,5})?$|^[A-Za-z0-9]$')]
    [string]$Target = '',
    [switch]$Transcript
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

# ===========================
# SHARED HELPERS
# ===========================
#region TK-SHARED
# Generated from shared/TechnicianToolkit.Shared.ps1 by build/Sync-TKShared.ps1. Do not edit
# this block: edit the source and re-run the script. The Pester suite fails on a stale block.

# Forcing UTF-8 keeps the box-drawing banners intact when a console is attached.
# When one is not — a GUI host, a WinRM runspace, a scheduled task — the setter
# throws "The handle is invalid", so this must never be fatal. Every tool runs
# this as the first line of its generated block.
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host ("  " + ("-" * 62)) -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host ("  " + ("-" * 62)) -ForegroundColor Cyan
    Write-Host ""
}

function Write-Warn { param([string]$Msg) Write-Host ("  [!] {0}" -f $Msg) -ForegroundColor Yellow  }

function Get-TKConfig {
    <#
    .SYNOPSIS
        Returns the toolkit configuration as a PSCustomObject.
        Reads config.json from the tool's own directory; missing keys are filled
        with empty-string defaults so callers never receive null.
    #>
    $configPath = Join-Path $PSScriptRoot 'config.json'

    $defaults = [PSCustomObject]@{
        OrgName       = ''
        LogDirectory  = ''
        Archive       = [PSCustomObject]@{ DefaultDestination   = '' }
        Revenant      = [PSCustomObject]@{ DefaultDestination   = '' }
        Covenant      = [PSCustomObject]@{ DefaultTimezone = ''; DefaultLocalAdminUser = '' }
        # DirectDownloads is an array of { Name; Url; Args; Sha256 } — installers
        # CONJURE fetches by URL because they have no winget/Chocolatey package.
        Conjure       = [PSCustomObject]@{ DirectDownloads = @() }
    }

    if (-not (Test-Path $configPath)) { return $defaults }

    try {
        $raw = Get-Content $configPath -Raw -ErrorAction Stop | ConvertFrom-Json

        # Legacy migration: v2.x used a 'Phantom' section; v3.0 renamed it to 'Revenant'.
        # Copy any populated 'Phantom' value forward so upgrading users don't lose their default destination.
        if ($raw.PSObject.Properties['Phantom'] -and
            $raw.Phantom -is [PSCustomObject] -and
            -not [string]::IsNullOrWhiteSpace($raw.Phantom.DefaultDestination)) {
            if ($raw.PSObject.Properties['Revenant'] -isnot [PSCustomObject] -or
                [string]::IsNullOrWhiteSpace($raw.Revenant.DefaultDestination)) {
                if (-not $raw.PSObject.Properties['Revenant']) {
                    $raw | Add-Member -NotePropertyName 'Revenant' -NotePropertyValue ([PSCustomObject]@{ DefaultDestination = '' }) -Force
                }
                $raw.Revenant.DefaultDestination = $raw.Phantom.DefaultDestination
            }
        }

        # Ensure top-level keys exist
        foreach ($key in ($defaults | Get-Member -MemberType NoteProperty).Name) {
            if ($null -eq $raw.$key) {
                $raw | Add-Member -NotePropertyName $key -NotePropertyValue $defaults.$key -Force
            }
        }
        # Ensure nested keys exist
        foreach ($section in @('Archive','Revenant','Covenant','Conjure')) {
            if ($raw.$section -isnot [PSCustomObject]) {
                $raw | Add-Member -NotePropertyName $section -NotePropertyValue $defaults.$section -Force
            } else {
                foreach ($key in ($defaults.$section | Get-Member -MemberType NoteProperty).Name) {
                    if ($null -eq $raw.$section.$key) {
                        # Carry the default's own type: filling an absent array key
                        # with '' would hand the caller a string to iterate.
                        $fallback = if ($defaults.$section.$key -is [Array]) { @() } else { '' }
                        $raw.$section | Add-Member -NotePropertyName $key -NotePropertyValue $fallback -Force
                    }
                }
            }
        }
        return $raw
    }
    catch { return $defaults }
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

function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Assert-AdminPrivilege {
    if (-not (Test-IsAdmin)) {
        Write-Host "  This script must be run as Administrator." -ForegroundColor Red
        exit 1
    }
}
#endregion TK-SHARED
Assert-AdminPrivilege

if ($Transcript) { Start-TKTranscript -LogRoot (Resolve-LogDirectory -FallbackPath $PSScriptRoot) }

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
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-LeylineBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██╗     ███████╗██╗   ██╗██╗     ██╗███╗   ██╗███████╗
  ██║     ██╔════╝╚██╗ ██╔╝██║     ██║████╗  ██║██╔════╝
  ██║     █████╗   ╚████╔╝ ██║     ██║██╔██╗ ██║█████╗
  ██║     ██╔══╝    ╚██╔╝  ██║     ██║██║╚██╗██║██╔══╝
  ███████╗███████╗   ██║   ███████╗██║██║ ╚████║███████╗
  ╚══════╝╚══════╝   ╚═╝   ╚══════╝╚═╝╚═╝  ╚═══╝╚══════╝

"@ -ForegroundColor Cyan
    Write-Host "    L.E.Y.L.I.N.E. — Locates, Examines & Yields Latency, Infrastructure, Network & Endpoints" -ForegroundColor Cyan
    Write-Host "    Network Diagnostics & Remediation Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function Test-TCPPort {
    param([string]$Hostname, [int]$Port, [int]$TimeoutMs = 2000)
    try {
        $tcp = [System.Net.Sockets.TcpClient]::new()
        $ar  = $tcp.BeginConnect($Hostname, $Port, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        $tcp.Close()
        return $ok
    } catch { return $false }
}

# ─────────────────────────────────────────────────────────────────────────────
# ADAPTER STATUS
# ─────────────────────────────────────────────────────────────────────────────

function Show-AdapterStatus {
    Write-Section "NETWORK ADAPTERS"

    $adapters = Get-NetAdapter | Sort-Object -Property Status -Descending

    if (-not $adapters) {
        Write-Host "  [-] No network adapters found." -ForegroundColor $C.Error
        return
    }

    foreach ($a in $adapters) {
        $statusColor = if ($a.Status -eq 'Up') { $C.Success } else { $C.Warning }
        $ipInfo = Get-NetIPAddress -InterfaceIndex $a.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                  Select-Object -First 1

        Write-Host ("  {0,-32} [{1}]" -f $a.Name, $a.Status) -ForegroundColor $statusColor
        Write-Host ("    MAC  : {0}" -f $a.MacAddress) -ForegroundColor $C.Info
        if ($ipInfo) {
            Write-Host ("    IPv4 : {0}/{1}" -f $ipInfo.IPAddress, $ipInfo.PrefixLength) -ForegroundColor $C.Info
        }
        Write-Host ""
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# CONNECTIVITY PING TESTS
# ─────────────────────────────────────────────────────────────────────────────

function Show-ConnectivityTests {
    Write-Section "CONNECTIVITY TESTS"

    $gateway = (Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
                Sort-Object RouteMetric | Select-Object -First 1).NextHop

    $targets = [ordered]@{}
    if ($gateway) { $targets["Gateway ($gateway)"] = $gateway }
    $targets["Google DNS (8.8.8.8)"]    = "8.8.8.8"
    $targets["Cloudflare (1.1.1.1)"]    = "1.1.1.1"
    $targets["DNS resolution (google.com)"] = "google.com"

    foreach ($label in $targets.Keys) {
        $addr = $targets[$label]
        Write-Host -NoNewline ("  {0,-40}" -f $label) -ForegroundColor $C.Info
        try {
            $ping = Test-Connection -ComputerName $addr -Count 1 -ErrorAction Stop
            $ms   = $ping.ResponseTime
            $color = if ($ms -lt 50) { $C.Success } elseif ($ms -lt 150) { $C.Warning } else { $C.Error }
            Write-Host ("[OK]  {0} ms" -f $ms) -ForegroundColor $color
        } catch {
            Write-Host "[FAIL]" -ForegroundColor $C.Error
        }
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# DNS INFO
# ─────────────────────────────────────────────────────────────────────────────

function Show-DNSInfo {
    Write-Section "DNS CONFIGURATION"

    $dnsClients = Get-DnsClientServerAddress -AddressFamily IPv4 |
                  Where-Object { $_.ServerAddresses -and $_.ServerAddresses.Count -gt 0 }

    foreach ($d in $dnsClients) {
        Write-Host ("  {0,-32} {1}" -f $d.InterfaceAlias, ($d.ServerAddresses -join ', ')) -ForegroundColor $C.Info
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PORT TEST
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-PortTest {
    param([string]$TargetInput = '')

    Write-Section "TCP PORT TEST"

    if (-not $TargetInput) {
        Write-Host "  Format: hostname:port  (e.g. 8.8.8.8:53  or  google.com:443)" -ForegroundColor $C.Info
        Write-Host ""
        Write-Host -NoNewline "  Enter target: " -ForegroundColor $C.Header
        $TargetInput = (Read-Host).Trim()
    }

    if ($TargetInput -notmatch '^(.+):(\d+)$') {
        Write-Host "  [-] Invalid format. Use host:port." -ForegroundColor $C.Error
        return
    }

    $host_ = $Matches[1]
    $port  = [int]$Matches[2]

    Write-Host ""
    Write-Host -NoNewline ("  Testing {0}:{1} ... " -f $host_, $port) -ForegroundColor $C.Progress
    $reachable = Test-TCPPort -Hostname $host_ -Port $port
    if ($reachable) {
        Write-Host "[OPEN]" -ForegroundColor $C.Success
    } else {
        Write-Host "[CLOSED / FILTERED]" -ForegroundColor $C.Error
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# TRACEROUTE
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-Traceroute {
    Write-Section "TRACEROUTE"

    Write-Host -NoNewline "  Target hostname or IP: " -ForegroundColor $C.Header
    $dest = (Read-Host).Trim()
    if (-not $dest) { return }

    Write-Host ""
    Write-Host "  [*] Tracing route to $dest (max 30 hops)..." -ForegroundColor $C.Progress
    Write-Host ""

    try {
        $hops = Test-NetConnection -ComputerName $dest -TraceRoute -ErrorAction Stop
        $i = 1
        foreach ($hop in $hops.TraceRoute) {
            $hopStr = if ($hop -eq '0.0.0.0' -or -not $hop) { '*' } else { $hop }
            Write-Host ("  {0,3}  {1}" -f $i, $hopStr) -ForegroundColor $C.Info
            $i++
        }
    } catch {
        Write-Host "  [-] Traceroute failed: $_" -ForegroundColor $C.Error
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# REMEDIATION — FLUSH DNS
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-FlushDNS {
    Write-Section "FLUSH DNS CACHE"

    Write-Host "  [*] Flushing DNS resolver cache..." -ForegroundColor $C.Progress
    try {
        Clear-DnsClientCache -ErrorAction Stop
        Write-Host "  [+] DNS cache flushed successfully." -ForegroundColor $C.Success
    } catch {
        Write-Host "  [-] Failed to flush DNS: $_" -ForegroundColor $C.Error
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# REMEDIATION — RELEASE / RENEW IP
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-RenewIP {
    Write-Section "RELEASE & RENEW IP (DHCP)"

    $dhcpAdapters = Get-NetIPInterface -AddressFamily IPv4 |
                    Where-Object { $_.Dhcp -eq 'Enabled' -and $_.ConnectionState -eq 'Connected' }

    if (-not $dhcpAdapters) {
        Write-Host "  [!!] No DHCP-enabled adapters found." -ForegroundColor $C.Warning
        return
    }

    foreach ($iface in $dhcpAdapters) {
        Write-Host ("  [*] Releasing IP on: {0}" -f $iface.InterfaceAlias) -ForegroundColor $C.Progress
        try {
            ipconfig /release $iface.InterfaceAlias 2>&1 | Out-Null
        } catch {
            Write-Host ("  [!] Release failed on {0}: {1}" -f $iface.InterfaceAlias, $_) -ForegroundColor $C.Warning
        }
    }

    Start-Sleep -Seconds 2

    foreach ($iface in $dhcpAdapters) {
        Write-Host ("  [*] Renewing IP on:  {0}" -f $iface.InterfaceAlias) -ForegroundColor $C.Progress
        try {
            ipconfig /renew $iface.InterfaceAlias 2>&1 | Out-Null
            $newIP = (Get-NetIPAddress -InterfaceAlias $iface.InterfaceAlias -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                      Select-Object -First 1).IPAddress
            Write-Host ("  [+] New IP: {0}" -f $(if ($newIP) { $newIP } else { 'unknown' })) -ForegroundColor $C.Success
        } catch {
            Write-Host ("  [-] Renew failed on {0}: {1}" -f $iface.InterfaceAlias, $_) -ForegroundColor $C.Error
        }
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# REMEDIATION — RESET NETWORK STACK
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-ResetStack {
    Write-Section "RESET NETWORK STACK"

    Write-Host "  [!!] This resets Winsock, TCP/IP, and firewall rules to defaults." -ForegroundColor $C.Warning
    Write-Host "       A reboot is required to complete the reset." -ForegroundColor $C.Warning
    Write-Host ""
    Write-Host -NoNewline "  Are you sure? (Y/N): " -ForegroundColor $C.Warning
    $confirm = (Read-Host).Trim().ToUpper()
    if ($confirm -ne 'Y') {
        Write-Host "  [*] Cancelled." -ForegroundColor $C.Info
        return
    }

    Write-Host ""
    $steps = @(
        @{ Cmd = "netsh winsock reset";         Label = "Resetting Winsock..."        },
        @{ Cmd = "netsh int ip reset";           Label = "Resetting TCP/IP stack..."   },
        @{ Cmd = "netsh advfirewall reset";      Label = "Resetting firewall rules..."  },
        @{ Cmd = "ipconfig /flushdns";           Label = "Flushing DNS cache..."       }
    )

    foreach ($step in $steps) {
        Write-Host ("  [*] {0}" -f $step.Label) -ForegroundColor $C.Progress
        try {
            Invoke-Expression $step.Cmd 2>&1 | Out-Null
            Write-Host "      Done." -ForegroundColor $C.Success
        } catch {
            Write-Host ("      Failed: {0}" -f $_) -ForegroundColor $C.Error
        }
    }

    Write-Host ""
    Write-Host "  [+] Network stack reset complete. Reboot to apply changes." -ForegroundColor $C.Success
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# NETWORK NEIGHBORS — ARP TABLE (IP + MAC SCAN)
# ─────────────────────────────────────────────────────────────────────────────

function Show-NetworkNeighbors {
    Write-Section "NETWORK NEIGHBORS  (IP / MAC ADDRESS TABLE)"

    # Optional ping sweep to populate the ARP cache before reading it
    if (-not $Unattended) {
        Write-Host -NoNewline "  Ping-sweep local subnet first to discover more hosts? (Y/N): " -ForegroundColor $C.Header
        $sweep = (Read-Host).Trim().ToUpper()
    } else {
        $sweep = 'N'
    }

    if ($sweep -eq 'Y') {
        $localIP = (Get-NetIPAddress -AddressFamily IPv4 |
                    Where-Object { $_.PrefixOrigin -ne 'WellKnown' -and $_.IPAddress -notmatch '^169\.' } |
                    Select-Object -First 1)

        if ($localIP) {
            $prefix = ($localIP.IPAddress -split '\.')[ 0..2] -join '.'
            Write-Host ("  [*] Sweeping {0}.1-254 — this may take a moment..." -f $prefix) -ForegroundColor $C.Progress
            1..254 | ForEach-Object {
                $addr = "$prefix.$_"
                [void](Test-Connection -ComputerName $addr -Count 1 -TimeoutSeconds 1 -ErrorAction SilentlyContinue -AsJob)
            }
            Get-Job | Wait-Job -Timeout 8 | Out-Null
            Get-Job | Remove-Job -Force
            Start-Sleep -Seconds 1
        } else {
            Write-Host "  [!!] Could not determine local subnet — skipping sweep." -ForegroundColor $C.Warning
        }
    }

    Write-Host ""

    $neighbors = Get-NetNeighbor -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                 Where-Object { $_.LinkLayerAddress -and $_.LinkLayerAddress -ne '00-00-00-00-00-00' -and $_.State -ne 'Unreachable' } |
                 Sort-Object -Property IPAddress

    if (-not $neighbors) {
        Write-Host "  [-] No neighbor entries found." -ForegroundColor $C.Warning
        Write-Host ""
        return
    }

    $header = "  {0,-18} {1,-20} {2,-30} {3}" -f "IP Address", "MAC Address", "Adapter", "State"
    Write-Host $header -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 74)) -ForegroundColor $C.Header

    foreach ($n in $neighbors) {
        $adapter = (Get-NetAdapter -InterfaceIndex $n.InterfaceIndex -ErrorAction SilentlyContinue).Name
        $stateColor = switch ($n.State) {
            'Reachable'   { $C.Success }
            'Stale'       { $C.Warning }
            'Incomplete'  { $C.Error   }
            default       { $C.Info    }
        }
        Write-Host ("  {0,-18} {1,-20} {2,-30} {3}" -f $n.IPAddress, $n.LinkLayerAddress, $adapter, $n.State) -ForegroundColor $stateColor
    }

    Write-Host ""
    Write-Host ("  Total entries: {0}" -f $neighbors.Count) -ForegroundColor $C.Info
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# FULL STATUS (adapters + connectivity + DNS)
# ─────────────────────────────────────────────────────────────────────────────

function Show-FullStatus {
    Show-AdapterStatus
    Show-ConnectivityTests
    Show-DNSInfo
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-LeylineBanner
    switch ($Action) {
        'Status'     { Show-FullStatus }
        'FlushDNS'   { Invoke-FlushDNS }
        'Renew'      { Invoke-RenewIP }
        'ResetStack' { Write-Host "  [!!] ResetStack requires interactive confirmation — run without -Unattended." -ForegroundColor $C.Warning }
        'PortTest'   { Invoke-PortTest -TargetInput $Target }
        'Trace'      { Write-Host "  [!!] Trace requires interactive input — run without -Unattended." -ForegroundColor $C.Warning }
        'ARPScan'    { Show-NetworkNeighbors }
    }
} else {
    $choice = ''

    do {
        Show-LeylineBanner
        Show-FullStatus

        Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS" -ForegroundColor $C.Header
        Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] TCP port test" -ForegroundColor $C.Info
        Write-Host "  [2] Traceroute" -ForegroundColor $C.Info
        Write-Host "  [3] Flush DNS cache" -ForegroundColor $C.Info
        Write-Host "  [4] Release & renew IP  (DHCP)" -ForegroundColor $C.Info
        Write-Host "  [5] Reset network stack  (requires reboot)" -ForegroundColor $C.Info
        Write-Host "  [6] Show network neighbors  (IP / MAC table)" -ForegroundColor $C.Info
        Write-Host "  [R] Refresh status" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' { Invoke-PortTest }
            '2' { Invoke-Traceroute }
            '3' { Invoke-FlushDNS }
            '4' { Invoke-RenewIP }
            '5' { Invoke-ResetStack }
            '6' { Show-NetworkNeighbors }
            'R' { }
            'Q' {
                Write-Host ""
                Write-Host "  Closing L.E.Y.L.I.N.E." -ForegroundColor $C.Header
                Write-Host ""
            }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1-6, R, or Q." -ForegroundColor $C.Warning
                Start-Sleep -Seconds 1
            }
        }

        if ($choice -notin @('Q','R')) {
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

    } while ($choice -ne 'Q')
}
if ($Transcript) { Stop-TKTranscript }
