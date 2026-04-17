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

    Color Schema
    ─────────────────────────────────────────
    Cyan     Headers and section dividers
    Magenta  Progress indicators
    Green    Success / reachable
    Yellow   Warnings / degraded
    Red      Critical errors / unreachable
    Gray     Information and details
#>

param(
    [switch]$Unattended,
    [ValidateSet('Status','FlushDNS','Renew','ResetStack','PortTest','Trace')]
    [string]$Action = 'Status',
    [string]$Target = ''
)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK
# ─────────────────────────────────────────────────────────────────────────────

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script must be run as Administrator." -ForegroundColor Red
    exit 1
}

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

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

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host "  $Title" -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host ""
}

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
        } catch {}
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
            'R' { }
            'Q' {
                Write-Host ""
                Write-Host "  Closing L.E.Y.L.I.N.E." -ForegroundColor $C.Header
                Write-Host ""
            }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1-5, R, or Q." -ForegroundColor $C.Warning
                Start-Sleep -Seconds 1
            }
        }

        if ($choice -notin @('Q','R')) {
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

    } while ($choice -ne 'Q')
}
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
