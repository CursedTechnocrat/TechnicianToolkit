<#
.SYNOPSIS
    L.A.N.T.E.R.N. — Locates & Audits Network Topology, Enumerating Resources & Nodes
    Network Discovery & Asset Inventory Tool for PowerShell 5.1+

.DESCRIPTION
    Scans the local /24 subnet to discover all live hosts via parallel ping sweep,
    resolves hostnames through DNS reverse lookup, reads MAC addresses from the ARP
    neighbor table, optionally performs TCP port scans against common service ports,
    and produces a dark-themed HTML inventory report with color-coded port badges
    and summary cards. Results can also be exported to CSV.

.USAGE
    PS C:\> .\lantern.ps1                           # Interactive menu
    PS C:\> .\lantern.ps1 -Unattended -Action Sweep # Run sweep and export HTML silently

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
    Green    Success / reachable
    Yellow   Warnings / degraded
    Red      Critical errors / unreachable
    Gray     Information and details
#>

param(
    [switch]$Unattended,
    [ValidateSet('Sweep')]
    [string]$Action = 'Sweep',
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
}

# ─────────────────────────────────────────────────────────────────────────────
# SHARED STATE
# ─────────────────────────────────────────────────────────────────────────────

$script:DiscoveredHosts = @()

# Common ports to scan
$script:ScanPorts = @(21, 22, 23, 80, 443, 445, 3389, 5985, 8080, 8443)

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-LanternBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██╗      █████╗ ███╗   ██╗████████╗███████╗██████╗ ███╗   ██╗
  ██║     ██╔══██╗████╗  ██║╚══██╔══╝██╔════╝██╔══██╗████╗  ██║
  ██║     ███████║██╔██╗ ██║   ██║   █████╗  ██████╔╝██╔██╗ ██║
  ██║     ██╔══██║██║╚██╗██║   ██║   ██╔══╝  ██╔══██╗██║╚██╗██║
  ███████╗██║  ██║██║ ╚████║   ██║   ███████╗██║  ██║██║ ╚████║
  ╚══════╝╚═╝  ╚═╝╚═╝  ╚═══╝   ╚═╝   ╚══════╝╚═╝  ╚═╝╚═╝  ╚═══╝

"@ -ForegroundColor Cyan
    Write-Host "    L.A.N.T.E.R.N. — Locates & Audits Network Topology, Enumerating Resources & Nodes" -ForegroundColor Cyan
    Write-Host "    Network Discovery & Asset Inventory Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function Test-TCPPort {
    param([string]$Hostname, [int]$Port, [int]$TimeoutMs = 1500)
    try {
        $tcp = [System.Net.Sockets.TcpClient]::new()
        $ar  = $tcp.BeginConnect($Hostname, $Port, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        $tcp.Close()
        return $ok
    } catch { return $false }
}

function Get-LocalSubnetPrefix {
    $ip = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
          Where-Object {
              $_.IPAddress -notmatch '^127\.' -and
              $_.IPAddress -notmatch '^169\.254\.' -and
              $_.PrefixOrigin -ne 'WellKnown'
          } |
          Sort-Object -Property PrefixLength -Descending |
          Select-Object -First 1

    if (-not $ip) { return $null }
    $parts = $ip.IPAddress -split '\.'
    return ($parts[0..2] -join '.')
}

# ─────────────────────────────────────────────────────────────────────────────
# SUBNET SWEEP
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-SubnetSweep {
    Write-Section "SUBNET DISCOVERY SWEEP"

    $prefix = Get-LocalSubnetPrefix
    if (-not $prefix) {
        Write-Host "  [-] Could not determine local IP address. Ensure a network adapter is connected." -ForegroundColor $C.Error
        Write-Host ""
        return
    }

    Write-Host ("  [*] Local subnet detected: {0}.0/24" -f $prefix) -ForegroundColor $C.Info
    Write-Host ("  [*] Sweeping {0}.1 – {0}.254 using parallel ping jobs..." -f $prefix) -ForegroundColor $C.Progress
    Write-Host ""

    # Build runspace pool for parallel pings
    $pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 50)
    $pool.Open()

    $jobs = [System.Collections.Generic.List[hashtable]]::new()

    $pingScript = {
        param([string]$IP)
        $result = @{ IP = $IP; Alive = $false; ResponseTimeMs = -1 }
        try {
            $ping = [System.Net.NetworkInformation.Ping]::new()
            $reply = $ping.Send($IP, 1000)
            if ($reply.Status -eq 'Success') {
                $result.Alive = $true
                $result.ResponseTimeMs = [int]$reply.RoundtripTime
            }
        } catch {}
        return $result
    }

    # Submit all 254 jobs
    for ($i = 1; $i -le 254; $i++) {
        $ip  = "$prefix.$i"
        $ps  = [System.Management.Automation.PowerShell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript($pingScript).AddArgument($ip)
        $handle = $ps.BeginInvoke()
        $jobs.Add(@{ PS = $ps; Handle = $handle; Index = $i })
    }

    # Wait and show progress
    $timeout   = (Get-Date).AddSeconds(15)
    $completed = 0
    $alive     = [System.Collections.Generic.List[hashtable]]::new()

    while ($completed -lt $jobs.Count -and (Get-Date) -lt $timeout) {
        foreach ($job in $jobs) {
            if (-not $job.ContainsKey('Done') -and $job.Handle.IsCompleted) {
                $res = $job.PS.EndInvoke($job.Handle)
                $job.PS.Dispose()
                $job['Done'] = $true
                $completed++
                if ($res -and $res.Alive) {
                    $alive.Add($res)
                }
            }
        }
        $pct = [int](($completed / 254) * 100)
        Write-Progress -Activity "Sweeping $prefix.x" `
                       -Status ("Checked {0}/254 hosts — {1} responding" -f $completed, $alive.Count) `
                       -PercentComplete $pct
        Start-Sleep -Milliseconds 50
    }

    Write-Progress -Activity "Sweeping $prefix.x" -Completed

    # Force-close any incomplete jobs
    foreach ($job in $jobs) {
        if (-not $job.ContainsKey('Done')) {
            try { $job.PS.Stop() } catch {}
            $job.PS.Dispose()
        }
    }
    $pool.Close()
    $pool.Dispose()

    Write-Host ("  [+] Sweep complete. Found {0} responding host(s)." -f $alive.Count) -ForegroundColor $C.Success
    Write-Host ""

    if ($alive.Count -eq 0) {
        Write-Host "  [!!] No hosts responded. Check your network connection." -ForegroundColor $C.Warning
        Write-Host ""
        return
    }

    # Read ARP table into a lookup dictionary
    Write-Host "  [*] Reading ARP neighbor table..." -ForegroundColor $C.Progress
    $arpLookup = @{}
    Get-NetNeighbor -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.LinkLayerAddress -and $_.LinkLayerAddress -ne '00-00-00-00-00-00' } |
        ForEach-Object { $arpLookup[$_.IPAddress] = $_.LinkLayerAddress }

    # Build discovered hosts list
    $script:DiscoveredHosts = @()
    $total  = $alive.Count
    $idx    = 0

    Write-Host "  [*] Resolving hostnames..." -ForegroundColor $C.Progress
    Write-Host ""

    foreach ($r in ($alive | Sort-Object { [version]$_.IP })) {
        $idx++
        Write-Progress -Activity "Resolving hostnames" `
                       -Status ("Resolving {0} ({1}/{2})" -f $r.IP, $idx, $total) `
                       -PercentComplete ([int](($idx / $total) * 100))

        $hostname = ''
        try {
            $entry    = [System.Net.Dns]::GetHostEntry($r.IP)
            $hostname = $entry.HostName
        } catch {}

        $mac = if ($arpLookup.ContainsKey($r.IP)) { $arpLookup[$r.IP] } else { '' }

        $script:DiscoveredHosts += [PSCustomObject]@{
            IP            = $r.IP
            Hostname      = $hostname
            MAC           = $mac
            Vendor        = ''
            OpenPorts     = @()
            ResponseTimeMs = $r.ResponseTimeMs
            LastSeen      = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        }
    }

    Write-Progress -Activity "Resolving hostnames" -Completed

    Write-Host ("  [+] Discovery complete. {0} host(s) inventoried." -f $script:DiscoveredHosts.Count) -ForegroundColor $C.Success
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PORT SCAN
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-PortScan {
    Write-Section "TCP PORT SCAN"

    if ($script:DiscoveredHosts.Count -eq 0) {
        Write-Host "  [!!] No hosts discovered yet. Run option 1 first." -ForegroundColor $C.Warning
        Write-Host ""
        return
    }

    $portList  = $script:ScanPorts
    $hostCount = $script:DiscoveredHosts.Count
    $totalOps  = $hostCount * $portList.Count
    $done      = 0

    Write-Host ("  [*] Scanning {0} host(s) across {1} ports ({2} total checks)..." -f $hostCount, $portList.Count, $totalOps) -ForegroundColor $C.Progress
    Write-Host ("      Ports: {0}" -f ($portList -join ', ')) -ForegroundColor $C.Info
    Write-Host ""

    foreach ($h in $script:DiscoveredHosts) {
        $h.OpenPorts = @()
        foreach ($port in $portList) {
            $done++
            $pct = [int](($done / $totalOps) * 100)
            Write-Progress -Activity "Port scanning" `
                           -Status ("Testing {0}:{1} ({2}/{3})" -f $h.IP, $port, $done, $totalOps) `
                           -PercentComplete $pct

            if (Test-TCPPort -Hostname $h.IP -Port $port -TimeoutMs 1500) {
                $h.OpenPorts += $port
            }
        }

        $portDisplay = if ($h.OpenPorts.Count -gt 0) {
            $h.OpenPorts -join ', '
        } else {
            'none'
        }
        $portColor = if ($h.OpenPorts.Count -gt 0) { $C.Warning } else { $C.Info }
        Write-Host ("  {0,-18} {1}" -f $h.IP, $portDisplay) -ForegroundColor $portColor
    }

    Write-Progress -Activity "Port scanning" -Completed
    Write-Host ""
    Write-Host "  [+] Port scan complete." -ForegroundColor $C.Success
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# SHOW HOST TABLE
# ─────────────────────────────────────────────────────────────────────────────

function Show-HostTable {
    Write-Section "DISCOVERED HOSTS"

    if ($script:DiscoveredHosts.Count -eq 0) {
        Write-Host "  [!!] No hosts discovered yet. Run option 1 first." -ForegroundColor $C.Warning
        Write-Host ""
        return
    }

    $hdr = "  {0,-18} {1,-30} {2,-20} {3,-8} {4}" -f "IP Address", "Hostname", "MAC Address", "RTT(ms)", "Open Ports"
    Write-Host $hdr -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 90)) -ForegroundColor $C.Header

    foreach ($h in $script:DiscoveredHosts | Sort-Object { [version]$_.IP }) {
        $ports    = if ($h.OpenPorts.Count -gt 0) { $h.OpenPorts -join ', ' } else { '—' }
        $hostname = if ($h.Hostname) { $h.Hostname } else { '—' }
        $mac      = if ($h.MAC)      { $h.MAC      } else { '—' }
        $rtt      = if ($h.ResponseTimeMs -ge 0) { $h.ResponseTimeMs.ToString() } else { '—' }

        $rowColor = if ($h.OpenPorts -contains 3389 -or $h.OpenPorts -contains 445) {
            $C.Warning
        } elseif ($h.OpenPorts.Count -gt 0) {
            $C.Success
        } else {
            $C.Info
        }

        Write-Host ("  {0,-18} {1,-30} {2,-20} {3,-8} {4}" -f $h.IP, $hostname, $mac, $rtt, $ports) -ForegroundColor $rowColor
    }

    Write-Host ""
    Write-Host ("  Total hosts: {0}" -f $script:DiscoveredHosts.Count) -ForegroundColor $C.Info

    $rdpExposed = ($script:DiscoveredHosts | Where-Object { $_.OpenPorts -contains 3389 }).Count
    $smbExposed = ($script:DiscoveredHosts | Where-Object { $_.OpenPorts -contains 445  }).Count

    if ($rdpExposed -gt 0) {
        Write-Host ("  [!!] RDP exposed (port 3389): {0} host(s)" -f $rdpExposed) -ForegroundColor $C.Warning
    }
    if ($smbExposed -gt 0) {
        Write-Host ("  [!!] SMB exposed (port 445): {0} host(s)"  -f $smbExposed) -ForegroundColor $C.Warning
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    Write-Section "EXPORT HTML REPORT"

    if ($script:DiscoveredHosts.Count -eq 0) {
        Write-Host "  [!!] No hosts discovered yet. Run option 1 first." -ForegroundColor $C.Warning
        Write-Host ""
        return
    }

    $timestamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reportName   = "LANTERN_$timestamp.html"
    $reportPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) $reportName
    $generated    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    $totalHosts   = $script:DiscoveredHosts.Count
    $rdpExposed   = ($script:DiscoveredHosts | Where-Object { $_.OpenPorts -contains 3389 }).Count
    $smbExposed   = ($script:DiscoveredHosts | Where-Object { $_.OpenPorts -contains 445  }).Count

    Write-Host "  [*] Building HTML report..." -ForegroundColor $C.Progress

    # Port badge helper
    function Get-PortBadgeHtml {
        param([int[]]$Ports)
        if (-not $Ports -or $Ports.Count -eq 0) { return '<span class="tk-mono" style="color:#666;">none</span>' }

        $badges = foreach ($p in $Ports) {
            $label = switch ($p) {
                21   { 'FTP'      }
                22   { 'SSH'      }
                23   { 'Telnet'   }
                80   { 'HTTP'     }
                443  { 'HTTPS'    }
                445  { 'SMB'      }
                3389 { 'RDP'      }
                5985 { 'WinRM'    }
                8080 { 'HTTP-Alt' }
                8443 { 'HTTPS-Alt'}
                default { "$p"   }
            }

            $badgeClass = switch ($p) {
                { $_ -in @(3389, 445, 23) } { 'tk-badge-err'  }
                { $_ -in @(80, 8080) }       { 'tk-badge-warn' }
                { $_ -in @(443, 8443) }      { 'tk-badge-ok'   }
                default                      { 'tk-badge-info' }
            }

            "<span class=`"$badgeClass`">$label</span>"
        }
        return $badges -join ' '
    }

    # Build table rows
    $rows = foreach ($h in $script:DiscoveredHosts | Sort-Object { [version]$_.IP }) {
        $hostname   = if ($h.Hostname) { [System.Net.WebUtility]::HtmlEncode($h.Hostname) } else { '<span style="color:#555;">-</span>' }
        $mac        = if ($h.MAC)      { $h.MAC } else { '<span style="color:#555;">-</span>' }
        $rtt        = if ($h.ResponseTimeMs -ge 0) { "$($h.ResponseTimeMs) ms" } else { '-' }
        $portBadges = Get-PortBadgeHtml -Ports $h.OpenPorts

        $rowRisk = if ($h.OpenPorts -contains 3389 -or $h.OpenPorts -contains 445 -or $h.OpenPorts -contains 23) {
            'err'
        } elseif ($h.OpenPorts.Count -gt 0) {
            'ok'
        } else {
            ''
        }

        @"
            <tr>
                <td>$($h.IP)</td>
                <td>$hostname</td>
                <td class="tk-mono">$mac</td>
                <td>$portBadges</td>
                <td>$rtt</td>
                $(if ($rowRisk) { "<td><span class=`"tk-badge-$rowRisk`">$(if ($rowRisk -eq 'err') {'Risk'} else {'Active'})</span></td>" } else { '<td></td>' })
            </tr>
"@
    }

    $tableRows = $rows -join "`n"

    $rdpClass = if ($rdpExposed -gt 0) { 'err' } else { 'ok' }
    $smbClass = if ($smbExposed -gt 0) { 'warn' } else { 'ok' }

    $subnet = Get-LocalSubnetPrefix
    $subnetLabel = if ($subnet) { "$subnet.0/24" } else { 'Unknown' }

    $htmlReport = (Get-TKHtmlHead `
        -Title      'L.A.N.T.E.R.N. Network Discovery Report' `
        -ScriptName 'L.A.N.T.E.R.N.' `
        -Subtitle   'Network Discovery & Asset Inventory' `
        -MetaItems  ([ordered]@{
            'Generated' = $generated
            'Subnet'    = $subnetLabel
            'Host'      = $env:COMPUTERNAME
        }) `
        -NavItems   @('Host Inventory')) + @"

<div class="tk-info-box">
  <span class="tk-info-label">Note</span>
  Port scan covers: 21 (FTP), 22 (SSH), 23 (Telnet), 80 (HTTP), 443 (HTTPS), 445 (SMB), 3389 (RDP), 5985 (WinRM), 8080, 8443.
  An empty port column means no port scan has been performed yet for that host.
</div>

<div class="tk-summary-row">
  <div class="tk-summary-card ok">
    <div class="tk-summary-num">$totalHosts</div>
    <div class="tk-summary-lbl">Total Hosts</div>
  </div>
  <div class="tk-summary-card ok">
    <div class="tk-summary-num">$totalHosts</div>
    <div class="tk-summary-lbl">Responding</div>
  </div>
  <div class="tk-summary-card $rdpClass">
    <div class="tk-summary-num">$rdpExposed</div>
    <div class="tk-summary-lbl">RDP Exposed (3389)</div>
  </div>
  <div class="tk-summary-card $smbClass">
    <div class="tk-summary-num">$smbExposed</div>
    <div class="tk-summary-lbl">SMB Exposed (445)</div>
  </div>
</div>

<div class="tk-section">
  <div class="tk-section-title">Host Inventory</div>
  <table class="tk-table">
    <thead>
      <tr>
        <th>IP Address</th>
        <th>Hostname</th>
        <th>MAC Address</th>
        <th>Open Ports</th>
        <th>Response Time</th>
        <th>Risk</th>
      </tr>
    </thead>
    <tbody>
$tableRows
    </tbody>
  </table>
</div>

"@ + (Get-TKHtmlFoot -ScriptName 'L.A.N.T.E.R.N. v1.0')

    try {
        [System.IO.File]::WriteAllText($reportPath, $htmlReport, [System.Text.Encoding]::UTF8)
        Write-Host ("  [+] HTML report saved: {0}" -f $reportPath) -ForegroundColor $C.Success
    } catch {
        Write-Host ("  [-] Failed to write report: {0}" -f $_) -ForegroundColor $C.Error
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# CSV EXPORT
# ─────────────────────────────────────────────────────────────────────────────

function Export-CsvReport {
    Write-Section "EXPORT CSV"

    if ($script:DiscoveredHosts.Count -eq 0) {
        Write-Host "  [!!] No hosts discovered yet. Run option 1 first." -ForegroundColor $C.Warning
        Write-Host ""
        return
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csvPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "LANTERN_$timestamp.csv"

    try {
        $script:DiscoveredHosts |
            Sort-Object { [version]$_.IP } |
            Select-Object IP, Hostname, MAC,
                @{ N='OpenPorts'; E={ $_.OpenPorts -join ',' } },
                ResponseTimeMs, LastSeen |
            Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

        Write-Host ("  [+] CSV saved: {0}" -f $csvPath) -ForegroundColor $C.Success
    } catch {
        Write-Host ("  [-] Failed to write CSV: {0}" -f $_) -ForegroundColor $C.Error
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — UNATTENDED OR INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-LanternBanner
    switch ($Action) {
        'Sweep' {
            Invoke-SubnetSweep
            Build-HtmlReport
        }
    }
} else {
    $choice = ''

    do {
        Show-LanternBanner

        # Status line if hosts are already known
        if ($script:DiscoveredHosts.Count -gt 0) {
            $portScanned = ($script:DiscoveredHosts | Where-Object { $_.OpenPorts.Count -gt 0 -or $_.OpenPorts -ne $null }).Count
            Write-Host ("  Inventory: {0} host(s) in memory" -f $script:DiscoveredHosts.Count) -ForegroundColor $C.Info

            $lastSeen = $script:DiscoveredHosts | Select-Object -Last 1 -ExpandProperty LastSeen
            Write-Host ("  Last sweep: {0}" -f $lastSeen) -ForegroundColor $C.Info
            Write-Host ""
        }

        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host "  ACTIONS" -ForegroundColor $C.Header
        Write-Host ("  " + ("-" * 62)) -ForegroundColor $C.Header
        Write-Host ""
        Write-Host "  [1] Run discovery          (ping sweep /24 subnet)" -ForegroundColor $C.Info
        Write-Host "  [2] Show discovered hosts  (console table)" -ForegroundColor $C.Info
        Write-Host "  [3] Port scan hosts        (TCP connect, 10 common ports)" -ForegroundColor $C.Info
        Write-Host "  [4] Export HTML report     (dark-themed inventory)" -ForegroundColor $C.Info
        Write-Host "  [5] Export CSV" -ForegroundColor $C.Info
        Write-Host "  [R] Re-run discovery" -ForegroundColor $C.Info
        Write-Host "  [Q] Quit" -ForegroundColor $C.Info
        Write-Host ""
        Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
        $choice = (Read-Host).Trim().ToUpper()

        switch ($choice) {
            '1' { Invoke-SubnetSweep }
            '2' { Show-HostTable }
            '3' { Invoke-PortScan }
            '4' { Build-HtmlReport }
            '5' { Export-CsvReport }
            'R' { Invoke-SubnetSweep }
            'Q' {
                Write-Host ""
                Write-Host "  Closing L.A.N.T.E.R.N." -ForegroundColor $C.Header
                Write-Host ""
            }
            default {
                Write-Host ""
                Write-Host "  [!!] Invalid selection. Enter 1-5, R, or Q." -ForegroundColor $C.Warning
                Start-Sleep -Seconds 1
            }
        }

        if ($choice -notin @('Q', 'R', '1')) {
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

    } while ($choice -ne 'Q')
}

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
