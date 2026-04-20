<#
.SYNOPSIS
    R.E.L.I.C. — Reviews & Enumerates Latent Identity Certificates
    Certificate Health Monitor for PowerShell 5.1+

.DESCRIPTION
    Audits the local Windows certificate stores (Personal, CA, Trusted Root) for
    expired and expiring-soon certificates. Optionally checks SSL/TLS certificate
    expiry on remote hosts by connecting and reading their presented certificate.
    Exports a dark-themed HTML report with color-coded expiry indicators.

.USAGE
    PS C:\> .\relic.ps1                                          # Interactive menu
    PS C:\> .\relic.ps1 -Unattended                              # Full audit (local + HTML report)
    PS C:\> .\relic.ps1 -Unattended -Targets "srv1.contoso.com,srv2.contoso.com:8443"

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
    Green    Healthy / valid
    Yellow   Expiring soon (30–90 days)
    Red      Expired or critical (<30 days)
    Gray     Information and details
#>

param(
    [switch]$Unattended,
    [string]$Targets = '',
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
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-RelicBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

  ██████╗ ███████╗██╗     ██╗ ██████╗
  ██╔══██╗██╔════╝██║     ██║██╔════╝
  ██████╔╝█████╗  ██║     ██║██║
  ██╔══██╗██╔══╝  ██║     ██║██║
  ██║  ██║███████╗███████╗██║╚██████╗
  ╚═╝  ╚═╝╚══════╝╚══════╝╚═╝ ╚═════╝

"@ -ForegroundColor Cyan
    Write-Host "    R.E.L.I.C. — Reviews & Enumerates Latent Identity Certificates" -ForegroundColor Cyan
    Write-Host "    Certificate Health & SSL Expiry Monitor" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

function HtmlEncode {
    param([string]$s)
    $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

# ─────────────────────────────────────────────────────────────────────────────
# CORE DATA FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Get-LocalCertHealth {
    param([string[]]$StoreNames = @('My', 'CA', 'Root', 'TrustedPublisher'))
    $results = @()
    foreach ($storeName in $StoreNames) {
        try {
            $store = [Security.Cryptography.X509Certificates.X509Store]::new($storeName, 'LocalMachine')
            $store.Open('ReadOnly')
            foreach ($cert in $store.Certificates) {
                $daysLeft = [int]($cert.NotAfter - (Get-Date)).TotalDays
                $status   = if ($daysLeft -lt 0)   { 'Expired'  }
                            elseif ($daysLeft -le 30)  { 'Critical' }
                            elseif ($daysLeft -le 90)  { 'Warning'  }
                            else                       { 'Healthy'  }
                $results += [PSCustomObject]@{
                    Store      = $storeName
                    Subject    = $cert.Subject
                    Issuer     = $cert.Issuer
                    Thumbprint = $cert.Thumbprint.Substring(0, [Math]::Min(16, $cert.Thumbprint.Length)) + '...'
                    Expiry     = $cert.NotAfter
                    DaysLeft   = $daysLeft
                    Status     = $status
                }
            }
            $store.Close()
        } catch {}
    }
    return $results
}

function Get-SslCertExpiry {
    param([string]$Hostname, [int]$Port = 443, [int]$TimeoutMs = 6000)

    $result = [PSCustomObject]@{
        Hostname = $Hostname
        Port     = $Port
        Subject  = 'N/A'
        Issuer   = 'N/A'
        Expiry   = $null
        DaysLeft = $null
        Status   = 'Unknown'
        Error    = $null
    }

    try {
        $tcp = New-Object Net.Sockets.TcpClient
        $connect = $tcp.BeginConnect($Hostname, $Port, $null, $null)
        $wait = $connect.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if (-not $wait) { throw "Connection timed out" }
        $tcp.EndConnect($connect)
        $ssl = New-Object Net.Security.SslStream($tcp.GetStream(), $false, { $true })
        $ssl.AuthenticateAsClient($Hostname)
        $cert  = $ssl.RemoteCertificate
        $cert2 = New-Object Security.Cryptography.X509Certificates.X509Certificate2($cert)
        $result.Subject  = $cert2.Subject
        $result.Issuer   = $cert2.Issuer
        $result.Expiry   = $cert2.NotAfter
        $result.DaysLeft = [int]($cert2.NotAfter - (Get-Date)).TotalDays
        $result.Status   = if ($result.DaysLeft -lt 0)   { 'Expired'  }
                           elseif ($result.DaysLeft -le 30)  { 'Critical' }
                           elseif ($result.DaysLeft -le 90)  { 'Warning'  }
                           else                              { 'Healthy'  }
        try { $ssl.Dispose() } catch {}
        try { $tcp.Dispose() } catch {}
    } catch {
        $result.Status = 'Error'
        $result.Error  = $_.Exception.Message
    }

    return $result
}

function ConvertTo-TargetList {
    param([string]$TargetString)

    $entries = @()

    if ([string]::IsNullOrWhiteSpace($TargetString)) {
        return $entries
    }

    # If it looks like a file path and the file exists, read it line by line
    if (Test-Path $TargetString -PathType Leaf) {
        $lines = Get-Content $TargetString -ErrorAction SilentlyContinue |
                 Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.TrimStart().StartsWith('#') }
    } else {
        $lines = $TargetString -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }

    foreach ($line in $lines) {
        $line = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        # Check for hostname:port format
        if ($line -match '^(.+):(\d+)$') {
            $entries += [PSCustomObject]@{
                Hostname = $Matches[1].Trim()
                Port     = [int]$Matches[2]
            }
        } else {
            $entries += [PSCustomObject]@{
                Hostname = $line
                Port     = 443
            }
        }
    }

    return $entries
}

# ─────────────────────────────────────────────────────────────────────────────
# CONSOLE DISPLAY FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Show-LocalCertSummary {
    param([array]$Certs)

    Write-Section "LOCAL CERTIFICATE STORE AUDIT"

    if (-not $Certs -or $Certs.Count -eq 0) {
        Write-Host "  No certificates found." -ForegroundColor $C.Warning
        return
    }

    $expired  = @($Certs | Where-Object { $_.Status -eq 'Expired'  })
    $critical = @($Certs | Where-Object { $_.Status -eq 'Critical' })
    $warning  = @($Certs | Where-Object { $_.Status -eq 'Warning'  })
    $healthy  = @($Certs | Where-Object { $_.Status -eq 'Healthy'  })

    # Summary counts
    Write-Host ("  Total: {0}   Expired: {1}   Critical: {2}   Warning: {3}   Healthy: {4}" -f `
        $Certs.Count, $expired.Count, $critical.Count, $warning.Count, $healthy.Count) -ForegroundColor $C.Info
    Write-Host ""

    $nonHealthy = @($Certs | Where-Object { $_.Status -ne 'Healthy' } | Sort-Object DaysLeft)

    if ($nonHealthy.Count -eq 0) {
        Write-Host "  [+] All certificates are healthy." -ForegroundColor $C.Success
        return
    }

    # Table header
    $col = @(10, 52, 22, 10, 10)
    Write-Host ("  {0,-$($col[0])} {1,-$($col[1])} {2,-$($col[2])} {3,-$($col[3])} {4}" -f `
        'Store', 'Subject', 'Expiry', 'Days Left', 'Status') -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 110)) -ForegroundColor $C.Header

    foreach ($cert in $nonHealthy) {
        $rowColor = switch ($cert.Status) {
            'Expired'  { $C.Error   }
            'Critical' { $C.Error   }
            'Warning'  { $C.Warning }
            default    { $C.Info    }
        }

        $subjectShort = if ($cert.Subject.Length -gt 50) { $cert.Subject.Substring(0,47) + '...' } else { $cert.Subject }
        $expiryStr    = $cert.Expiry.ToString('yyyy-MM-dd HH:mm')

        Write-Host ("  {0,-$($col[0])} {1,-$($col[1])} {2,-$($col[2])} {3,-$($col[3])} {4}" -f `
            $cert.Store, $subjectShort, $expiryStr, $cert.DaysLeft, $cert.Status) -ForegroundColor $rowColor
    }

    Write-Host ""
}

function Show-SslSummary {
    param([array]$SslResults)

    Write-Section "SSL/TLS REMOTE CERTIFICATE CHECKS"

    if (-not $SslResults -or $SslResults.Count -eq 0) {
        Write-Host "  No remote hosts checked." -ForegroundColor $C.Warning
        return
    }

    # Table header
    Write-Host ("  {0,-35} {1,-45} {2,-22} {3,-10} {4}" -f `
        'Host:Port', 'Subject', 'Expiry', 'Days Left', 'Status') -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 120)) -ForegroundColor $C.Header

    foreach ($r in $SslResults) {
        $hostPort = "$($r.Hostname):$($r.Port)"
        if ($hostPort.Length -gt 34) { $hostPort = $hostPort.Substring(0,31) + '...' }

        $rowColor = switch ($r.Status) {
            'Expired'  { $C.Error   }
            'Critical' { $C.Error   }
            'Warning'  { $C.Warning }
            'Error'    { $C.Error   }
            'Healthy'  { $C.Success }
            default    { $C.Info    }
        }

        if ($r.Status -eq 'Error') {
            $subjectShort = if ($r.Error -and $r.Error.Length -gt 44) { $r.Error.Substring(0,41) + '...' } else { $r.Error }
            Write-Host ("  {0,-35} {1,-45} {2,-22} {3,-10} {4}" -f `
                $hostPort, $subjectShort, 'N/A', 'N/A', 'Error') -ForegroundColor $rowColor
        } else {
            $subjectShort = if ($r.Subject.Length -gt 44) { $r.Subject.Substring(0,41) + '...' } else { $r.Subject }
            $expiryStr    = if ($r.Expiry) { $r.Expiry.ToString('yyyy-MM-dd HH:mm') } else { 'N/A' }
            $daysStr      = if ($null -ne $r.DaysLeft) { "$($r.DaysLeft)" } else { 'N/A' }

            Write-Host ("  {0,-35} {1,-45} {2,-22} {3,-10} {4}" -f `
                $hostPort, $subjectShort, $expiryStr, $daysStr, $r.Status) -ForegroundColor $rowColor
        }
    }

    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param(
        [array]$LocalCerts,
        [array]$SslResults,
        [string]$MachineName,
        [string]$ReportTimestamp
    )

    # Summary counts
    $totalCerts   = if ($LocalCerts) { $LocalCerts.Count } else { 0 }
    $expiredCount = if ($LocalCerts) { ($LocalCerts | Where-Object { $_.Status -eq 'Expired'  } | Measure-Object).Count } else { 0 }
    $critCount    = if ($LocalCerts) { ($LocalCerts | Where-Object { $_.Status -eq 'Critical' } | Measure-Object).Count } else { 0 }
    $warnCount    = if ($LocalCerts) { ($LocalCerts | Where-Object { $_.Status -eq 'Warning'  } | Measure-Object).Count } else { 0 }
    $healthyCount = if ($LocalCerts) { ($LocalCerts | Where-Object { $_.Status -eq 'Healthy'  } | Measure-Object).Count } else { 0 }

    # Build local cert rows — ALL certs included in HTML, sorted by DaysLeft ascending
    $localRows = ''
    if ($LocalCerts) {
        foreach ($cert in ($LocalCerts | Sort-Object DaysLeft)) {
            $badgeClass = switch ($cert.Status) {
                'Expired'  { 'badge-crit' }
                'Critical' { 'badge-crit' }
                'Warning'  { 'badge-warn' }
                default    { 'badge-ok'   }
            }
            $expiryStr = $cert.Expiry.ToString('yyyy-MM-dd HH:mm')
            $daysStr   = if ($cert.DaysLeft -lt 0) { "$($cert.DaysLeft)" } else { "$($cert.DaysLeft)" }

            $localRows += "            <tr>
                <td>$(HtmlEncode $cert.Store)</td>
                <td style='font-size:12px;word-break:break-all;'>$(HtmlEncode $cert.Subject)</td>
                <td style='font-size:11px;color:#888;word-break:break-all;'>$(HtmlEncode $cert.Issuer)</td>
                <td>$expiryStr</td>
                <td>$daysStr</td>
                <td><span class='badge $badgeClass'>$(HtmlEncode $cert.Status)</span></td>
            </tr>`n"
        }
    }

    # Build SSL rows
    $sslSection = ''
    if ($SslResults -and $SslResults.Count -gt 0) {
        $sslRows = ''
        foreach ($r in $SslResults) {
            $badgeClass = switch ($r.Status) {
                'Expired'  { 'badge-crit' }
                'Critical' { 'badge-crit' }
                'Warning'  { 'badge-warn' }
                'Error'    { 'badge-crit' }
                default    { 'badge-ok'   }
            }

            if ($r.Status -eq 'Error') {
                $sslRows += "            <tr>
                <td><strong>$(HtmlEncode $r.Hostname)</strong>:$($r.Port)</td>
                <td style='font-size:12px;color:#e74c3c;'>$(HtmlEncode $r.Error)</td>
                <td>N/A</td>
                <td>N/A</td>
                <td><span class='badge badge-crit'>Error</span></td>
            </tr>`n"
            } else {
                $expiryStr = if ($r.Expiry) { $r.Expiry.ToString('yyyy-MM-dd HH:mm') } else { 'N/A' }
                $daysStr   = if ($null -ne $r.DaysLeft) { "$($r.DaysLeft)" } else { 'N/A' }
                $sslRows += "            <tr>
                <td><strong>$(HtmlEncode $r.Hostname)</strong>:$($r.Port)</td>
                <td style='font-size:12px;word-break:break-all;'>$(HtmlEncode $r.Subject)</td>
                <td>$expiryStr</td>
                <td>$daysStr</td>
                <td><span class='badge $badgeClass'>$(HtmlEncode $r.Status)</span></td>
            </tr>`n"
            }
        }

        $sslSection = @"

<div class="section-title">SSL/TLS Remote Certificate Checks</div>
<table>
  <thead>
    <tr>
      <th>Host:Port</th>
      <th>Subject</th>
      <th>Expiry</th>
      <th>Days Left</th>
      <th>Status</th>
    </tr>
  </thead>
  <tbody>
$sslRows
  </tbody>
</table>
"@
    }

    $orgLine = ''
    $cfg = Get-TKConfig
    if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) {
        $orgLine = "$(HtmlEncode $cfg.OrgName) &nbsp;|&nbsp; "
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>R.E.L.I.C. Certificate Audit — $MachineName</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: #1a1a2e; color: #e0e0e0; font-family: 'Segoe UI', Consolas, monospace; font-size: 14px; padding: 24px; }
  h1 { color: #00d4ff; font-size: 22px; margin-bottom: 4px; }
  .subtitle { color: #888; font-size: 13px; margin-bottom: 24px; }
  .summary { display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 28px; }
  .card { background: #16213e; border: 1px solid #0f3460; border-radius: 8px; padding: 16px 24px; min-width: 120px; text-align: center; }
  .card .val { font-size: 28px; font-weight: bold; color: #00d4ff; }
  .card .lbl { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 1px; margin-top: 4px; }
  .card.warn .val { color: #f39c12; }
  .card.crit .val { color: #e74c3c; }
  .card.ok   .val { color: #2ecc71; }
  table { width: 100%; border-collapse: collapse; margin-top: 8px; }
  th { background: #0f3460; color: #00d4ff; padding: 10px 12px; text-align: left; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; }
  td { padding: 9px 12px; border-bottom: 1px solid #1e2d4d; vertical-align: top; }
  tr:hover td { background: #1e2d4d; }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: bold; }
  .badge-ok   { background: #1a4a2e; color: #2ecc71; }
  .badge-warn { background: #4a3a10; color: #f39c12; }
  .badge-crit { background: #4a1a1a; color: #e74c3c; }
  .section-title { color: #00d4ff; font-size: 15px; margin: 28px 0 10px; border-bottom: 1px solid #0f3460; padding-bottom: 6px; }
  .footer { margin-top: 32px; color: #555; font-size: 11px; }
</style>
</head>
<body>
<h1>R.E.L.I.C. — Certificate Audit Report</h1>
<div class="subtitle">${orgLine}Machine: <strong>$MachineName</strong> &nbsp;|&nbsp; Generated: $ReportTimestamp</div>

<div class="summary">
  <div class="card"><div class="val">$totalCerts</div><div class="lbl">Total Certs</div></div>
  <div class="card crit"><div class="val">$expiredCount</div><div class="lbl">Expired</div></div>
  <div class="card crit"><div class="val">$critCount</div><div class="lbl">Critical (&lt;30d)</div></div>
  <div class="card warn"><div class="val">$warnCount</div><div class="lbl">Warning (&lt;90d)</div></div>
  <div class="card ok"><div class="val">$healthyCount</div><div class="lbl">Healthy</div></div>
</div>

<div class="section-title">Local Certificate Stores</div>
<table>
  <thead>
    <tr>
      <th>Store</th>
      <th>Subject</th>
      <th>Issuer</th>
      <th>Expiry</th>
      <th>Days Left</th>
      <th>Status</th>
    </tr>
  </thead>
  <tbody>
$localRows
  </tbody>
</table>
$sslSection
<div class="footer">
  Generated by R.E.L.I.C. — Technician Toolkit &nbsp;|&nbsp; Stores audited: My, CA, Root, TrustedPublisher
</div>
</body>
</html>
"@

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# EXPORT HELPER
# ─────────────────────────────────────────────────────────────────────────────

function Save-HtmlReport {
    param(
        [array]$LocalCerts,
        [array]$SslResults
    )

    $machineName     = $env:COMPUTERNAME
    $reportTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    Write-Host "  [*] Building HTML report..." -ForegroundColor $C.Progress

    $html = Build-HtmlReport `
        -LocalCerts      $LocalCerts `
        -SslResults      $SslResults `
        -MachineName     $machineName `
        -ReportTimestamp $reportTimestamp

    $reportFilename = "RELIC_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $reportPath     = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) $reportFilename

    try {
        [System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
        Write-Host "  [+] Report saved: $reportPath" -ForegroundColor $C.Success
    } catch {
        Write-Host "  [-] Could not save report: $_" -ForegroundColor $C.Error
        $reportPath = $null
    }

    return $reportPath
}

# ─────────────────────────────────────────────────────────────────────────────
# UNATTENDED MODE
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended) {
    Show-RelicBanner

    $machineName = $env:COMPUTERNAME
    Write-Host "  [*] Unattended mode — Machine: $machineName" -ForegroundColor $C.Progress
    Write-Host ""

    # Local cert audit
    Write-Host "  [*] Auditing local certificate stores..." -ForegroundColor $C.Progress
    $localCerts = Get-LocalCertHealth

    $expiredCount = ($localCerts | Where-Object { $_.Status -eq 'Expired'  } | Measure-Object).Count
    $critCount    = ($localCerts | Where-Object { $_.Status -eq 'Critical' } | Measure-Object).Count
    $warnCount    = ($localCerts | Where-Object { $_.Status -eq 'Warning'  } | Measure-Object).Count
    $healthyCount = ($localCerts | Where-Object { $_.Status -eq 'Healthy'  } | Measure-Object).Count

    Write-Host ("  [+] Found {0} certificate(s): Expired={1}, Critical={2}, Warning={3}, Healthy={4}" -f `
        $localCerts.Count, $expiredCount, $critCount, $warnCount, $healthyCount) -ForegroundColor $C.Success

    # SSL checks
    $sslResults = @()
    if (-not [string]::IsNullOrWhiteSpace($Targets)) {
        $targetList = ConvertTo-TargetList -TargetString $Targets
        if ($targetList.Count -gt 0) {
            Write-Host ""
            Write-Host "  [*] Checking SSL/TLS certificates on $($targetList.Count) remote host(s)..." -ForegroundColor $C.Progress
            foreach ($t in $targetList) {
                Write-Host "  [*] Checking $($t.Hostname):$($t.Port)..." -ForegroundColor $C.Progress
                $sslResults += Get-SslCertExpiry -Hostname $t.Hostname -Port $t.Port
            }
        }
    }

    # Generate HTML report
    Write-Host ""
    $reportPath = Save-HtmlReport -LocalCerts $localCerts -SslResults $sslResults

    # Console summary
    Write-Host ""
    Write-Host ("  " + ("═" * 62)) -ForegroundColor $C.Header
    Write-Host "  R.E.L.I.C. AUDIT SUMMARY" -ForegroundColor $C.Header
    Write-Host ("  " + ("═" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  Total Certificates : $($localCerts.Count)" -ForegroundColor $C.Info
    Write-Host "  Expired            : $expiredCount" -ForegroundColor $(if ($expiredCount -gt 0) { $C.Error } else { $C.Info })
    Write-Host "  Critical (<30d)    : $critCount"   -ForegroundColor $(if ($critCount    -gt 0) { $C.Error } else { $C.Info })
    Write-Host "  Warning  (<90d)    : $warnCount"   -ForegroundColor $(if ($warnCount    -gt 0) { $C.Warning } else { $C.Info })
    Write-Host "  Healthy            : $healthyCount" -ForegroundColor $C.Success
    if ($sslResults.Count -gt 0) {
        $sslErrors  = ($sslResults | Where-Object { $_.Status -eq 'Error'    } | Measure-Object).Count
        $sslExpired = ($sslResults | Where-Object { $_.Status -eq 'Expired'  } | Measure-Object).Count
        $sslCrit    = ($sslResults | Where-Object { $_.Status -eq 'Critical' } | Measure-Object).Count
        Write-Host ""
        Write-Host "  SSL Hosts Checked  : $($sslResults.Count)" -ForegroundColor $C.Info
        Write-Host "  SSL Errors         : $sslErrors"  -ForegroundColor $(if ($sslErrors  -gt 0) { $C.Error   } else { $C.Info })
        Write-Host "  SSL Expired        : $sslExpired" -ForegroundColor $(if ($sslExpired -gt 0) { $C.Error   } else { $C.Info })
        Write-Host "  SSL Critical       : $sslCrit"    -ForegroundColor $(if ($sslCrit    -gt 0) { $C.Error   } else { $C.Info })
    }
    Write-Host ""
    Write-Host ("  " + ("═" * 62)) -ForegroundColor $C.Header
    Write-Host "  R.E.L.I.C. UNATTENDED RUN COMPLETE" -ForegroundColor $C.Header
    Write-Host ("  " + ("═" * 62)) -ForegroundColor $C.Header
    Write-Host ""

    if ($Transcript) { Stop-TKTranscript }
    if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
    exit 0
}

# ─────────────────────────────────────────────────────────────────────────────
# INTERACTIVE MENU
# ─────────────────────────────────────────────────────────────────────────────

Show-RelicBanner

$choice = ''

do {
    Show-RelicBanner

    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host "  MAIN MENU" -ForegroundColor $C.Header
    Write-Host ("  " + ("─" * 62)) -ForegroundColor $C.Header
    Write-Host ""
    Write-Host "  [1] Audit local certificate stores" -ForegroundColor $C.Info
    Write-Host "  [2] Check SSL certificate expiry on remote hosts" -ForegroundColor $C.Info
    Write-Host "  [3] Full audit — local stores + SSL + HTML report" -ForegroundColor $C.Info
    Write-Host "  [Q] Quit" -ForegroundColor $C.Info
    Write-Host ""
    Write-Host -NoNewline "  Enter selection: " -ForegroundColor $C.Header
    $choice = (Read-Host).Trim().ToUpper()

    switch ($choice) {

        # ── Option 1: Local cert audit ──────────────────────────────────────
        '1' {
            Show-RelicBanner
            Write-Host "  [*] Auditing local certificate stores (My, CA, Root, TrustedPublisher)..." -ForegroundColor $C.Progress
            $localCerts = Get-LocalCertHealth
            Write-Host "  [+] Found $($localCerts.Count) certificate(s)." -ForegroundColor $C.Success

            Show-LocalCertSummary -Certs $localCerts

            Write-Host -NoNewline "  Export HTML report? (Y/N): " -ForegroundColor $C.Header
            $exportChoice = (Read-Host).Trim().ToUpper()
            if ($exportChoice -eq 'Y') {
                $reportPath = Save-HtmlReport -LocalCerts $localCerts -SslResults @()
                if ($reportPath) {
                    Write-Host ""
                    Write-Host -NoNewline "  Open report in browser? (Y/N): " -ForegroundColor $C.Header
                    $openChoice = (Read-Host).Trim().ToUpper()
                    if ($openChoice -eq 'Y') { Start-Process $reportPath }
                }
            }

            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

        # ── Option 2: SSL remote checks ─────────────────────────────────────
        '2' {
            Show-RelicBanner
            Write-Section "SSL/TLS REMOTE CERTIFICATE CHECK"

            Write-Host "  Enter target hosts to check." -ForegroundColor $C.Info
            Write-Host "  Format: hostname or hostname:port (comma-separated, or a file path)" -ForegroundColor $C.Info
            Write-Host "  Examples: server1.contoso.com, mail.contoso.com:443, 10.0.0.5:8443" -ForegroundColor $C.Info
            Write-Host ""
            Write-Host -NoNewline "  Targets: " -ForegroundColor $C.Header
            $rawTargets = (Read-Host).Trim()

            if ([string]::IsNullOrWhiteSpace($rawTargets)) {
                Write-Host "  [!!] No targets entered." -ForegroundColor $C.Warning
            } else {
                $targetList = ConvertTo-TargetList -TargetString $rawTargets
                if ($targetList.Count -eq 0) {
                    Write-Host "  [!!] No valid targets parsed." -ForegroundColor $C.Warning
                } else {
                    $sslResults = @()
                    Write-Host ""
                    foreach ($t in $targetList) {
                        Write-Host "  [*] Checking $($t.Hostname):$($t.Port)..." -ForegroundColor $C.Progress
                        $sslResults += Get-SslCertExpiry -Hostname $t.Hostname -Port $t.Port
                    }

                    Show-SslSummary -SslResults $sslResults

                    Write-Host -NoNewline "  Export HTML report? (Y/N): " -ForegroundColor $C.Header
                    $exportChoice = (Read-Host).Trim().ToUpper()
                    if ($exportChoice -eq 'Y') {
                        $reportPath = Save-HtmlReport -LocalCerts @() -SslResults $sslResults
                        if ($reportPath) {
                            Write-Host ""
                            Write-Host -NoNewline "  Open report in browser? (Y/N): " -ForegroundColor $C.Header
                            $openChoice = (Read-Host).Trim().ToUpper()
                            if ($openChoice -eq 'Y') { Start-Process $reportPath }
                        }
                    }
                }
            }

            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

        # ── Option 3: Full audit ────────────────────────────────────────────
        '3' {
            Show-RelicBanner

            # Local cert audit
            Write-Host "  [*] Auditing local certificate stores..." -ForegroundColor $C.Progress
            $localCerts = Get-LocalCertHealth
            Write-Host "  [+] Found $($localCerts.Count) certificate(s)." -ForegroundColor $C.Success

            # SSL targets
            Write-Host ""
            Write-Host "  Enter SSL/TLS target hosts (optional — press Enter to skip)." -ForegroundColor $C.Info
            Write-Host "  Format: hostname or hostname:port (comma-separated, or a file path)" -ForegroundColor $C.Info
            Write-Host ""
            Write-Host -NoNewline "  Targets: " -ForegroundColor $C.Header
            $rawTargets = (Read-Host).Trim()

            $sslResults = @()
            if (-not [string]::IsNullOrWhiteSpace($rawTargets)) {
                $targetList = ConvertTo-TargetList -TargetString $rawTargets
                if ($targetList.Count -gt 0) {
                    Write-Host ""
                    foreach ($t in $targetList) {
                        Write-Host "  [*] Checking $($t.Hostname):$($t.Port)..." -ForegroundColor $C.Progress
                        $sslResults += Get-SslCertExpiry -Hostname $t.Hostname -Port $t.Port
                    }
                }
            }

            # Show console summaries
            Show-LocalCertSummary -Certs $localCerts
            if ($sslResults.Count -gt 0) {
                Show-SslSummary -SslResults $sslResults
            }

            # Generate report automatically
            $reportPath = Save-HtmlReport -LocalCerts $localCerts -SslResults $sslResults

            if ($reportPath) {
                Write-Host ""
                Write-Host -NoNewline "  Open report in browser? (Y/N): " -ForegroundColor $C.Header
                $openChoice = (Read-Host).Trim().ToUpper()
                if ($openChoice -eq 'Y') { Start-Process $reportPath }
            }

            Write-Host ""
            Write-Host -NoNewline "  Press Enter to return to menu..." -ForegroundColor $C.Info
            Read-Host | Out-Null
        }

        # ── Quit ─────────────────────────────────────────────────────────────
        'Q' {
            Write-Host ""
            Write-Host "  Closing R.E.L.I.C." -ForegroundColor $C.Header
            Write-Host ""
        }

        default {
            Write-Host ""
            Write-Host "  [!!] Invalid selection. Enter 1, 2, 3, or Q." -ForegroundColor $C.Warning
            Start-Sleep -Seconds 1
        }
    }

} while ($choice -ne 'Q')

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
