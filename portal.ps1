<#
.SYNOPSIS
    P.O.R.T.A.L. — Profiles, Observes & Reports Tunnels, Authentication & Links
    VPN / Always-On VPN Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Audits the VPN posture of a Windows endpoint. Enumerates built-in
    Windows VPN connections at user and all-user scope via
    `Get-VpnConnection`, surfaces Always-On VPN app triggers via
    `Get-VpnConnectionTrigger`, dumps the Name Resolution Policy Table
    (NRPT) via `Get-DnsClientNrptPolicy`, lists VPN tunnel interfaces,
    and detects installed third-party VPN clients via the service
    table (Cisco AnyConnect / Cisco Secure Client, Palo Alto
    GlobalProtect, Ivanti / Pulse, OpenVPN, WireGuard, Tailscale,
    ZeroTier, Cloudflare WARP, NordVPN, ProtonVPN).

    Read-only audit -- no state-changing actions are performed.

.USAGE
    PS C:\> .\portal.ps1                    # Interactive run
    PS C:\> .\portal.ps1 -Unattended        # Silent: export HTML and exit

.NOTES
    Version : 3.6

#>

param(
    [switch]$Unattended,
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

$C = @{
    Header   = 'Cyan'
    Success  = 'Green'
    Warning  = 'Yellow'
    Error    = 'Red'
    Info     = 'Gray'
    Progress = 'Magenta'
}

function Show-PortalBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  P.O.R.T.A.L. -- Profiles, Observes & Reports Tunnels, Authentication & Links" -ForegroundColor Green
    Write-Host "  VPN / Always-On VPN Audit  v3.6" -ForegroundColor Green
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

# Authentication-method strength tier. PAP transmits credentials in cleartext;
# CHAP and MS-CHAPv1 are deprecated. MS-CHAPv2 is acceptable for non-EAP VPNs.
# EAP and MachineCertificate are the strong tier.
$AuthStrength = @{
    'Pap'                = 'Insecure'
    'Chap'               = 'Weak'
    'MSChapv2'           = 'Acceptable'
    'Eap'                = 'Strong'
    'MachineCertificate' = 'Strong'
}

# Encryption level strength. NoEncryption and Optional are red flags; Required
# and Maximum are the safe tiers.
$EncryptionStrength = @{
    'NoEncryption' = 'Insecure'
    'Optional'     = 'Weak'
    'Required'     = 'Strong'
    'Maximum'      = 'Strong'
}

# Curated catalog of third-party VPN client signals. Each entry is matched by
# service-name pattern (Get-Service -Name) -- absent services are silently
# skipped, present services are surfaced in the report. Patterns ending in '*'
# are wildcarded (e.g. WireGuardTunnel$<name>). The friendly label is what
# appears in the report; vendor lets the technician spot which org owns the
# client when the service name is opaque.
$ThirdPartyClients = @(
    @{ Pattern = 'vpnagent';                Friendly = 'Cisco AnyConnect / Secure Client'; Vendor = 'Cisco' }
    @{ Pattern = 'csc_vpnagent';            Friendly = 'Cisco Secure Client (newer)';      Vendor = 'Cisco' }
    @{ Pattern = 'PanGPS';                  Friendly = 'GlobalProtect (service)';          Vendor = 'Palo Alto' }
    @{ Pattern = 'PanGPA';                  Friendly = 'GlobalProtect (agent)';            Vendor = 'Palo Alto' }
    @{ Pattern = 'JuniperNetworksTunnelService'; Friendly = 'Juniper / Pulse Tunnel';      Vendor = 'Ivanti' }
    @{ Pattern = 'PulseService';            Friendly = 'Pulse Secure Service';             Vendor = 'Ivanti' }
    @{ Pattern = 'OpenVPNService';          Friendly = 'OpenVPN (interactive service)';    Vendor = 'OpenVPN' }
    @{ Pattern = 'OpenVPNServiceLegacy';    Friendly = 'OpenVPN (legacy)';                 Vendor = 'OpenVPN' }
    @{ Pattern = 'OpenVPNServiceInteractive'; Friendly = 'OpenVPN GUI';                    Vendor = 'OpenVPN' }
    @{ Pattern = 'WireGuardManager';        Friendly = 'WireGuard manager';                Vendor = 'WireGuard' }
    @{ Pattern = 'WireGuardTunnel*';        Friendly = 'WireGuard tunnel (per-config)';    Vendor = 'WireGuard' }
    @{ Pattern = 'Tailscale';               Friendly = 'Tailscale';                        Vendor = 'Tailscale' }
    @{ Pattern = 'ZeroTierOneService';      Friendly = 'ZeroTier One';                     Vendor = 'ZeroTier' }
    @{ Pattern = 'CloudflareWARP';          Friendly = 'Cloudflare WARP';                  Vendor = 'Cloudflare' }
    @{ Pattern = 'nordvpn-service';         Friendly = 'NordVPN';                          Vendor = 'Nord Security' }
    @{ Pattern = 'ProtonVPNService';        Friendly = 'Proton VPN';                       Vendor = 'Proton AG' }
    @{ Pattern = 'F5 BIG-IP Edge Client';   Friendly = 'F5 BIG-IP Edge Client';            Vendor = 'F5' }
    @{ Pattern = 'CitrixWorkspaceUpdater';  Friendly = 'Citrix Workspace (heuristic)';     Vendor = 'Citrix' }
)

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function _vpnRowFromCmdlet {
    param($V, [string]$Scope)

    # AuthenticationMethod is an array; encode as a comma-separated string for
    # display while keeping the original for verdict logic.
    $authArr = @()
    if ($V.AuthenticationMethod) { $authArr = @($V.AuthenticationMethod | ForEach-Object { "$_" }) }

    return [PSCustomObject]@{
        Name                  = $V.Name
        Scope                 = $Scope
        ServerAddress         = $V.ServerAddress
        TunnelType            = "$($V.TunnelType)"
        AuthMethods           = $authArr
        AuthMethodsDisplay    = ($authArr -join ', ')
        EncryptionLevel       = "$($V.EncryptionLevel)"
        L2tpIPsecAuth         = "$($V.L2tpIPsecAuth)"
        UseWinlogonCredential = [bool]$V.UseWinlogonCredential
        ConnectionStatus      = "$($V.ConnectionStatus)"
        RememberCredential    = [bool]$V.RememberCredential
        SplitTunneling        = [bool]$V.SplitTunneling
        DnsSuffix             = $V.DnsSuffix
        IdleDisconnectSeconds = $V.IdleDisconnectSeconds
        ProfileType           = "$($V.ProfileType)"
    }
}

function Get-VpnConnections {
    $rows = New-Object 'System.Collections.Generic.List[object]'
    try {
        $user = @(Get-VpnConnection -ErrorAction Stop)
        foreach ($v in $user) { $rows.Add((_vpnRowFromCmdlet -V $v -Scope 'User')) }
    } catch {
        Write-TKError -ScriptName 'portal.ps1' -Message "Get-VpnConnection (user scope) failed: $($_.Exception.Message)" -Category 'VPN'
    }
    try {
        $all = @(Get-VpnConnection -AllUserConnection -ErrorAction Stop)
        foreach ($v in $all) { $rows.Add((_vpnRowFromCmdlet -V $v -Scope 'AllUser')) }
    } catch {
        Write-TKError -ScriptName 'portal.ps1' -Message "Get-VpnConnection (all-user scope) failed: $($_.Exception.Message)" -Category 'VPN'
    }
    return @($rows)
}

function Get-AlwaysOnTriggers {
    param([array]$Connections)

    $rows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($c in $Connections) {
        try {
            $params = @{ ConnectionName = $c.Name; ErrorAction = 'Stop' }
            if ($c.Scope -eq 'AllUser') { $params.AllUserConnection = $true }
            $trigger = Get-VpnConnectionTrigger @params
            if (-not $trigger) { continue }

            $apps = @()
            if ($trigger.ApplicationID) {
                $apps = @($trigger.ApplicationID | Where-Object { $_ })
            }
            $dns = @()
            if ($trigger.DnsConfiguration) {
                $dns = @($trigger.DnsConfiguration | ForEach-Object {
                    "$($_.DnsSuffix) -> $(($_.DnsServers -join ', '))"
                })
            }
            if ($apps.Count -gt 0 -or $dns.Count -gt 0) {
                $rows.Add([PSCustomObject]@{
                    ConnectionName = $c.Name
                    Scope          = $c.Scope
                    Apps           = $apps
                    DnsRules       = $dns
                })
            }
        } catch {
            # Connection has no trigger -- expected for plain manual VPNs.
            continue
        }
    }
    return @($rows)
}

function Get-NrptEntries {
    try {
        $rules = @(Get-DnsClientNrptPolicy -ErrorAction Stop)
    } catch {
        return @()
    }
    return @($rules | ForEach-Object {
        [PSCustomObject]@{
            Namespace          = ($_.Namespace -join ', ')
            DnsServers         = ($_.NameServers -join ', ')
            DirectAccessServers= (@($_.DirectAccessDnsServers) -join ', ')
            IpsecRequired      = [bool]$_.IPsecRequired
            EncryptionType     = "$($_.DnsSecValidationRequired)"
            Comment            = $_.Comment
        }
    })
}

function Get-VpnTunnelInterfaces {
    try {
        # Tunnel interfaces show up with InterfaceAlias starting with the VPN
        # connection name OR with PPP* / WireGuardTunnel*. Filter by both
        # ConnectionState (Connected) and a heuristic name match so we don't
        # pollute the table with every IPv6 tunnel pseudo-interface.
        $ifaces = @(Get-NetIPInterface -ErrorAction Stop | Where-Object {
            $_.ConnectionState -eq 'Connected' -and (
                $_.InterfaceAlias -match 'VPN|Wintun|WireGuard|Tailscale|GlobalProtect|AnyConnect|PPP'
            )
        })
    } catch {
        return @()
    }
    return @($ifaces | ForEach-Object {
        [PSCustomObject]@{
            InterfaceAlias  = $_.InterfaceAlias
            AddressFamily   = "$($_.AddressFamily)"
            ConnectionState = "$($_.ConnectionState)"
            Forwarding      = "$($_.Forwarding)"
            Dhcp            = "$($_.Dhcp)"
            InterfaceMetric = $_.InterfaceMetric
        }
    })
}

function Get-ThirdPartyVpnClients {
    $rows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($entry in $ThirdPartyClients) {
        try {
            $svcs = @(Get-Service -Name $entry.Pattern -ErrorAction SilentlyContinue)
        } catch { $svcs = @() }
        foreach ($s in $svcs) {
            $rows.Add([PSCustomObject]@{
                Name      = $s.Name
                Friendly  = $entry.Friendly
                Vendor    = $entry.Vendor
                Status    = "$($s.Status)"
                StartType = "$($s.StartType)"
            })
        }
    }
    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# CLASSIFICATION
# ─────────────────────────────────────────────────────────────────────────────

function Get-VpnAuthTier {
    param([array]$Methods)
    if (-not $Methods -or $Methods.Count -eq 0) { return 'Unknown' }
    # The weakest method in the configured list governs the tier -- a VPN that
    # accepts PAP as one of multiple methods is still PAP-vulnerable.
    $tier = 'Strong'
    foreach ($m in $Methods) {
        $t = if ($AuthStrength.ContainsKey("$m")) { $AuthStrength["$m"] } else { 'Unknown' }
        switch ($t) {
            'Insecure'   { return 'Insecure' }
            'Weak'       { if ($tier -ne 'Insecure') { $tier = 'Weak' } }
            'Acceptable' { if ($tier -in @('Strong')) { $tier = 'Acceptable' } }
            'Unknown'    { if ($tier -eq 'Strong') { $tier = 'Unknown' } }
        }
    }
    return $tier
}

function Get-VpnEncryptionTier {
    param([string]$Level)
    if ($EncryptionStrength.ContainsKey($Level)) { return $EncryptionStrength[$Level] }
    return 'Unknown'
}

# ─────────────────────────────────────────────────────────────────────────────
# VERDICT
# ─────────────────────────────────────────────────────────────────────────────

function Get-PortalVerdict {
    param([array]$Vpns, [array]$Triggers, [array]$Nrpt, [array]$ThirdParty)

    $issues = New-Object 'System.Collections.Generic.List[string]'
    $warns  = New-Object 'System.Collections.Generic.List[string]'

    if ($Vpns.Count -eq 0 -and $ThirdParty.Count -eq 0) {
        return [PSCustomObject]@{
            Verdict = 'NONE CONFIGURED'
            Class   = 'info'
            Issues  = @()
            Warns   = @('No built-in Windows VPN connections and no third-party VPN clients detected. This is a posture statement, not a finding.')
        }
    }

    foreach ($v in $Vpns) {
        $authTier = Get-VpnAuthTier -Methods $v.AuthMethods
        $encTier  = Get-VpnEncryptionTier -Level $v.EncryptionLevel

        if ($authTier -eq 'Insecure') {
            $msg = "VPN '$($v.Name)' accepts PAP authentication -- credentials transmitted in cleartext."
            $issues.Add($msg)
            Write-TKError -ScriptName 'portal.ps1' -Message $msg -Category 'VPN'
        } elseif ($authTier -eq 'Weak') {
            $warns.Add("VPN '$($v.Name)' accepts CHAP -- deprecated; require MS-CHAPv2 or EAP.")
        }

        if ($encTier -eq 'Insecure') {
            $msg = "VPN '$($v.Name)' uses NoEncryption -- traffic flows in the clear over the tunnel."
            $issues.Add($msg)
            Write-TKError -ScriptName 'portal.ps1' -Message $msg -Category 'VPN'
        } elseif ($encTier -eq 'Weak') {
            $warns.Add("VPN '$($v.Name)' has Optional encryption -- raise to Required or Maximum.")
        }

        if ($v.SplitTunneling) {
            $hasNrptForThis = @($Nrpt | Where-Object { $_.Namespace -and $v.DnsSuffix -and ($_.Namespace -match [regex]::Escape($v.DnsSuffix)) }).Count -gt 0
            if (-not $hasNrptForThis) {
                $warns.Add("VPN '$($v.Name)' is split-tunnel with no NRPT entry covering its DNS suffix '$($v.DnsSuffix)' -- internal-name resolution may leak to the public resolver.")
            }
        }

        if ($v.Scope -eq 'AllUser' -and $v.ConnectionStatus -eq 'Disconnected') {
            $warns.Add("All-user VPN '$($v.Name)' is currently Disconnected -- expected if Always-On has not triggered yet, but worth confirming.")
        }
    }

    # Concurrent third-party clients -> routing-conflict risk.
    $vendors = @($ThirdParty | Select-Object -ExpandProperty Vendor -Unique)
    if ($vendors.Count -gt 1) {
        $warns.Add("$($vendors.Count) different VPN vendors detected on this machine ($([string]::Join(', ', $vendors))) -- competing route tables and DNS resolvers can produce intermittent connectivity.")
    }

    $verdict = if ($issues.Count -gt 0) { 'AT RISK' }
               elseif ($warns.Count -gt 0) { 'ATTENTION NEEDED' }
               else { 'OK' }
    $class   = if ($issues.Count -gt 0) { 'err' }
               elseif ($warns.Count -gt 0) { 'warn' }
               else { 'ok' }

    return [PSCustomObject]@{
        Verdict = $verdict
        Class   = $class
        Issues  = @($issues)
        Warns   = @($warns)
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-PortalReport {
    param([array]$Vpns, [array]$Triggers, [array]$Nrpt, [array]$Tunnels, [array]$ThirdParty, $Verdict)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    function _tierBadge { param([string]$Tier)
        switch ($Tier) {
            'Strong'     { "<span class='tk-badge-ok'>$(EscHtml $Tier)</span>" }
            'Acceptable' { "<span class='tk-badge-info'>$(EscHtml $Tier)</span>" }
            'Weak'       { "<span class='tk-badge-warn'>$(EscHtml $Tier)</span>" }
            'Insecure'   { "<span class='tk-badge-err'>$(EscHtml $Tier)</span>" }
            default      { "<span class='tk-badge-info'>$(EscHtml $Tier)</span>" }
        }
    }
    function _yn { param($b) if ($b) { "<span class='tk-badge-ok'>Yes</span>" } else { "<span class='tk-badge-info'>No</span>" } }
    function _ynWarn { param($b) if ($b) { "<span class='tk-badge-warn'>Yes</span>" } else { "<span class='tk-badge-ok'>No</span>" } }

    # VPN connections table
    $vRows = [System.Text.StringBuilder]::new()
    if ($Vpns.Count -eq 0) {
        [void]$vRows.Append("<tr><td colspan='9' class='tk-badge-info' style='text-align:center;'>No built-in Windows VPN connections configured.</td></tr>")
    } else {
        foreach ($v in $Vpns | Sort-Object Scope, Name) {
            $authTier = Get-VpnAuthTier -Methods $v.AuthMethods
            $encTier  = Get-VpnEncryptionTier -Level $v.EncryptionLevel
            $stateBadge = switch ($v.ConnectionStatus) {
                'Connected'    { "<span class='tk-badge-ok'>Connected</span>" }
                'Disconnected' { "<span class='tk-badge-info'>Disconnected</span>" }
                default        { "<span class='tk-badge-warn'>$(EscHtml $v.ConnectionStatus)</span>" }
            }
            [void]$vRows.Append(
                "<tr><td>$(EscHtml $v.Name)</td><td>$(EscHtml $v.Scope)</td>" +
                "<td class='tk-mono'>$(EscHtml $v.ServerAddress)</td>" +
                "<td>$(EscHtml $v.TunnelType)</td>" +
                "<td>$(EscHtml $v.AuthMethodsDisplay) $(_tierBadge $authTier)</td>" +
                "<td>$(EscHtml $v.EncryptionLevel) $(_tierBadge $encTier)</td>" +
                "<td>$(_ynWarn $v.SplitTunneling)</td>" +
                "<td>$stateBadge</td>" +
                "<td>$(EscHtml $v.ProfileType)</td></tr>"
            )
        }
    }

    # Always-On triggers
    $tRows = [System.Text.StringBuilder]::new()
    if ($Triggers.Count -eq 0) {
        [void]$tRows.Append("<tr><td colspan='4' class='tk-badge-info' style='text-align:center;'>No app-triggered Always-On entries configured.</td></tr>")
    } else {
        foreach ($t in $Triggers) {
            $appList = if ($t.Apps.Count -gt 0) { ($t.Apps | ForEach-Object { "<code class='tk-mono'>$(EscHtml $_)</code>" }) -join '<br/>' } else { '<span class=''tk-badge-info''>none</span>' }
            $dnsList = if ($t.DnsRules.Count -gt 0) { ($t.DnsRules | ForEach-Object { "<code class='tk-mono'>$(EscHtml $_)</code>" }) -join '<br/>' } else { '<span class=''tk-badge-info''>none</span>' }
            [void]$tRows.Append(
                "<tr><td>$(EscHtml $t.ConnectionName)</td><td>$(EscHtml $t.Scope)</td>" +
                "<td>$appList</td><td>$dnsList</td></tr>"
            )
        }
    }

    # NRPT
    $nRows = [System.Text.StringBuilder]::new()
    if ($Nrpt.Count -eq 0) {
        [void]$nRows.Append("<tr><td colspan='5' class='tk-badge-info' style='text-align:center;'>No NRPT rules configured.</td></tr>")
    } else {
        foreach ($n in $Nrpt) {
            [void]$nRows.Append(
                "<tr><td class='tk-mono'>$(EscHtml $n.Namespace)</td>" +
                "<td class='tk-mono'>$(EscHtml $n.DnsServers)</td>" +
                "<td class='tk-mono'>$(EscHtml $n.DirectAccessServers)</td>" +
                "<td>$(_ynWarn $n.IpsecRequired)</td>" +
                "<td>$(EscHtml $n.Comment)</td></tr>"
            )
        }
    }

    # Tunnel interfaces
    $iRows = [System.Text.StringBuilder]::new()
    if ($Tunnels.Count -eq 0) {
        [void]$iRows.Append("<tr><td colspan='5' class='tk-badge-info' style='text-align:center;'>No active VPN tunnel interfaces detected.</td></tr>")
    } else {
        foreach ($i in $Tunnels) {
            [void]$iRows.Append(
                "<tr><td>$(EscHtml $i.InterfaceAlias)</td><td>$(EscHtml $i.AddressFamily)</td>" +
                "<td>$(EscHtml $i.ConnectionState)</td><td>$(EscHtml $i.Dhcp)</td>" +
                "<td>$(EscHtml $i.InterfaceMetric)</td></tr>"
            )
        }
    }

    # Third-party clients
    $cRows = [System.Text.StringBuilder]::new()
    if ($ThirdParty.Count -eq 0) {
        [void]$cRows.Append("<tr><td colspan='5' class='tk-badge-info' style='text-align:center;'>No third-party VPN clients detected.</td></tr>")
    } else {
        foreach ($c in $ThirdParty) {
            $statusBadge = if ($c.Status -eq 'Running') { "<span class='tk-badge-ok'>Running</span>" } else { "<span class='tk-badge-info'>$(EscHtml $c.Status)</span>" }
            [void]$cRows.Append(
                "<tr><td>$(EscHtml $c.Friendly)</td><td>$(EscHtml $c.Vendor)</td>" +
                "<td class='tk-mono'>$(EscHtml $c.Name)</td><td>$statusBadge</td>" +
                "<td>$(EscHtml $c.StartType)</td></tr>"
            )
        }
    }

    $findingsList = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$findingsList.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>") }
    foreach ($w in $Verdict.Warns)  { [void]$findingsList.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$findingsList.Append("<li class='tk-badge-ok'>No insecure VPN configurations detected.</li>")
    }

    $allUserCount = @($Vpns | Where-Object { $_.Scope -eq 'AllUser' }).Count
    $userCount    = @($Vpns | Where-Object { $_.Scope -eq 'User' }).Count
    $insecureCount = @($Vpns | Where-Object {
        (Get-VpnAuthTier -Methods $_.AuthMethods) -eq 'Insecure' -or
        (Get-VpnEncryptionTier -Level $_.EncryptionLevel) -eq 'Insecure'
    }).Count

    $htmlHead = Get-TKHtmlHead `
        -Title      'P.O.R.T.A.L. VPN / Always-On VPN Audit' `
        -ScriptName 'P.O.R.T.A.L.' `
        -Subtitle   "${orgPrefix}VPN / Always-On VPN Audit -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'             = $machine
            'Generated'           = $reportDate
            'Verdict'             = $Verdict.Verdict
            'Built-in VPNs'       = $Vpns.Count
            'Third-Party Clients' = $ThirdParty.Count
        }) `
        -NavItems   @('Verdict', 'Built-in VPNs', 'Always-On Triggers', 'NRPT', 'Tunnel Interfaces', 'Third-Party Clients')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'P.O.R.T.A.L. v3.6'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">VPN Posture</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Vpns.Count)</div><div class="tk-summary-lbl">Built-in VPNs</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$allUserCount</div><div class="tk-summary-lbl">All-user (Always-On candidate)</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Triggers.Count)</div><div class="tk-summary-lbl">App Triggers</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Nrpt.Count)</div><div class="tk-summary-lbl">NRPT Entries</div></div>
    <div class="tk-summary-card $(if ($ThirdParty.Count -gt 0) { 'warn' } else { 'info' })"><div class="tk-summary-num">$($ThirdParty.Count)</div><div class="tk-summary-lbl">Third-Party Clients</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Verdict &amp; Findings</span><span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span></div>
    <div class="tk-card"><ul class="tk-info-box" style="list-style:none;padding-left:0;">$($findingsList.ToString())</ul></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Built-in Windows VPN Connections</span><span class="tk-section-num">$userCount user / $allUserCount all-user</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Name</th><th>Scope</th><th>Server</th><th>Tunnel</th><th>Auth</th><th>Encryption</th><th>Split-Tunnel</th><th>State</th><th>Profile</th></tr></thead>
        <tbody>$($vRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Always-On VPN App Triggers</span><span class="tk-section-num">$($Triggers.Count) configured</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Connection</th><th>Scope</th><th>Triggering Apps</th><th>DNS Triggers</th></tr></thead>
        <tbody>$($tRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">NRPT Entries</span><span class="tk-section-num">$($Nrpt.Count) rules</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Namespace</th><th>DNS Servers</th><th>DirectAccess Servers</th><th>IPsec required</th><th>Comment</th></tr></thead>
        <tbody>$($nRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">VPN Tunnel Interfaces</span><span class="tk-section-num">$($Tunnels.Count) connected</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Interface</th><th>Family</th><th>State</th><th>DHCP</th><th>Metric</th></tr></thead>
        <tbody>$($iRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Third-Party VPN Clients</span><span class="tk-section-num">$($ThirdParty.Count) detected</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Product</th><th>Vendor</th><th>Service</th><th>Status</th><th>Start Type</th></tr></thead>
        <tbody>$($cRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-PortalBanner

Write-Section "BUILT-IN VPN CONNECTIONS"
$vpns = Get-VpnConnections
Write-Host ("  Built-in VPN connections : {0}" -f $vpns.Count) -ForegroundColor $C.Info
foreach ($v in $vpns) {
    $authTier = Get-VpnAuthTier -Methods $v.AuthMethods
    $encTier  = Get-VpnEncryptionTier -Level $v.EncryptionLevel
    $color = if ($authTier -eq 'Insecure' -or $encTier -eq 'Insecure') { $C.Error }
             elseif ($authTier -in @('Weak','Acceptable') -or $encTier -eq 'Weak') { $C.Warning }
             else { $C.Success }
    Write-Host ("  - {0,-25} [{1}] auth={2} enc={3} state={4}" -f $v.Name, $v.Scope, $v.AuthMethodsDisplay, $v.EncryptionLevel, $v.ConnectionStatus) -ForegroundColor $color
}
Write-Host ""

Write-Section "ALWAYS-ON TRIGGERS"
$triggers = Get-AlwaysOnTriggers -Connections $vpns
Write-Host ("  App-trigger / DNS-trigger entries : {0}" -f $triggers.Count) -ForegroundColor $C.Info
Write-Host ""

Write-Section "NRPT ENTRIES"
$nrpt = Get-NrptEntries
Write-Host ("  NRPT rules                : {0}" -f $nrpt.Count) -ForegroundColor $C.Info
Write-Host ""

Write-Section "ACTIVE TUNNEL INTERFACES"
$tunnels = Get-VpnTunnelInterfaces
Write-Host ("  Connected tunnel ifaces   : {0}" -f $tunnels.Count) -ForegroundColor $C.Info
Write-Host ""

Write-Section "THIRD-PARTY VPN CLIENTS"
$thirdParty = Get-ThirdPartyVpnClients
Write-Host ("  Detected client services  : {0}" -f $thirdParty.Count) -ForegroundColor $C.Info
foreach ($c in $thirdParty) {
    $color = if ($c.Status -eq 'Running') { $C.Success } else { $C.Info }
    Write-Host ("    - {0,-30} ({1})  service: {2,-30} {3}" -f $c.Friendly, $c.Vendor, $c.Name, $c.Status) -ForegroundColor $color
}
Write-Host ""

$verdict = Get-PortalVerdict -Vpns $vpns -Triggers $triggers -Nrpt $nrpt -ThirdParty $thirdParty

Write-Section "VPN VERDICT"
$verdictColor = switch ($verdict.Class) {
    'ok'   { $C.Success }
    'warn' { $C.Warning }
    'err'  { $C.Error }
    default { $C.Info }
}
Write-Host "  $($verdict.Verdict)" -ForegroundColor $verdictColor
foreach ($i in $verdict.Issues) { Write-Host "    [!!] $i" -ForegroundColor $C.Error }
foreach ($w in $verdict.Warns)  { Write-Host "    [~ ] $w" -ForegroundColor $C.Warning }
if ($verdict.Class -eq 'ok' -and $verdict.Issues.Count -eq 0 -and $verdict.Warns.Count -eq 0) {
    Write-Host "    [+ ] All checks passed." -ForegroundColor $C.Success
}
Write-Host ""

Write-Step "Generating HTML report..."
$html      = Build-PortalReport -Vpns $vpns -Triggers $triggers -Nrpt $nrpt -Tunnels $tunnels -ThirdParty $thirdParty -Verdict $verdict
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "PORTAL_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
