<#
.SYNOPSIS
    B.E.A.C.O.N. — Broadcasts, Encryption, Authentication & Connections Of Networks
    Wi-Fi Profile Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Audits saved Wi-Fi (WLAN) profiles on a Windows endpoint. Exports
    each profile to XML via `netsh wlan export profile ... key=clear`,
    parses the locale-stable XML schema, and produces a dark-themed
    HTML report covering authentication / encryption / connection mode
    / auto-switch / hidden-SSID / MAC-randomization for every saved
    profile, plus filtered tables for open, weak, and auto-connecting
    networks.

    Read-only audit. Key material is masked by default and only
    rendered in cleartext when `-IncludeKey` is explicitly passed.

.USAGE
    PS C:\> .\beacon.ps1                    # Interactive run; key material masked
    PS C:\> .\beacon.ps1 -Unattended        # Silent: export HTML and exit
    PS C:\> .\beacon.ps1 -IncludeKey        # Render cleartext PSKs in the report (technician-managed audit only)

.NOTES
    Version : 3.6

#>

param(
    [switch]$Unattended,
    [switch]$Transcript,
    [switch]$IncludeKey
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

function Show-BeaconBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  B.E.A.C.O.N. -- Broadcasts, Encryption, Authentication & Connections Of Networks" -ForegroundColor Yellow
    Write-Host "  Wi-Fi Profile Audit  v3.6" -ForegroundColor Yellow
    if ($IncludeKey) {
        Write-Host "  *** Key material WILL be rendered in cleartext (-IncludeKey) ***" -ForegroundColor Magenta
    }
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

# Authentication strength tier. Drives badge colour and verdict.
# Values come from the WLAN profile XML schema:
# learn.microsoft.com/en-us/windows/win32/nativewifi/wlan-profileschema-authentication-authencryption-element
$AuthStrength = @{
    'open'       = 'Insecure'
    'shared'     = 'Insecure'   # Shared-key WEP
    'WEP'        = 'Insecure'
    'WPA'        = 'Weak'
    'WPAPSK'     = 'Weak'
    'WPA2'       = 'Strong'     # WPA2 enterprise
    'WPA2PSK'    = 'Strong'
    'WPA3'       = 'Strong'
    'WPA3SAE'    = 'Strong'
    'WPA3ENT'    = 'Strong'
    'WPA3ENT192' = 'Strong'
    'OWE'        = 'Strong'     # Open with opportunistic encryption
}

# Cipher tier. TKIP and WEP are deprecated; AES (CCMP) and GCMP are current.
$CipherStrength = @{
    'none' = 'Insecure'
    'WEP'  = 'Insecure'
    'TKIP' = 'Weak'
    'AES'  = 'Strong'   # CCMP
    'GCMP' = 'Strong'
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-WifiAdapters {
    # PhysicalMediaType 'Native 802.11' is the canonical filter for Wi-Fi NICs.
    try {
        $rows = Get-NetAdapter -Physical -ErrorAction Stop | Where-Object {
            $_.PhysicalMediaType -eq 'Native 802.11' -or $_.MediaType -eq '802.11'
        } | ForEach-Object {
            [PSCustomObject]@{
                Name              = $_.Name
                InterfaceDescription = $_.InterfaceDescription
                Status            = $_.Status
                LinkSpeed         = $_.LinkSpeed
                MacAddress        = $_.MacAddress
                DriverVersion     = $_.DriverVersion
                DriverDate        = $_.DriverDate
            }
        }
        return @($rows)
    } catch {
        return @()
    }
}

function Get-WlanProfileList {
    # `netsh wlan show profiles` output is locale-dependent for headers but
    # the per-profile lines all have the form `    <something> : <profileName>`.
    # We extract the trailing token (the profile name) from any line containing
    # a colon and indented two-plus spaces. Empty / non-matching lines are skipped.
    $raw = & netsh wlan show profiles 2>$null
    if (-not $raw) { return @() }

    $names = New-Object 'System.Collections.Generic.List[string]'
    foreach ($line in $raw) {
        if ($line -match '^\s+[^:]+:\s+(.+?)\s*$') {
            $candidate = $matches[1].Trim()
            # Skip anything that's clearly a header (contains keywords we never see in real names).
            if ($candidate -and $candidate -notmatch '^(Hosted network|Group policy|User|All User|Granted Permission|Restrictions)') {
                $names.Add($candidate)
            }
        }
    }
    # Names can repeat across User / All-User scopes. De-duplicate but keep order.
    return @($names | Select-Object -Unique)
}

function Export-WlanProfileXml {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [string]$Folder
    )

    # `key=clear` requires elevation. If we're elevated (Invoke-AdminElevation
    # already ensured this) we get cleartext keys; otherwise we get DPAPI blobs
    # we can't decode. Suppress the netsh stdout / stderr noise -- callers only
    # care about the XML file landing on disk.
    $null = & netsh wlan export profile name="$Name" folder="$Folder" key=clear 2>&1
    # netsh emits the file name from the SSID name with characters sanitised:
    # spaces -> hyphens, etc. The simplest reliable approach is to glob the
    # folder for any new XML file. Caller passes a dedicated empty folder per
    # profile, so we just take the first XML there.
    $xml = Get-ChildItem -Path $Folder -Filter '*.xml' -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $xml) { return $null }
    return $xml.FullName
}

function ConvertFrom-WlanProfileXml {
    param([Parameter(Mandatory)] [string]$XmlPath)

    try {
        [xml]$doc = Get-Content -Path $XmlPath -Raw -ErrorAction Stop
    } catch {
        return $null
    }

    # The XML uses a default namespace; XPath queries must declare it. We pluck
    # values by walking the DOM directly to avoid an XmlNamespaceManager dance.
    $profile = $doc.WLANProfile
    if (-not $profile) { return $null }

    $name           = $profile.name
    $ssidName       = $null
    $ssidHex        = $null
    $nonBroadcast   = $false
    if ($profile.SSIDConfig -and $profile.SSIDConfig.SSID) {
        $ssidName = $profile.SSIDConfig.SSID.name
        $ssidHex  = $profile.SSIDConfig.SSID.hex
    }
    if ($profile.SSIDConfig -and $profile.SSIDConfig.nonBroadcast) {
        $nonBroadcast = ($profile.SSIDConfig.nonBroadcast -eq 'true')
    }

    $connectionType = $profile.connectionType
    $connectionMode = $profile.connectionMode
    $autoSwitch     = ($profile.autoSwitch -eq 'true')

    $auth         = $null
    $encryption   = $null
    $useOneX      = $false
    $keyType      = $null
    $keyProtected = $null
    $keyMaterial  = $null
    if ($profile.MSM -and $profile.MSM.security) {
        $sec = $profile.MSM.security
        if ($sec.authEncryption) {
            $auth       = $sec.authEncryption.authentication
            $encryption = $sec.authEncryption.encryption
            $useOneX    = ($sec.authEncryption.useOneX -eq 'true')
        }
        if ($sec.sharedKey) {
            $keyType      = $sec.sharedKey.keyType
            $keyProtected = ($sec.sharedKey.protected -eq 'true')
            $keyMaterial  = $sec.sharedKey.keyMaterial
        }
    }

    $macRand = $null
    # MacRandomization sits in a v3 child namespace. .NET's XmlElement returns
    # the property if present regardless of namespace, but if it's missing the
    # accessor returns $null in PS5.1 (no exception).
    if ($profile.MacRandomization) {
        $macRand = ($profile.MacRandomization.enableRandomization -eq 'true')
    }

    return [PSCustomObject]@{
        Name             = $name
        SsidName         = $ssidName
        SsidHex          = $ssidHex
        NonBroadcast     = $nonBroadcast
        ConnectionType   = $connectionType
        ConnectionMode   = $connectionMode
        AutoSwitch       = $autoSwitch
        Authentication   = $auth
        Encryption       = $encryption
        UseOneX          = $useOneX
        KeyType          = $keyType
        KeyProtected     = $keyProtected
        KeyMaterial      = $keyMaterial
        MacRandomization = $macRand
    }
}

function Get-WlanProfiles {
    $names = Get-WlanProfileList
    if (-not $names -or $names.Count -eq 0) { return @() }

    $rootTemp = Join-Path ([System.IO.Path]::GetTempPath()) ("TK-BEACON-" + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $rootTemp -Force -ErrorAction SilentlyContinue

    $rows = New-Object 'System.Collections.Generic.List[object]'
    try {
        foreach ($n in $names) {
            $sub = Join-Path $rootTemp ([guid]::NewGuid().ToString('N'))
            $null = New-Item -ItemType Directory -Path $sub -Force -ErrorAction SilentlyContinue
            $xmlPath = Export-WlanProfileXml -Name $n -Folder $sub
            if ($null -eq $xmlPath) {
                Write-TKError -ScriptName 'beacon.ps1' -Message "Could not export profile '$n' (netsh produced no XML)." -Category 'Wi-Fi'
                continue
            }
            $row = ConvertFrom-WlanProfileXml -XmlPath $xmlPath
            if ($row) { $rows.Add($row) }
        }
    } finally {
        Remove-Item -Path $rootTemp -Recurse -Force -ErrorAction SilentlyContinue
    }
    return @($rows)
}

# ─────────────────────────────────────────────────────────────────────────────
# CLASSIFICATION
# ─────────────────────────────────────────────────────────────────────────────

function Get-AuthTier {
    param([string]$Auth)
    if ([string]::IsNullOrWhiteSpace($Auth)) { return 'Unknown' }
    if ($AuthStrength.ContainsKey($Auth)) { return $AuthStrength[$Auth] }
    return 'Unknown'
}

function Get-CipherTier {
    param([string]$Cipher)
    if ([string]::IsNullOrWhiteSpace($Cipher)) { return 'Unknown' }
    if ($CipherStrength.ContainsKey($Cipher)) { return $CipherStrength[$Cipher] }
    return 'Unknown'
}

function Test-IsOpenProfile {
    param($P)
    $a = "$($P.Authentication)".ToLower()
    return ($a -eq 'open' -and ("$($P.Encryption)".ToLower() -in @('none', '')))
}

function Test-IsWeakProfile {
    param($P)
    if (Test-IsOpenProfile $P) { return $true }
    $auth = "$($P.Authentication)"
    $cipher = "$($P.Encryption)"
    if ((Get-AuthTier $auth) -eq 'Insecure') { return $true }
    if ((Get-CipherTier $cipher) -eq 'Insecure') { return $true }
    if ((Get-AuthTier $auth) -eq 'Weak')   { return $true }
    if ((Get-CipherTier $cipher) -eq 'Weak') { return $true }
    return $false
}

# ─────────────────────────────────────────────────────────────────────────────
# VERDICT
# ─────────────────────────────────────────────────────────────────────────────

function Get-BeaconVerdict {
    param([array]$Adapters, [array]$Profiles)

    $issues = New-Object 'System.Collections.Generic.List[string]'
    $warns  = New-Object 'System.Collections.Generic.List[string]'

    if ($Adapters.Count -eq 0) {
        # Desktops without Wi-Fi are a normal state, not a finding.
        $warns.Add('No Wi-Fi adapter present on this machine -- desktop / wired-only configuration.')
    }

    foreach ($p in $Profiles) {
        $auth   = "$($p.Authentication)"
        $cipher = "$($p.Encryption)"
        $auto   = ($p.ConnectionMode -eq 'auto')

        if (Test-IsOpenProfile $p) {
            if ($auto) {
                $issues.Add("Open profile '$($p.Name)' is set to AUTO-CONNECT -- this machine will silently associate to any AP advertising that SSID.")
            } else {
                $warns.Add("Open profile '$($p.Name)' is configured (manual-connect).")
            }
        }
        if ((Get-CipherTier $cipher) -eq 'Insecure' -and $cipher -ne 'none') {
            $issues.Add("Profile '$($p.Name)' uses deprecated $cipher encryption -- WEP is broken; remove the profile.")
        }
        if ((Get-AuthTier $auth) -eq 'Weak') {
            $warns.Add("Profile '$($p.Name)' uses legacy $auth -- upgrade the network to WPA2-PSK / WPA3-SAE.")
        }
        if ((Get-CipherTier $cipher) -eq 'Weak') {
            $warns.Add("Profile '$($p.Name)' uses TKIP cipher -- upgrade to AES (CCMP) or GCMP.")
        }
        if ($p.AutoSwitch -and $auto) {
            $warns.Add("Profile '$($p.Name)' has autoSwitch enabled -- machine may roam between this and other known networks without prompt.")
        }
        if ($p.NonBroadcast -and $auto) {
            $warns.Add("Hidden-SSID profile '$($p.Name)' is auto-connecting -- the client probes for it constantly, leaking the SSID.")
        }
        if ($p.MacRandomization -eq $false) {
            $warns.Add("Profile '$($p.Name)' has MAC randomization disabled.")
        }
    }

    if ($Profiles.Count -gt 25) {
        $warns.Add("$($Profiles.Count) saved Wi-Fi profiles -- a large profile estate increases roaming surprises and probe-request leakage. Consider periodic cleanup.")
    }

    $verdict = if ($issues.Count -gt 0) { 'AT RISK' }
               elseif ($warns.Count -gt 0) { 'ATTENTION NEEDED' }
               else { 'CLEAN' }
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

function Build-BeaconReport {
    param([array]$Adapters, [array]$Profiles, $Verdict)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    function _tierBadge { param([string]$Tier)
        switch ($Tier) {
            'Strong'   { "<span class='tk-badge-ok'>$(EscHtml $Tier)</span>" }
            'Weak'     { "<span class='tk-badge-warn'>$(EscHtml $Tier)</span>" }
            'Insecure' { "<span class='tk-badge-err'>$(EscHtml $Tier)</span>" }
            default    { "<span class='tk-badge-info'>$(EscHtml $Tier)</span>" }
        }
    }
    function _yn { param($b) if ($b) { "<span class='tk-badge-ok'>Yes</span>" } else { "<span class='tk-badge-info'>No</span>" } }
    function _ynWarn { param($b) if ($b) { "<span class='tk-badge-warn'>Yes</span>" } else { "<span class='tk-badge-ok'>No</span>" } }

    function _keyCell { param($P)
        if ([string]::IsNullOrEmpty($P.KeyMaterial)) { return "<span class='tk-badge-info'>None / enterprise</span>" }
        if ($IncludeKey) { return "<code class='tk-mono'>$(EscHtml $P.KeyMaterial)</code>" }
        return "<span class='tk-badge-warn'>masked (use -IncludeKey to show)</span>"
    }

    # Adapters table
    $adRows = [System.Text.StringBuilder]::new()
    if ($Adapters.Count -eq 0) {
        [void]$adRows.Append("<tr><td colspan='6' class='tk-badge-info' style='text-align:center;'>No Wi-Fi adapters detected.</td></tr>")
    } else {
        foreach ($a in $Adapters) {
            $statBadge = if ($a.Status -eq 'Up') { "<span class='tk-badge-ok'>Up</span>" } else { "<span class='tk-badge-warn'>$(EscHtml $a.Status)</span>" }
            [void]$adRows.Append(
                "<tr><td>$(EscHtml $a.Name)</td><td>$(EscHtml $a.InterfaceDescription)</td>" +
                "<td>$statBadge</td><td>$(EscHtml $a.LinkSpeed)</td>" +
                "<td class='tk-mono'>$(EscHtml $a.MacAddress)</td>" +
                "<td>$(EscHtml $a.DriverVersion) ($(EscHtml $a.DriverDate))</td></tr>"
            )
        }
    }

    # Profile table
    $pRows = [System.Text.StringBuilder]::new()
    if ($Profiles.Count -eq 0) {
        [void]$pRows.Append("<tr><td colspan='9' class='tk-badge-info' style='text-align:center;'>No saved Wi-Fi profiles found.</td></tr>")
    } else {
        foreach ($p in ($Profiles | Sort-Object Name)) {
            $authBadge   = _tierBadge (Get-AuthTier $p.Authentication)
            $cipherBadge = _tierBadge (Get-CipherTier $p.Encryption)
            $modeBadge   = if ($p.ConnectionMode -eq 'auto') { "<span class='tk-badge-warn'>auto</span>" } else { "<span class='tk-badge-info'>$(EscHtml $p.ConnectionMode)</span>" }
            [void]$pRows.Append(
                "<tr><td>$(EscHtml $p.Name)</td>" +
                "<td>$(EscHtml $p.Authentication) $authBadge</td>" +
                "<td>$(EscHtml $p.Encryption) $cipherBadge</td>" +
                "<td>$modeBadge</td>" +
                "<td>$(_ynWarn $p.AutoSwitch)</td>" +
                "<td>$(_ynWarn $p.NonBroadcast)</td>" +
                "<td>$(if ($null -eq $p.MacRandomization) { '<span class=''tk-badge-info''>n/a</span>' } else { _yn $p.MacRandomization })</td>" +
                "<td>$(_yn $p.UseOneX)</td>" +
                "<td>$(_keyCell $p)</td></tr>"
            )
        }
    }

    # Open / weak profiles
    $weak = @($Profiles | Where-Object { Test-IsWeakProfile $_ })
    $weakRows = [System.Text.StringBuilder]::new()
    if ($weak.Count -eq 0) {
        [void]$weakRows.Append("<tr><td colspan='4' class='tk-badge-ok' style='text-align:center;'>No open or weak-cipher profiles detected.</td></tr>")
    } else {
        foreach ($p in $weak | Sort-Object Name) {
            $reason = @()
            if (Test-IsOpenProfile $p)             { $reason += 'open authentication' }
            if ((Get-AuthTier $p.Authentication) -eq 'Insecure' -and "$($p.Authentication)".ToLower() -ne 'open') { $reason += "insecure auth $($p.Authentication)" }
            if ((Get-AuthTier $p.Authentication) -eq 'Weak')      { $reason += "weak auth $($p.Authentication)" }
            if ((Get-CipherTier $p.Encryption) -eq 'Insecure')    { $reason += "insecure cipher $($p.Encryption)" }
            if ((Get-CipherTier $p.Encryption) -eq 'Weak')        { $reason += "weak cipher $($p.Encryption)" }
            $reasonStr = ($reason -join ', ')
            [void]$weakRows.Append(
                "<tr><td>$(EscHtml $p.Name)</td><td>$(EscHtml $p.Authentication)</td>" +
                "<td>$(EscHtml $p.Encryption)</td><td>$(EscHtml $reasonStr)</td></tr>"
            )
        }
    }

    # Auto-connect profiles
    $auto = @($Profiles | Where-Object { $_.ConnectionMode -eq 'auto' })
    $autoRows = [System.Text.StringBuilder]::new()
    if ($auto.Count -eq 0) {
        [void]$autoRows.Append("<tr><td colspan='4' class='tk-badge-ok' style='text-align:center;'>No auto-connect profiles configured.</td></tr>")
    } else {
        foreach ($p in $auto | Sort-Object Name) {
            $rowClass = if (Test-IsOpenProfile $p) { 'tk-badge-err' }
                        elseif (Test-IsWeakProfile $p) { 'tk-badge-warn' }
                        else { 'tk-badge-info' }
            [void]$autoRows.Append(
                "<tr><td>$(EscHtml $p.Name)</td><td>$(EscHtml $p.Authentication)</td>" +
                "<td>$(EscHtml $p.Encryption)</td><td><span class='$rowClass'>$(if (Test-IsOpenProfile $p) { 'open auto-connect' } elseif (Test-IsWeakProfile $p) { 'weak auto-connect' } else { 'auto-connect' })</span></td></tr>"
            )
        }
    }

    $findingsList = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$findingsList.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>") }
    foreach ($w in $Verdict.Warns)  { [void]$findingsList.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$findingsList.Append("<li class='tk-badge-ok'>No insecure or surprising Wi-Fi profiles found.</li>")
    }

    $strongCount = @($Profiles | Where-Object { (Get-AuthTier $_.Authentication) -eq 'Strong' -and (Get-CipherTier $_.Encryption) -eq 'Strong' }).Count

    $htmlHead = Get-TKHtmlHead `
        -Title      'B.E.A.C.O.N. Wi-Fi Profile Audit' `
        -ScriptName 'B.E.A.C.O.N.' `
        -Subtitle   "${orgPrefix}Wi-Fi Profile Audit -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'        = $machine
            'Generated'      = $reportDate
            'Verdict'        = $Verdict.Verdict
            'Saved Profiles' = $Profiles.Count
            'Adapters'       = $Adapters.Count
        }) `
        -NavItems   @('Verdict', 'Adapters', 'Saved Profiles', 'Open / Weak', 'Auto-connect')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'B.E.A.C.O.N. v3.6'

    $keyNoteBadge = if ($IncludeKey) {
        "<span class='tk-badge-warn'>Key material rendered in cleartext (-IncludeKey)</span>"
    } else {
        "<span class='tk-badge-ok'>Key material masked (default)</span>"
    }

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Wi-Fi Posture</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Adapters.Count)</div><div class="tk-summary-lbl">Wi-Fi Adapter(s)</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Profiles.Count)</div><div class="tk-summary-lbl">Saved Profiles</div></div>
    <div class="tk-summary-card $(if ($weak.Count -gt 0) { 'err' } else { 'ok' })"><div class="tk-summary-num">$($weak.Count)</div><div class="tk-summary-lbl">Open / Weak</div></div>
    <div class="tk-summary-card $(if ($auto.Count -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$($auto.Count)</div><div class="tk-summary-lbl">Auto-connect</div></div>
    <div class="tk-summary-card ok"><div class="tk-summary-num">$strongCount</div><div class="tk-summary-lbl">Strong (WPA2/WPA3 + AES)</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Verdict &amp; Findings</span><span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span></div>
    <div class="tk-card">
      <ul class="tk-info-box" style="list-style:none;padding-left:0;">$($findingsList.ToString())</ul>
      <div class="tk-info-label">Privacy</div>
      <div class="tk-info-box">$keyNoteBadge</div>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Wi-Fi Adapters</span><span class="tk-section-num">$($Adapters.Count) detected</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Name</th><th>Description</th><th>Status</th><th>Link</th><th>MAC</th><th>Driver</th></tr></thead>
        <tbody>$($adRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Saved Profiles</span><span class="tk-section-num">$($Profiles.Count) total</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Profile</th><th>Authentication</th><th>Cipher</th><th>Connection</th><th>AutoSwitch</th><th>Hidden</th><th>MAC random</th><th>802.1X</th><th>Key</th></tr></thead>
        <tbody>$($pRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Open / Weak Profiles</span><span class="tk-section-num">$($weak.Count) flagged</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Profile</th><th>Authentication</th><th>Cipher</th><th>Reason</th></tr></thead>
        <tbody>$($weakRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Auto-connecting Profiles</span><span class="tk-section-num">$($auto.Count) auto-connect</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Profile</th><th>Authentication</th><th>Cipher</th><th>Risk Tier</th></tr></thead>
        <tbody>$($autoRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-BeaconBanner

Write-Section "WI-FI ADAPTERS"
$adapters = Get-WifiAdapters
if ($adapters.Count -eq 0) {
    Write-Info "No Wi-Fi adapters detected (desktop or wired-only configuration)."
} else {
    foreach ($a in $adapters) {
        $color = if ($a.Status -eq 'Up') { $C.Success } else { $C.Warning }
        Write-Host ("  {0,-12} {1,-8} {2}" -f $a.Name, $a.Status, $a.InterfaceDescription) -ForegroundColor $color
    }
}
Write-Host ""

Write-Section "SAVED PROFILES"
Write-Step "Exporting WLAN profiles to XML and parsing..."
$profiles = Get-WlanProfiles
Write-Host ("  Profiles parsed      : {0}" -f $profiles.Count) -ForegroundColor $C.Info
$openCount = @($profiles | Where-Object { Test-IsOpenProfile $_ }).Count
$weakCount = @($profiles | Where-Object { Test-IsWeakProfile $_ }).Count
Write-Host ("  Open profiles        : {0}" -f $openCount) -ForegroundColor $(if ($openCount -gt 0) { $C.Error } else { $C.Success })
Write-Host ("  Open / weak profiles : {0}" -f $weakCount) -ForegroundColor $(if ($weakCount -gt 0) { $C.Warning } else { $C.Success })
$autoCount = @($profiles | Where-Object { $_.ConnectionMode -eq 'auto' }).Count
Write-Host ("  Auto-connect         : {0}" -f $autoCount) -ForegroundColor $C.Info
Write-Host ""

$verdict = Get-BeaconVerdict -Adapters $adapters -Profiles $profiles

Write-Section "WI-FI VERDICT"
$verdictColor = switch ($verdict.Class) { 'ok' { $C.Success } 'warn' { $C.Warning } default { $C.Error } }
Write-Host "  $($verdict.Verdict)" -ForegroundColor $verdictColor
foreach ($i in $verdict.Issues) { Write-Host "    [!!] $i" -ForegroundColor $C.Error }
foreach ($w in $verdict.Warns)  { Write-Host "    [~ ] $w" -ForegroundColor $C.Warning }
if ($verdict.Issues.Count -eq 0 -and $verdict.Warns.Count -eq 0) {
    Write-Host "    [+ ] All checks passed." -ForegroundColor $C.Success
}
Write-Host ""

Write-Step "Generating HTML report..."
$html      = Build-BeaconReport -Adapters $adapters -Profiles $profiles -Verdict $verdict
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "BEACON_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
