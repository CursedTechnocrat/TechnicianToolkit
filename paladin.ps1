<#
.SYNOPSIS
    P.A.L.A.D.I.N. — Protection Auditor: Logs Antivirus, Defender, Intrusions & Notifications
    Antivirus / Microsoft Defender Health Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Audits the antivirus posture of a Windows endpoint. Reads Microsoft
    Defender state via `Get-MpComputerStatus` / `Get-MpPreference`, threat
    history via `Get-MpThreat` / `Get-MpThreatDetection`, registered AV
    products via the SecurityCenter2 WMI namespace, AV-related service
    state, and recent Defender Operational events. Produces a dark-themed
    HTML report with a red / yellow / green verdict and per-section detail
    tables: core state, real-time protections, cloud / sample submission,
    signatures, scan history, threats, exclusions, ASR rules, third-party
    AV products, service health, and recent events.

    Read-only audit -- no state-changing actions are performed.

.USAGE
    PS C:\> .\paladin.ps1                    # Interactive run
    PS C:\> .\paladin.ps1 -Unattended        # Silent: export HTML and exit
    PS C:\> .\paladin.ps1 -EventDays 14      # Look back 14 days for Defender events (default 7)
    PS C:\> .\paladin.ps1 -SignatureMaxAgeDays 3   # Tighter signature-age threshold (default 7)

.NOTES
    Version : 3.5

#>

param(
    [switch]$Unattended,
    [switch]$Transcript,
    [ValidateRange(1, 90)]
    [int]$EventDays = 7,
    [ValidateRange(1, 30)]
    [int]$SignatureMaxAgeDays = 7
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

function Show-PaladinBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  P.A.L.A.D.I.N. -- Protection Auditor: Logs Antivirus, Defender, Intrusions & Notifications" -ForegroundColor Magenta
    Write-Host "  AV / Microsoft Defender Health Audit  v3.5" -ForegroundColor Magenta
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

# ASR rule GUID -> friendly name. Microsoft publishes these IDs at
# learn.microsoft.com/en-us/defender-endpoint/attack-surface-reduction-rules-reference.
# Unknown GUIDs fall back to the raw GUID so a tenant deploying preview rules
# still sees them in the report.
$AsrRuleNames = @{
    '56a863a9-875e-4185-98a7-b882c64b5ce5' = 'Block abuse of exploited vulnerable signed drivers'
    '7674ba52-37eb-4a4f-a9a1-f0f9a1619a2c' = 'Block Adobe Reader from creating child processes'
    'd4f940ab-401b-4efc-aadc-ad5f3c50688a' = 'Block Office apps from creating child processes'
    '9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2' = 'Block credential stealing from LSASS'
    'be9ba2d9-53ea-4cdc-84e5-9b1eeee46550' = 'Block executable content from email/webmail'
    '01443614-cd74-433a-b99e-2ecdc07bfc25' = 'Block executable files unless meeting prevalence/age/trusted criteria'
    '5beb7efe-fd9a-4556-801d-275e5ffc04cc' = 'Block execution of potentially obfuscated scripts'
    'd3e037e1-3eb8-44c8-a917-57927947596d' = 'Block JavaScript/VBScript from launching downloaded content'
    '3b576869-a4ec-4529-8536-b80a7769e899' = 'Block Office apps from creating executable content'
    '75668c1f-73b5-4cf0-bb93-3ecf5cb7cc84' = 'Block Office apps from injecting code into other processes'
    '26190899-1602-49e8-8b27-eb1d0a1ce869' = 'Block Office communication apps from creating child processes'
    'e6db77e5-3df2-4cf1-b95a-636979351e5b' = 'Block persistence through WMI event subscription'
    'd1e49aac-8f56-4280-b9ba-993a6d77406c' = 'Block process creations from PsExec and WMI commands'
    'b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4' = 'Block untrusted/unsigned processes from USB'
    '92e97fa1-2edf-4476-bdd6-9dd0b4dddc7b' = 'Block Win32 API calls from Office macros'
    'c1db55ab-c21a-4637-bb3f-a12568109d35' = 'Use advanced protection against ransomware'
    'a8f5898e-1dc8-49a9-9878-85004b8a61e6' = 'Block Webshell creation for Servers'
    '33ddedf1-c6e0-47cb-833e-de6133960387' = 'Block rebooting machine in Safe Mode (preview)'
    'c0033c00-d16d-4114-a5a0-dc9b3a7d2ceb' = 'Block use of copied or impersonated system tools (preview)'
}

function Get-AsrActionLabel {
    param([int]$Action)
    switch ($Action) {
        0       { 'Not Configured' }
        1       { 'Block' }
        2       { 'Audit' }
        6       { 'Warn' }
        default { "Unknown ($Action)" }
    }
}

# Defender services we care about.
$DefenderServices = @(
    @{ Name = 'WinDefend';   Friendly = 'Microsoft Defender Antivirus Service'; Critical = $true  }
    @{ Name = 'WdNisSvc';    Friendly = 'Defender Network Inspection';          Critical = $false }
    @{ Name = 'Sense';       Friendly = 'Defender for Endpoint (EDR)';          Critical = $false }
    @{ Name = 'WdFilter';    Friendly = 'Defender mini-filter driver';          Critical = $true  }
    @{ Name = 'SecurityHealthService'; Friendly = 'Windows Security UI host';   Critical = $false }
    @{ Name = 'mpssvc';      Friendly = 'Windows Defender Firewall';            Critical = $false }
)

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-DefenderState {
    try {
        $s = Get-MpComputerStatus -ErrorAction Stop
    } catch {
        Write-TKError -ScriptName 'paladin.ps1' -Message "Get-MpComputerStatus failed: $($_.Exception.Message)" -Category 'Defender'
        return [PSCustomObject]@{ CollectorError = $_.Exception.Message }
    }

    return [PSCustomObject]@{
        CollectorError                 = $null
        ComputerState                  = $s.ComputerState
        AMServiceEnabled               = [bool]$s.AMServiceEnabled
        AMRunningMode                  = $s.AMRunningMode
        AMEngineVersion                = $s.AMEngineVersion
        AMServiceVersion               = $s.AMServiceVersion
        AMProductVersion               = $s.AMProductVersion
        AntivirusEnabled               = [bool]$s.AntivirusEnabled
        AntispywareEnabled             = [bool]$s.AntispywareEnabled
        RealTimeProtectionEnabled      = [bool]$s.RealTimeProtectionEnabled
        BehaviorMonitorEnabled         = [bool]$s.BehaviorMonitorEnabled
        IoavProtectionEnabled          = [bool]$s.IoavProtectionEnabled
        OnAccessProtectionEnabled      = [bool]$s.OnAccessProtectionEnabled
        NISEnabled                     = [bool]$s.NISEnabled
        TamperProtected                = [bool]$s.TamperProtected
        IsTamperProtected              = [bool]$s.IsTamperProtected
        AntivirusSignatureVersion      = $s.AntivirusSignatureVersion
        AntivirusSignatureLastUpdated  = $s.AntivirusSignatureLastUpdated
        AntivirusSignatureAge          = $s.AntivirusSignatureAge
        AntispywareSignatureVersion    = $s.AntispywareSignatureVersion
        AntispywareSignatureLastUpdated= $s.AntispywareSignatureLastUpdated
        AntispywareSignatureAge        = $s.AntispywareSignatureAge
        NISSignatureVersion            = $s.NISSignatureVersion
        NISSignatureLastUpdated        = $s.NISSignatureLastUpdated
        NISSignatureAge                = $s.NISSignatureAge
        QuickScanStartTime             = $s.QuickScanStartTime
        QuickScanEndTime               = $s.QuickScanEndTime
        QuickScanAge                   = $s.QuickScanAge
        FullScanStartTime              = $s.FullScanStartTime
        FullScanEndTime                = $s.FullScanEndTime
        FullScanAge                    = $s.FullScanAge
    }
}

function Get-DefenderPreference {
    try {
        $p = Get-MpPreference -ErrorAction Stop
    } catch {
        return [PSCustomObject]@{ CollectorError = $_.Exception.Message }
    }

    # ASR rule IDs and matching actions arrive as parallel arrays.
    $asr = [System.Collections.Generic.List[object]]::new()
    if ($p.AttackSurfaceReductionRules_Ids -and $p.AttackSurfaceReductionRules_Actions) {
        $ids     = @($p.AttackSurfaceReductionRules_Ids)
        $actions = @($p.AttackSurfaceReductionRules_Actions)
        for ($i = 0; $i -lt $ids.Count; $i++) {
            $id     = "$($ids[$i])".ToLower()
            $action = if ($i -lt $actions.Count) { [int]$actions[$i] } else { 0 }
            $asr.Add([PSCustomObject]@{
                Id      = $id
                Name    = if ($AsrRuleNames.ContainsKey($id)) { $AsrRuleNames[$id] } else { 'Unknown ASR rule' }
                Action  = $action
                Label   = Get-AsrActionLabel -Action $action
            })
        }
    }

    return [PSCustomObject]@{
        CollectorError              = $null
        ExclusionPath               = @($p.ExclusionPath      | Where-Object { $_ })
        ExclusionExtension          = @($p.ExclusionExtension | Where-Object { $_ })
        ExclusionProcess            = @($p.ExclusionProcess   | Where-Object { $_ })
        ExclusionIpAddress          = @($p.ExclusionIpAddress | Where-Object { $_ })
        AsrRules                    = @($asr)
        # Cloud / sample submission posture
        MAPSReporting               = $p.MAPSReporting
        SubmitSamplesConsent        = $p.SubmitSamplesConsent
        CloudBlockLevel             = $p.CloudBlockLevel
        CloudExtendedTimeout        = $p.CloudExtendedTimeout
        # Scan policies
        DisableArchiveScanning      = [bool]$p.DisableArchiveScanning
        DisableEmailScanning        = [bool]$p.DisableEmailScanning
        DisableScriptScanning       = [bool]$p.DisableScriptScanning
        DisableRemovableDriveScanning = [bool]$p.DisableRemovableDriveScanning
        PUAProtection               = $p.PUAProtection
        # Real-time controls (preference side; status side comes from Get-MpComputerStatus)
        DisableRealtimeMonitoring   = [bool]$p.DisableRealtimeMonitoring
        DisableBehaviorMonitoring   = [bool]$p.DisableBehaviorMonitoring
        DisableIOAVProtection       = [bool]$p.DisableIOAVProtection
    }
}

function Get-ThreatHistorySnapshot {
    $threats = @()
    try {
        $threats = @(Get-MpThreat -ErrorAction Stop)
    } catch {
        # Get-MpThreat throws when the cache is empty on some builds; treat as
        # empty rather than a hard collector failure.
        return [PSCustomObject]@{
            CollectorError = $null
            Threats        = @()
            UnresolvedHigh = 0
        }
    }

    $unresolvedHigh = 0
    $rows = foreach ($t in $threats) {
        $sevName = switch ([int]$t.SeverityID) {
            1 { 'Low' }    2 { 'Moderate' }   4 { 'High' }   5 { 'Severe' }
            default { "Sev$($t.SeverityID)" }
        }
        $isHigh = ($t.SeverityID -ge 4)
        # ActiveThreatExecutionStatus is non-zero when the threat is still acting on the system.
        # IsActive is true for any threat that hasn't been remediated.
        $active = [bool]$t.IsActive
        if ($isHigh -and $active) { $unresolvedHigh++ }

        [PSCustomObject]@{
            ThreatID       = $t.ThreatID
            ThreatName     = $t.ThreatName
            Severity       = $sevName
            SeverityID     = [int]$t.SeverityID
            CategoryID     = $t.CategoryID
            IsActive       = $active
            DetectionCount = $t.DetectionCount
            Resources      = (@($t.Resources) -join '; ')
        }
    }

    return [PSCustomObject]@{
        CollectorError = $null
        Threats        = @($rows)
        UnresolvedHigh = $unresolvedHigh
    }
}

function Get-RecentDetectionSnapshot {
    try {
        $detections = @(Get-MpThreatDetection -ErrorAction Stop)
    } catch {
        return [PSCustomObject]@{ CollectorError = $null; Detections = @() }
    }

    $rows = foreach ($d in $detections | Sort-Object InitialDetectionTime -Descending | Select-Object -First 50) {
        [PSCustomObject]@{
            DetectionID         = $d.DetectionID
            ThreatID            = $d.ThreatID
            InitialDetectionTime= $d.InitialDetectionTime
            LastThreatStatusChangeTime = $d.LastThreatStatusChangeTime
            ProcessName         = $d.ProcessName
            DomainUser          = $d.DomainUser
            Resources           = (@($d.Resources) -join '; ')
            ActionSuccess       = $d.ActionSuccess
            CleaningActionID    = $d.CleaningActionID
        }
    }

    return [PSCustomObject]@{ CollectorError = $null; Detections = @($rows) }
}

function Get-ThirdPartyAvProducts {
    # SecurityCenter2 namespace is absent on Server SKUs and Server Core; treat
    # absence as "no third-party AV registered" rather than a collector failure.
    try {
        $items = @(Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName 'AntiVirusProduct' -ErrorAction Stop)
    } catch {
        return [PSCustomObject]@{ Available = $false; Products = @() }
    }

    # productState is a packed bitfield; bits 12-13 are real-time protection,
    # bits 16-17 are signature freshness. The decoded values surface in the
    # report so a technician can spot a third-party AV silently in passive mode.
    $rows = foreach ($p in $items) {
        $state = [int]$p.productState
        # Documented decoding: see learn.microsoft.com / community references.
        $rt    = ($state -band 0x1000) -ne 0   # bit 12 set => real-time on
        $up    = ($state -band 0x10) -eq 0     # bit 4 clear => up to date
        [PSCustomObject]@{
            DisplayName  = $p.displayName
            ProductState = ('0x{0:X}' -f $state)
            RealTimeOn   = $rt
            UpToDate     = $up
            ExePath      = $p.pathToSignedReportingExe
        }
    }
    return [PSCustomObject]@{ Available = $true; Products = @($rows) }
}

function Get-DefenderServiceStatus {
    $rows = foreach ($svc in $DefenderServices) {
        $s = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
        if ($null -eq $s) {
            [PSCustomObject]@{
                Name      = $svc.Name
                Friendly  = $svc.Friendly
                Status    = 'NotInstalled'
                StartType = $null
                Critical  = $svc.Critical
            }
        } else {
            [PSCustomObject]@{
                Name      = $svc.Name
                Friendly  = $svc.Friendly
                Status    = $s.Status
                StartType = $s.StartType
                Critical  = $svc.Critical
            }
        }
    }
    return @($rows)
}

function Get-DefenderEvents {
    param([int]$Days)
    $since = (Get-Date).AddDays(-$Days)
    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName   = 'Microsoft-Windows-Windows Defender/Operational'
            StartTime = $since
        } -ErrorAction Stop | Select-Object -First 100
    } catch {
        return [PSCustomObject]@{ Available = $false; Events = @() }
    }

    $rows = foreach ($e in $events) {
        # Trim message to one line so the table stays readable.
        $msg = if ($e.Message) { ($e.Message -split "`r?`n")[0] } else { '' }
        [PSCustomObject]@{
            TimeCreated = $e.TimeCreated
            Id          = $e.Id
            Level       = $e.LevelDisplayName
            Provider    = $e.ProviderName
            Message     = $msg
        }
    }
    return [PSCustomObject]@{ Available = $true; Events = @($rows) }
}

# ─────────────────────────────────────────────────────────────────────────────
# VERDICT
# ─────────────────────────────────────────────────────────────────────────────

function Get-PaladinVerdict {
    param($State, $Pref, $Threats, $ThirdParty, $Services)

    $issues = [System.Collections.Generic.List[string]]::new()
    $warns  = [System.Collections.Generic.List[string]]::new()

    if ($State.CollectorError) {
        $issues.Add("Could not read Defender state: $($State.CollectorError). Defender may be uninstalled, blocked by GP, or this is Server Core without the AV role.")
        $verdict = 'AT RISK'
        $class   = 'err'
        return [PSCustomObject]@{ Verdict = $verdict; Class = $class; Issues = @($issues); Warns = @($warns) }
    }

    if (-not $State.AntivirusEnabled)                  { $issues.Add('Defender antivirus is DISABLED.') }
    if (-not $State.AMServiceEnabled)                  { $issues.Add('Defender AM service is not running.') }
    if (-not $State.RealTimeProtectionEnabled)         { $issues.Add('Real-time protection is OFF.') }
    if (-not $State.IsTamperProtected -and -not $State.TamperProtected) {
        $issues.Add('Tamper protection is OFF -- attackers can disable Defender from a local admin context.')
    }
    if ($null -ne $State.AntivirusSignatureAge -and $State.AntivirusSignatureAge -gt ($SignatureMaxAgeDays * 2)) {
        $issues.Add("Antivirus signatures are $($State.AntivirusSignatureAge) days old (> $($SignatureMaxAgeDays * 2) day red threshold).")
    }
    if ($Threats.UnresolvedHigh -gt 0) {
        $issues.Add("$($Threats.UnresolvedHigh) unresolved high/severe threat(s) recorded -- investigate before clearing.")
    }
    # No registered AV product at all (Defender absent and no third party) is a hard fail.
    if ($State.CollectorError -and -not $ThirdParty.Available) {
        $issues.Add('No registered AV product detected on the SecurityCenter2 surface.')
    }

    # WARNS
    if ($null -ne $State.AntivirusSignatureAge -and $State.AntivirusSignatureAge -gt $SignatureMaxAgeDays -and $State.AntivirusSignatureAge -le ($SignatureMaxAgeDays * 2)) {
        $warns.Add("Antivirus signatures are $($State.AntivirusSignatureAge) days old (> $SignatureMaxAgeDays day yellow threshold).")
    }
    if (-not $State.BehaviorMonitorEnabled)            { $warns.Add('Behavior monitoring is off.') }
    if (-not $State.IoavProtectionEnabled)             { $warns.Add('Downloaded-file scan (IOAV) is off.') }
    if (-not $State.OnAccessProtectionEnabled)         { $warns.Add('On-access scanning is off.') }
    if (-not $State.NISEnabled)                        { $warns.Add('Network Inspection (NIS) is off.') }
    if ($null -ne $State.QuickScanAge -and $State.QuickScanAge -gt 7) {
        $warns.Add("Last quick scan was $($State.QuickScanAge) day(s) ago.")
    }
    if ($null -ne $State.FullScanAge -and $State.FullScanAge -gt 30) {
        $warns.Add("Last full scan was $($State.FullScanAge) day(s) ago.")
    } elseif ($null -eq $State.FullScanAge) {
        $warns.Add('No full scan recorded on this machine.')
    }

    if ($Pref -and -not $Pref.CollectorError) {
        # MAPSReporting: 0=Disabled, 1=Basic, 2=Advanced
        if ([int]$Pref.MAPSReporting -eq 0) { $warns.Add('Cloud-delivered protection (MAPSReporting) is DISABLED.') }
        # SubmitSamplesConsent: 0=Always prompt, 1=Send safe, 2=Never send, 3=Send all
        if ([int]$Pref.SubmitSamplesConsent -eq 2) { $warns.Add('Sample submission is set to NEVER -- cloud blocks will miss novel threats.') }
        # CloudBlockLevel: 0=Default, 2=High, 4=High+, 6=Zero-tolerance
        if ([int]$Pref.CloudBlockLevel -eq 0) { $warns.Add('Cloud block level is at default (lowest tier).') }

        $auditOnly = @($Pref.AsrRules | Where-Object { $_.Action -eq 2 })
        if ($auditOnly.Count -gt 0) {
            $warns.Add("$($auditOnly.Count) ASR rule(s) are in Audit-only mode -- detections logged but not blocked.")
        }
        $totalExclusions = $Pref.ExclusionPath.Count + $Pref.ExclusionExtension.Count + $Pref.ExclusionProcess.Count + $Pref.ExclusionIpAddress.Count
        if ($totalExclusions -gt 20) {
            $warns.Add("$totalExclusions exclusions are configured -- a large exclusion footprint reduces effective protection coverage.")
        }
        if ($Pref.DisableArchiveScanning)        { $warns.Add('Archive (.zip / .iso / .7z) scanning is disabled.') }
        if ($Pref.DisableEmailScanning)          { $warns.Add('Email scanning is disabled.') }
        if ($Pref.DisableScriptScanning)         { $warns.Add('Script scanning is disabled.') }
        if ($Pref.DisableRemovableDriveScanning) { $warns.Add('Removable-drive scanning is disabled.') }
        # PUAProtection: 0=Disabled, 1=Block, 2=Audit
        if ([int]$Pref.PUAProtection -eq 0)      { $warns.Add('Potentially Unwanted Application (PUA) protection is DISABLED.') }
    }

    # Critical service health
    foreach ($svc in $Services) {
        if ($svc.Critical -and $svc.Status -ne 'Running') {
            $issues.Add("Critical service $($svc.Name) ($($svc.Friendly)) is $($svc.Status).")
        }
    }

    # Third-party AV in real-time mode while Defender is also active in real-time mode is a posture conflict, not an outright failure.
    if ($State.RealTimeProtectionEnabled -and $ThirdParty.Available) {
        $rtThird = @($ThirdParty.Products | Where-Object { $_.RealTimeOn })
        if ($rtThird.Count -gt 0) {
            $warns.Add("$($rtThird.Count) third-party AV product(s) running in real-time alongside Defender -- only one should own real-time protection.")
        }
    }

    $verdict = if ($issues.Count -gt 0) { 'AT RISK' }
               elseif ($warns.Count -gt 0) { 'ATTENTION NEEDED' }
               else { 'PROTECTED' }
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

function Build-PaladinReport {
    param($State, $Pref, $Threats, $Detections, $ThirdParty, $Services, $Events, $Verdict)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    function _yn { param($b) if ($b) { "<span class='tk-badge-ok'>Yes</span>" } else { "<span class='tk-badge-err'>No</span>" } }
    function _ynWarn { param($b) if ($b) { "<span class='tk-badge-ok'>Yes</span>" } else { "<span class='tk-badge-warn'>No</span>" } }

    # Verdict findings list
    $findingsList = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$findingsList.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>`n") }
    foreach ($w in $Verdict.Warns)  { [void]$findingsList.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>`n") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$findingsList.Append("<li class='tk-badge-ok'>Defender posture is clean -- real-time on, signatures fresh, no unresolved threats.</li>")
    }

    # When Get-MpPreference fails (Server Core, blocked GP, missing Defender)
    # the $Pref object only carries CollectorError -- the numeric properties
    # don't exist. PowerShell silently returns $null for missing properties
    # and `[int]$null` lands on `0`, which would make the Cloud / Sample
    # Submission table render every setting as "Disabled" (the 0-case label).
    # That's misleading. Render the per-setting labels only when the
    # collector succeeded; otherwise render an "unavailable" marker.
    $prefAvailable = ($Pref -and -not $Pref.CollectorError)

    if ($prefAvailable) {
        $mapsLabel = switch ([int]$Pref.MAPSReporting) {
            0 { 'Disabled' } 1 { 'Basic' } 2 { 'Advanced (MAPS)' } default { "Code $($Pref.MAPSReporting)" }
        }
        $sampleLabel = switch ([int]$Pref.SubmitSamplesConsent) {
            0 { 'Always prompt' } 1 { 'Send safe samples' } 2 { 'Never send' } 3 { 'Send all samples' } default { "Code $($Pref.SubmitSamplesConsent)" }
        }
        $cloudLevelLabel = switch ([int]$Pref.CloudBlockLevel) {
            0 { 'Default' } 2 { 'High' } 4 { 'High+' } 6 { 'Zero-tolerance' } default { "Code $($Pref.CloudBlockLevel)" }
        }
        $puaLabel = switch ([int]$Pref.PUAProtection) {
            0 { 'Disabled' } 1 { 'Block' } 2 { 'Audit' } default { "Code $($Pref.PUAProtection)" }
        }
    } else {
        $mapsLabel       = 'unavailable'
        $sampleLabel     = 'unavailable'
        $cloudLevelLabel = 'unavailable'
        $puaLabel        = 'unavailable'
    }

    # The Cloud & Sample Submission card body is built conditionally: full
    # table when preferences were readable, single notice card when not.
    # Building it here (rather than embedding the if-else inside the heredoc)
    # avoids interpolating boolean operators inside `$(...)` -- those don't
    # evaluate cleanly across PS5.1's heredoc parser.
    $cloudSampleBody = if ($prefAvailable) { @"
<table class="tk-table"><tbody>
        <tr><th>MAPS Reporting</th><td>$(EscHtml $mapsLabel)</td></tr>
        <tr><th>Sample Submission Consent</th><td>$(EscHtml $sampleLabel)</td></tr>
        <tr><th>Cloud Block Level</th><td>$(EscHtml $cloudLevelLabel)</td></tr>
        <tr><th>Cloud Extended Timeout (s)</th><td>$(EscHtml $Pref.CloudExtendedTimeout)</td></tr>
        <tr><th>PUA Protection</th><td>$(EscHtml $puaLabel)</td></tr>
        <tr><th>Archive scanning</th><td>$(_ynWarn (-not $Pref.DisableArchiveScanning))</td></tr>
        <tr><th>Email scanning</th><td>$(_ynWarn (-not $Pref.DisableEmailScanning))</td></tr>
        <tr><th>Script scanning</th><td>$(_ynWarn (-not $Pref.DisableScriptScanning))</td></tr>
        <tr><th>Removable-drive scanning</th><td>$(_ynWarn (-not $Pref.DisableRemovableDriveScanning))</td></tr>
      </tbody></table>
"@
    } else { @"
<div class="tk-info-box tk-badge-warn">Defender preferences could not be read on this machine -- cloud / sample submission posture is unknown.$(if ($Pref.CollectorError) { " Collector error: $(EscHtml $Pref.CollectorError)" } else { '' })</div>
"@
    }

    # Signature age cell helper
    function _ageBadge { param($age, $maxOk, $maxWarn)
        if ($null -eq $age) { return "<span class='tk-badge-warn'>Unknown</span>" }
        if ($age -le $maxOk)   { return "<span class='tk-badge-ok'>$age day(s)</span>" }
        if ($age -le $maxWarn) { return "<span class='tk-badge-warn'>$age day(s)</span>" }
        return "<span class='tk-badge-err'>$age day(s)</span>"
    }

    # Threats table
    $threatRows = [System.Text.StringBuilder]::new()
    if ($Threats.Threats.Count -eq 0) {
        [void]$threatRows.Append("<tr><td colspan='6' class='tk-badge-ok' style='text-align:center;'>No threats recorded.</td></tr>")
    } else {
        foreach ($t in $Threats.Threats | Sort-Object SeverityID -Descending) {
            $sevBadge = switch ($t.SeverityID) {
                { $_ -ge 5 } { "<span class='tk-badge-err'>$(EscHtml $t.Severity)</span>" }
                { $_ -ge 4 } { "<span class='tk-badge-err'>$(EscHtml $t.Severity)</span>" }
                { $_ -ge 2 } { "<span class='tk-badge-warn'>$(EscHtml $t.Severity)</span>" }
                default      { "<span class='tk-badge-info'>$(EscHtml $t.Severity)</span>" }
            }
            $activeBadge = if ($t.IsActive) { "<span class='tk-badge-err'>Active</span>" } else { "<span class='tk-badge-ok'>Resolved</span>" }
            [void]$threatRows.Append(
                "<tr><td>$(EscHtml $t.ThreatName)</td><td>$sevBadge</td><td>$activeBadge</td>" +
                "<td class='tk-mono'>$(EscHtml $t.ThreatID)</td><td>$(EscHtml $t.DetectionCount)</td>" +
                "<td class='tk-mono'>$(EscHtml $t.Resources)</td></tr>`n"
            )
        }
    }

    # Detections table
    $detRows = [System.Text.StringBuilder]::new()
    if ($Detections.Detections.Count -eq 0) {
        [void]$detRows.Append("<tr><td colspan='5' class='tk-badge-ok' style='text-align:center;'>No recent detections.</td></tr>")
    } else {
        foreach ($d in $Detections.Detections) {
            $okBadge = if ($d.ActionSuccess) { "<span class='tk-badge-ok'>Yes</span>" } else { "<span class='tk-badge-warn'>No</span>" }
            [void]$detRows.Append(
                "<tr><td>$(EscHtml ($d.InitialDetectionTime))</td><td>$(EscHtml $d.ProcessName)</td>" +
                "<td>$(EscHtml $d.DomainUser)</td><td class='tk-mono'>$(EscHtml $d.Resources)</td><td>$okBadge</td></tr>`n"
            )
        }
    }

    # Exclusions tables (one per category, only render the section if anything is configured)
    function _excTable {
        param([string]$Heading, [string[]]$Items)
        $sb = [System.Text.StringBuilder]::new()
        [void]$sb.Append("<div class='tk-info-label'>$Heading ($($Items.Count))</div>")
        if ($Items.Count -eq 0) {
            [void]$sb.Append("<div class='tk-info-box tk-badge-ok'>None configured.</div>")
        } else {
            [void]$sb.Append('<table class="tk-table"><thead><tr><th>Value</th></tr></thead><tbody>')
            foreach ($x in $Items) { [void]$sb.Append("<tr><td class='tk-mono'>$(EscHtml $x)</td></tr>") }
            [void]$sb.Append('</tbody></table>')
        }
        return $sb.ToString()
    }

    $exclusionsBlock = ''
    if ($Pref -and -not $Pref.CollectorError) {
        $exclusionsBlock =
            (_excTable -Heading 'Path exclusions'      -Items $Pref.ExclusionPath) +
            (_excTable -Heading 'Extension exclusions' -Items $Pref.ExclusionExtension) +
            (_excTable -Heading 'Process exclusions'   -Items $Pref.ExclusionProcess) +
            (_excTable -Heading 'IP exclusions'        -Items $Pref.ExclusionIpAddress)
    } else {
        $exclusionsBlock = "<div class='tk-info-box tk-badge-warn'>Could not read Defender preferences.</div>"
    }

    # ASR rules
    $asrRows = [System.Text.StringBuilder]::new()
    if (-not $Pref -or $Pref.CollectorError -or $Pref.AsrRules.Count -eq 0) {
        [void]$asrRows.Append("<tr><td colspan='3' class='tk-badge-warn' style='text-align:center;'>No ASR rules configured (or preferences unavailable).</td></tr>")
    } else {
        foreach ($r in $Pref.AsrRules | Sort-Object Name) {
            $badge = switch ($r.Action) {
                1 { "<span class='tk-badge-ok'>Block</span>" }
                2 { "<span class='tk-badge-warn'>Audit</span>" }
                6 { "<span class='tk-badge-warn'>Warn</span>" }
                0 { "<span class='tk-badge-info'>Not Configured</span>" }
                default { "<span class='tk-badge-info'>$(EscHtml $r.Label)</span>" }
            }
            [void]$asrRows.Append("<tr><td>$(EscHtml $r.Name)</td><td>$badge</td><td class='tk-mono'>$(EscHtml $r.Id)</td></tr>")
        }
    }

    # Third-party AV
    $tpRows = [System.Text.StringBuilder]::new()
    if (-not $ThirdParty.Available -or $ThirdParty.Products.Count -eq 0) {
        [void]$tpRows.Append("<tr><td colspan='4' class='tk-badge-info' style='text-align:center;'>No third-party AV products registered.</td></tr>")
    } else {
        foreach ($p in $ThirdParty.Products) {
            $rt = if ($p.RealTimeOn) { "<span class='tk-badge-ok'>On</span>" } else { "<span class='tk-badge-warn'>Off</span>" }
            $up = if ($p.UpToDate)   { "<span class='tk-badge-ok'>Yes</span>" } else { "<span class='tk-badge-warn'>No</span>" }
            [void]$tpRows.Append("<tr><td>$(EscHtml $p.DisplayName)</td><td>$rt</td><td>$up</td><td class='tk-mono'>$(EscHtml $p.ExePath)</td></tr>")
        }
    }

    # Services
    $svcRows = [System.Text.StringBuilder]::new()
    foreach ($s in $Services) {
        $statusBadge = switch ($s.Status) {
            'Running'      { "<span class='tk-badge-ok'>Running</span>" }
            'Stopped'      { if ($s.Critical) { "<span class='tk-badge-err'>Stopped</span>" } else { "<span class='tk-badge-warn'>Stopped</span>" } }
            'NotInstalled' { "<span class='tk-badge-info'>Not Installed</span>" }
            default        { "<span class='tk-badge-warn'>$(EscHtml $s.Status)</span>" }
        }
        $crit = if ($s.Critical) { "<span class='tk-badge-err'>Critical</span>" } else { "<span class='tk-badge-info'>Optional</span>" }
        [void]$svcRows.Append("<tr><td class='tk-mono'>$(EscHtml $s.Name)</td><td>$(EscHtml $s.Friendly)</td><td>$statusBadge</td><td>$(EscHtml $s.StartType)</td><td>$crit</td></tr>")
    }

    # Events
    $evRows = [System.Text.StringBuilder]::new()
    if (-not $Events.Available) {
        [void]$evRows.Append("<tr><td colspan='4' class='tk-badge-warn' style='text-align:center;'>Defender Operational log unavailable (no events in window or log not present).</td></tr>")
    } elseif ($Events.Events.Count -eq 0) {
        [void]$evRows.Append("<tr><td colspan='4' class='tk-badge-ok' style='text-align:center;'>No Defender events in the last $EventDays day(s).</td></tr>")
    } else {
        foreach ($e in $Events.Events) {
            $lvlBadge = switch ($e.Level) {
                'Critical'      { "<span class='tk-badge-err'>$(EscHtml $e.Level)</span>" }
                'Error'         { "<span class='tk-badge-err'>$(EscHtml $e.Level)</span>" }
                'Warning'       { "<span class='tk-badge-warn'>$(EscHtml $e.Level)</span>" }
                'Information'   { "<span class='tk-badge-info'>$(EscHtml $e.Level)</span>" }
                default         { "<span class='tk-badge-info'>$(EscHtml $e.Level)</span>" }
            }
            [void]$evRows.Append("<tr><td>$(EscHtml ($e.TimeCreated))</td><td>$(EscHtml $e.Id)</td><td>$lvlBadge</td><td>$(EscHtml $e.Message)</td></tr>")
        }
    }

    # Summary cards
    $rtClass = if ($State.RealTimeProtectionEnabled) { 'ok' } else { 'err' }
    $tpClass = if ($State.IsTamperProtected -or $State.TamperProtected) { 'ok' } else { 'err' }
    $sigClass = if ($State.AntivirusSignatureAge -le $SignatureMaxAgeDays) { 'ok' }
                elseif ($State.AntivirusSignatureAge -le ($SignatureMaxAgeDays * 2)) { 'warn' }
                else { 'err' }
    $threatClass = if ($Threats.UnresolvedHigh -gt 0) { 'err' } elseif ($Threats.Threats.Count -gt 0) { 'warn' } else { 'ok' }

    $htmlHead = Get-TKHtmlHead `
        -Title      'P.A.L.A.D.I.N. AV / Defender Health Report' `
        -ScriptName 'P.A.L.A.D.I.N.' `
        -Subtitle   "${orgPrefix}AV / Microsoft Defender Health Audit -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'       = $machine
            'Generated'     = $reportDate
            'Verdict'       = $Verdict.Verdict
            'Engine'        = $State.AMEngineVersion
            'AV Sig Version'= $State.AntivirusSignatureVersion
        }) `
        -NavItems   @('Verdict', 'Defender Core', 'Cloud & Sample', 'Signatures', 'Scans', 'Threats', 'Detections', 'Exclusions', 'ASR Rules', 'Third-Party AV', 'Services', 'Events')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'P.A.L.A.D.I.N. v3.5'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Posture</div></div>
    <div class="tk-summary-card $rtClass"><div class="tk-summary-num">$(if ($State.RealTimeProtectionEnabled) { 'On' } else { 'Off' })</div><div class="tk-summary-lbl">Real-time Protection</div></div>
    <div class="tk-summary-card $tpClass"><div class="tk-summary-num">$(if ($State.IsTamperProtected -or $State.TamperProtected) { 'On' } else { 'Off' })</div><div class="tk-summary-lbl">Tamper Protection</div></div>
    <div class="tk-summary-card $sigClass"><div class="tk-summary-num">$(if ($null -eq $State.AntivirusSignatureAge) { '?' } else { "$($State.AntivirusSignatureAge)d" })</div><div class="tk-summary-lbl">AV Signature Age</div></div>
    <div class="tk-summary-card $threatClass"><div class="tk-summary-num">$($Threats.Threats.Count)</div><div class="tk-summary-lbl">Threats in History</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$(EscHtml $State.AMRunningMode)</div><div class="tk-summary-lbl">AM Running Mode</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Verdict &amp; Findings</span><span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span></div>
    <div class="tk-card"><ul class="tk-info-box" style="list-style:none;padding-left:0;">$($findingsList.ToString())</ul></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Defender Core State</span></div>
    <div class="tk-card">
      <table class="tk-table"><tbody>
        <tr><th>Antivirus Enabled</th><td>$(_yn $State.AntivirusEnabled)</td></tr>
        <tr><th>Antispyware Enabled</th><td>$(_yn $State.AntispywareEnabled)</td></tr>
        <tr><th>AM Service Enabled</th><td>$(_yn $State.AMServiceEnabled)</td></tr>
        <tr><th>AM Running Mode</th><td>$(EscHtml $State.AMRunningMode)</td></tr>
        <tr><th>AM Engine Version</th><td class='tk-mono'>$(EscHtml $State.AMEngineVersion)</td></tr>
        <tr><th>AM Service Version</th><td class='tk-mono'>$(EscHtml $State.AMServiceVersion)</td></tr>
        <tr><th>AM Product Version</th><td class='tk-mono'>$(EscHtml $State.AMProductVersion)</td></tr>
        <tr><th>Real-time Protection</th><td>$(_yn $State.RealTimeProtectionEnabled)</td></tr>
        <tr><th>Behavior Monitor</th><td>$(_ynWarn $State.BehaviorMonitorEnabled)</td></tr>
        <tr><th>IOAV (downloaded files)</th><td>$(_ynWarn $State.IoavProtectionEnabled)</td></tr>
        <tr><th>On-access Protection</th><td>$(_ynWarn $State.OnAccessProtectionEnabled)</td></tr>
        <tr><th>Network Inspection (NIS)</th><td>$(_ynWarn $State.NISEnabled)</td></tr>
        <tr><th>Tamper Protected</th><td>$(_yn ($State.IsTamperProtected -or $State.TamperProtected))</td></tr>
      </tbody></table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Cloud &amp; Sample Submission</span></div>
    <div class="tk-card">
      $cloudSampleBody
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Signatures</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Signature Set</th><th>Version</th><th>Last Updated</th><th>Age</th></tr></thead>
        <tbody>
          <tr><td>Antivirus</td><td class='tk-mono'>$(EscHtml $State.AntivirusSignatureVersion)</td><td>$(EscHtml $State.AntivirusSignatureLastUpdated)</td><td>$(_ageBadge $State.AntivirusSignatureAge $SignatureMaxAgeDays ($SignatureMaxAgeDays * 2))</td></tr>
          <tr><td>Antispyware</td><td class='tk-mono'>$(EscHtml $State.AntispywareSignatureVersion)</td><td>$(EscHtml $State.AntispywareSignatureLastUpdated)</td><td>$(_ageBadge $State.AntispywareSignatureAge $SignatureMaxAgeDays ($SignatureMaxAgeDays * 2))</td></tr>
          <tr><td>Network Inspection (NIS)</td><td class='tk-mono'>$(EscHtml $State.NISSignatureVersion)</td><td>$(EscHtml $State.NISSignatureLastUpdated)</td><td>$(_ageBadge $State.NISSignatureAge $SignatureMaxAgeDays ($SignatureMaxAgeDays * 2))</td></tr>
        </tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Scan History</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Scan</th><th>Last Started</th><th>Last Ended</th><th>Age</th></tr></thead>
        <tbody>
          <tr><td>Quick Scan</td><td>$(EscHtml $State.QuickScanStartTime)</td><td>$(EscHtml $State.QuickScanEndTime)</td><td>$(_ageBadge $State.QuickScanAge 7 30)</td></tr>
          <tr><td>Full Scan</td><td>$(EscHtml $State.FullScanStartTime)</td><td>$(EscHtml $State.FullScanEndTime)</td><td>$(_ageBadge $State.FullScanAge 30 90)</td></tr>
        </tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Threat History</span><span class="tk-section-num">$($Threats.Threats.Count) entries</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Threat</th><th>Severity</th><th>State</th><th>Threat ID</th><th>Detections</th><th>Resources</th></tr></thead>
        <tbody>$($threatRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Recent Detections</span><span class="tk-section-num">$($Detections.Detections.Count) shown</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Initial Detection</th><th>Process</th><th>User</th><th>Resources</th><th>Cleanup OK?</th></tr></thead>
        <tbody>$($detRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Exclusions</span></div>
    <div class="tk-card">$exclusionsBlock</div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Attack Surface Reduction Rules</span><span class="tk-section-num">$(if ($Pref.AsrRules) { $Pref.AsrRules.Count } else { 0 }) configured</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Rule</th><th>Mode</th><th>Rule ID</th></tr></thead>
        <tbody>$($asrRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Third-Party AV Products</span><span class="tk-section-num">SecurityCenter2</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Product</th><th>Real-time</th><th>Up to Date</th><th>Reporting EXE</th></tr></thead>
        <tbody>$($tpRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Service Health</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Service</th><th>Description</th><th>Status</th><th>Start Type</th><th>Tier</th></tr></thead>
        <tbody>$($svcRows.ToString())</tbody>
      </table>
    </div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Recent Defender Events</span><span class="tk-section-num">last $EventDays day(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Time</th><th>Event ID</th><th>Level</th><th>Message</th></tr></thead>
        <tbody>$($evRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-PaladinBanner

Write-Section "DEFENDER STATE"
$state = Get-DefenderState
if ($state.CollectorError) {
    Write-Fail "Get-MpComputerStatus failed: $($state.CollectorError)"
    Write-Warn "Defender may be uninstalled, blocked by GP, or this is Server Core. The report will still render with a partial picture."
} else {
    Write-Host ("  Antivirus enabled    : {0}" -f $state.AntivirusEnabled) -ForegroundColor $(if ($state.AntivirusEnabled) { $C.Success } else { $C.Error })
    Write-Host ("  Real-time protection : {0}" -f $state.RealTimeProtectionEnabled) -ForegroundColor $(if ($state.RealTimeProtectionEnabled) { $C.Success } else { $C.Error })
    Write-Host ("  Tamper protection    : {0}" -f ($state.IsTamperProtected -or $state.TamperProtected)) -ForegroundColor $(if ($state.IsTamperProtected -or $state.TamperProtected) { $C.Success } else { $C.Error })
    Write-Host ("  AV sig age (days)    : {0}" -f $state.AntivirusSignatureAge) -ForegroundColor $(
        if ($null -eq $state.AntivirusSignatureAge) { $C.Warning }
        elseif ($state.AntivirusSignatureAge -le $SignatureMaxAgeDays) { $C.Success }
        elseif ($state.AntivirusSignatureAge -le ($SignatureMaxAgeDays * 2)) { $C.Warning }
        else { $C.Error }
    )
    Write-Host ("  Engine               : {0}" -f $state.AMEngineVersion) -ForegroundColor $C.Info
}
Write-Host ""

Write-Section "PREFERENCES (cloud, exclusions, ASR)"
$pref = Get-DefenderPreference
if ($pref.CollectorError) {
    Write-Warn "Get-MpPreference failed: $($pref.CollectorError)"
} else {
    Write-Host ("  Path exclusions      : {0}" -f $pref.ExclusionPath.Count) -ForegroundColor $C.Info
    Write-Host ("  Process exclusions   : {0}" -f $pref.ExclusionProcess.Count) -ForegroundColor $C.Info
    Write-Host ("  ASR rules configured : {0}" -f $pref.AsrRules.Count) -ForegroundColor $C.Info
}
Write-Host ""

Write-Section "THREATS"
$threats = Get-ThreatHistorySnapshot
Write-Host ("  Threats in history   : {0}" -f $threats.Threats.Count) -ForegroundColor $C.Info
Write-Host ("  Unresolved high/sev  : {0}" -f $threats.UnresolvedHigh) -ForegroundColor $(if ($threats.UnresolvedHigh -gt 0) { $C.Error } else { $C.Success })
foreach ($t in $threats.Threats | Where-Object { $_.IsActive -and $_.SeverityID -ge 4 }) {
    # Telemetry path: every active high/severe threat warrants a Teams ping if a webhook is configured.
    Write-TKError -ScriptName 'paladin.ps1' -Message "Active high/severe threat: $($t.ThreatName) (Severity $($t.Severity), DetectionCount $($t.DetectionCount))" -Category 'Defender'
}
$detections = Get-RecentDetectionSnapshot
Write-Host ("  Recent detections    : {0}" -f $detections.Detections.Count) -ForegroundColor $C.Info
Write-Host ""

Write-Section "THIRD-PARTY AV (SecurityCenter2)"
$thirdParty = Get-ThirdPartyAvProducts
if (-not $thirdParty.Available) {
    Write-Info "SecurityCenter2 namespace unavailable -- no third-party data collected (normal on Server SKUs)."
} else {
    Write-Host ("  Registered products  : {0}" -f $thirdParty.Products.Count) -ForegroundColor $C.Info
    foreach ($p in $thirdParty.Products) {
        Write-Host ("    - {0}  RT: {1}  UpToDate: {2}" -f $p.DisplayName, $p.RealTimeOn, $p.UpToDate) -ForegroundColor $(if ($p.RealTimeOn) { $C.Success } else { $C.Warning })
    }
}
Write-Host ""

Write-Section "SERVICES"
$services = Get-DefenderServiceStatus
foreach ($s in $services) {
    $color = if ($s.Status -eq 'Running')         { $C.Success }
             elseif ($s.Status -eq 'NotInstalled') { $C.Info }
             elseif ($s.Critical)                 { $C.Error }
             else                                  { $C.Warning }
    Write-Host ("  {0,-22} {1,-12} {2}" -f $s.Name, $s.Status, $s.Friendly) -ForegroundColor $color
}
Write-Host ""

Write-Section "RECENT EVENTS (last $EventDays day(s))"
$events = Get-DefenderEvents -Days $EventDays
if (-not $events.Available) {
    Write-Info "Defender Operational log returned no events (or is unavailable)."
} else {
    Write-Host ("  Events captured      : {0}" -f $events.Events.Count) -ForegroundColor $C.Info
}
Write-Host ""

$verdict = Get-PaladinVerdict -State $state -Pref $pref -Threats $threats -ThirdParty $thirdParty -Services $services

Write-Section "AV / DEFENDER VERDICT"
$verdictColor = switch ($verdict.Class) { 'ok' { $C.Success } 'warn' { $C.Warning } default { $C.Error } }
Write-Host "  $($verdict.Verdict)" -ForegroundColor $verdictColor
foreach ($i in $verdict.Issues) { Write-Host "    [!!] $i" -ForegroundColor $C.Error }
foreach ($w in $verdict.Warns)  { Write-Host "    [~ ] $w" -ForegroundColor $C.Warning }
if ($verdict.Issues.Count -eq 0 -and $verdict.Warns.Count -eq 0) {
    Write-Host "    [+ ] All checks passed." -ForegroundColor $C.Success
}
Write-Host ""

Write-Step "Generating HTML report..."
$html      = Build-PaladinReport -State $state -Pref $pref -Threats $threats -Detections $detections -ThirdParty $thirdParty -Services $services -Events $events -Verdict $verdict
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "PALADIN_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Show-TKReportResult -Path $outPath -Unattended:$Unattended
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
