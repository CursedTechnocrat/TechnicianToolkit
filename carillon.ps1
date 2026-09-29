# carillon.ps1 - C.A.R.I.L.L.O.N. — Catalogs Attendants, Routing, Inbound Lines, Listeners, Overflow & Numbers
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
    C.A.R.I.L.L.O.N. — Catalogs Attendants, Routing, Inbound Lines, Listeners, Overflow & Numbers
    Teams Phone Call Queue & Auto Attendant Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Inventories every Teams Phone call queue and auto attendant in the tenant
    and shows who is actually connected to them -- by display name, never by
    object ID:

      - Call queues: phone numbers, routing method, presence-based routing,
        conference mode, overflow / timeout / no-agent handling and where
        those calls go, and every agent with their opt-in state, whether
        they are voice-enabled, and how they got there (added directly,
        through a group, or through a Teams channel).
      - Auto attendants: phone numbers, language and time zone, operator,
        the business-hours menu (each key and where it transfers), and the
        after-hours and holiday call flows with their schedules.
      - Resource accounts: every account, its number, and the queue or
        attendant it fronts.
      - Routing: which attendants and queues hand calls to each queue or
        attendant, so an unreachable one stands out.

    Flags queues nobody can answer (no agents, or every agent opted out),
    agents who are not voice-enabled or are disabled, routing targets that
    point at deleted users or groups, queues and attendants with no number
    and nothing routing to them, and resource accounts assigned to nothing.

    Read-only -- nothing in the tenant is changed. Needs the MicrosoftTeams
    module for the queues and attendants, and Microsoft.Graph.Authentication
    (Directory.Read.All) to put names on groups and to show which group
    brought each agent in. Both are offered for install if missing.
    -SkipGraph avoids the second sign-in; users are still named through the
    Teams module, but group names are not resolved.

    Writes an HTML report and a CSV of every queue / agent pairing to the
    log directory.

.USAGE
    PS C:\> .\carillon.ps1                     # Interactive: sign in, audit, HTML + CSV
    PS C:\> .\carillon.ps1 -Name 'Sales*'      # Only queues / attendants whose name matches
    PS C:\> .\carillon.ps1 -SkipGraph          # Teams sign-in only; group names left unresolved
    PS C:\> .\carillon.ps1 -Unattended         # Silent: sign in, audit, export

.NOTES
    Version : 5.1

#>

param(
    [switch]$Unattended,
    [switch]$Transcript,
    [string]$Name,
    [switch]$SkipGraph
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
# REFERENCE TABLES
# ─────────────────────────────────────────────────────────────────────────────

$GraphScopes = @('Directory.Read.All')

# The application IDs a resource account carries say what it can front. They
# are the same in every tenant.
$ResourceAccountApps = @{
    '11cd3e2e-fccb-42ad-ad00-878b93575e07' = 'Call queue'
    'ce933385-9390-45d1-9512-c8d228074e07' = 'Auto attendant'
}

# Queue actions that end the call rather than send it somewhere.
$DisconnectActions = @('Disconnect', 'DisconnectWithBusy')

# Readable text for the routing enums the Teams module returns.
$RoutingMethodNames = @{
    'Attendant'   = 'Attendant (ring all)'
    'Serial'      = 'Serial'
    'RoundRobin'  = 'Round robin'
    'LongestIdle' = 'Longest idle'
}
$MenuActionNames = @{
    'DisconnectCall'         = 'Disconnect'
    'TransferCallToOperator' = 'Transfer to operator'
    'TransferCallToTarget'   = 'Transfer'
    'Announcement'           = 'Play announcement'
}

# ─────────────────────────────────────────────────────────────────────────────
# FINDING CATALOG
#
# Every condition C.A.R.I.L.L.O.N. can report, keyed by a stable code. The
# Pester suite extracts it by AST lookup and checks it against the codes the
# script raises.
# ─────────────────────────────────────────────────────────────────────────────

$CarillonFindings = @{
    'QueueNoAgents' = @{
        Severity = 'Error'
        Title    = 'Call queue has no agents'
        Summary  = 'Nobody is connected to this queue, so every call waits out the timeout or goes straight to the no-agent action.'
        Remedy   = 'Add users, a group, or a Teams channel as the queue''s agents in the Teams admin center (Voice > Call queues > Agents).'
    }
    'QueueAllOptedOut' = @{
        Severity = 'Error'
        Title    = 'Every agent has opted out of the queue'
        Summary  = 'The queue has agents, but none is taking calls. Calls are not offered to opted-out agents.'
        Remedy   = 'Have at least one agent opt back in (Teams > Settings > Calls > Call queues), or turn off "Agents can opt out" on the queue.'
    }
    'QueueFewAgentsOptedIn' = @{
        Severity = 'Warning'
        Title    = 'Only one agent is taking calls'
        Summary  = 'A single opted-in agent means the queue stops being answered the moment that person is away.'
        Remedy   = 'Confirm the coverage is intended, or add or opt in a second agent.'
    }
    'AgentNotVoiceEnabled' = @{
        Severity = 'Warning'
        Title    = 'Agent is not voice-enabled'
        Summary  = 'The agent has no Enterprise Voice, which usually means no Teams Phone licence. Queue calls from the phone network may not reach them.'
        Remedy   = 'Assign a Teams Phone licence to the agent, or remove them from the queue.'
    }
    'AgentDisabled' = @{
        Severity = 'Warning'
        Title    = 'Agent account is disabled'
        Summary  = 'A disabled account is still listed as an agent but will never answer.'
        Remedy   = 'Remove the account from the queue, or from the group that adds it.'
    }
    'StaleReference' = @{
        Severity = 'Error'
        Title    = 'Routing points at a deleted user or group'
        Summary  = 'An agent list, overflow / timeout target, menu option or operator names an object that no longer exists. Calls sent there fail.'
        Remedy   = 'Edit the queue or attendant and replace the missing target.'
    }
    'OverflowImmediate' = @{
        Severity = 'Warning'
        Title    = 'Queue overflows on the first call'
        Summary  = 'The overflow threshold is 0, so every call takes the overflow action instead of ringing an agent.'
        Remedy   = 'Raise the overflow threshold unless the queue is meant to be bypassed.'
    }
    'CallsDisconnected' = @{
        Severity = 'Info'
        Title    = 'Queue hangs up on overflow or timeout'
        Summary  = 'Callers who hit the limit are disconnected rather than sent to voicemail or another target.'
        Remedy   = 'Consider shared voicemail or a forward so those calls are not lost.'
    }
    'Unreachable' = @{
        Severity = 'Warning'
        Title    = 'Nothing routes calls here'
        Summary  = 'No resource account with a phone number fronts it, and no auto attendant or queue transfers to it.'
        Remedy   = 'Give its resource account a number, or add it as a target in an attendant menu or queue overflow -- or delete it if it is no longer used.'
    }
    'NoAfterHours' = @{
        Severity = 'Info'
        Title    = 'Auto attendant has no after-hours handling'
        Summary  = 'Every call gets the business-hours flow, day and night.'
        Remedy   = 'Fine for a 24/7 line. Otherwise add business hours and an after-hours call flow.'
    }
    'ResourceAccountUnassigned' = @{
        Severity = 'Info'
        Title    = 'Resource account not assigned to anything'
        Summary  = 'The account fronts no call queue or auto attendant. Any number or licence on it is doing nothing.'
        Remedy   = 'Assign it to the queue or attendant it was made for, or delete it and reclaim the number and licence.'
    }
    'GroupsNotResolved' = @{
        Severity = 'Info'
        Title    = 'Group names were not resolved'
        Summary  = 'Microsoft Graph was skipped or unavailable, so groups appear by object ID and agents are not traced to the group that added them.'
        Remedy   = 'Run again without -SkipGraph and consent to Directory.Read.All.'
    }
    'QueryFailed' = @{
        Severity = 'Warning'
        Title    = 'Part of the tenant could not be read'
        Summary  = 'A Teams or Microsoft Graph query failed, so the results below may be incomplete.'
        Remedy   = 'Sign in with Teams Administrator (or Global Reader plus Teams Communications Support Engineer) and retry.'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

$Findings = [System.Collections.Generic.List[object]]::new()

function Add-CarillonFinding {
    param(
        [Parameter(Mandatory)][string]$Code,
        [string]$Subject = '',
        [string]$Detail = ''
    )
    $meta = $CarillonFindings[$Code]
    if (-not $meta) { $meta = @{ Severity = 'Warning'; Title = $Code; Summary = ''; Remedy = '' } }
    [void]$Findings.Add([PSCustomObject]@{
        Code = $Code; Severity = $meta.Severity; Title = $meta.Title; Summary = $meta.Summary
        Remedy = $meta.Remedy; Subject = $Subject; Detail = $Detail
    })
}

function Get-SeverityClass {
    param([string]$Severity)
    switch ($Severity) {
        'Error'   { return 'err'  }
        'Warning' { return 'warn' }
        default   { return 'info' }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────

function Show-CarillonBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host @"

   ██████╗ █████╗ ██████╗ ██╗██╗     ██╗      ██████╗ ███╗   ██╗
  ██╔════╝██╔══██╗██╔══██╗██║██║     ██║     ██╔═══██╗████╗  ██║
  ██║     ███████║██████╔╝██║██║     ██║     ██║   ██║██╔██╗ ██║
  ██║     ██╔══██║██╔══██╗██║██║     ██║     ██║   ██║██║╚██╗██║
  ╚██████╗██║  ██║██║  ██║██║███████╗███████╗╚██████╔╝██║ ╚████║
   ╚═════╝╚═╝  ╚═╝╚═╝  ╚═╝╚═╝╚══════╝╚══════╝ ╚═════╝ ╚═╝  ╚═══╝

"@ -ForegroundColor Cyan
    Write-Host "    C.A.R.I.L.L.O.N. — Catalogs Attendants, Routing, Inbound Lines, Listeners, Overflow & Numbers" -ForegroundColor Cyan
    Write-Host "    Teams Phone Call Queue & Auto Attendant Audit Tool" -ForegroundColor Cyan
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# PURE HELPERS
#
# Queue and attendant objects come from the Teams module as typed objects; the
# tests pass PSCustomObjects or hashtables. Member access reads all of them the
# same way. No I/O until the COLLECTORS section.
# ─────────────────────────────────────────────────────────────────────────────

function Get-CarillonList {
    # The Teams module returns missing collections as $null and single values
    # unwrapped; GUIDs arrive as [guid]. Normalise to a flat string array.
    param($Value)
    if ($null -eq $Value) { return @() }
    return @($Value | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
}

function Get-CarillonName {
    # Display label for a directory object ID -- never the bare GUID when the
    # object resolved. $Names maps ID -> @{ DisplayName; Upn }.
    param([string]$Id, [hashtable]$Names)
    if ($Names -and $Id -and $Names.ContainsKey($Id.ToLowerInvariant())) {
        $n = $Names[$Id.ToLowerInvariant()]
        if ($n.Upn -and $n.Upn -ne $n.DisplayName) { return "$($n.DisplayName) <$($n.Upn)>" }
        return "$($n.DisplayName)"
    }
    return "[unresolved $Id]"
}

function Format-CarillonNumber {
    param([string]$Uri)
    return ("$Uri" -replace '^(?i)tel:', '').Trim()
}

function Format-CarillonTarget {
    # Renders a callable entity -- a queue overflow / timeout target, an
    # attendant menu option or operator -- as readable text. $Endpoints maps a
    # resource-account ID to what it fronts; $Configs maps a queue or attendant
    # identity to its label.
    param($Target, [hashtable]$Names, [hashtable]$Endpoints = @{}, [hashtable]$Configs = @{})
    if ($null -eq $Target -or -not "$($Target.Id)") { return '' }
    $id  = "$($Target.Id)".Trim()
    $key = $id.ToLowerInvariant()
    switch ("$($Target.Type)") {
        'ExternalPstn'    { return "External number $(Format-CarillonNumber $id)" }
        'User'            { return "User $(Get-CarillonName -Id $id -Names $Names)" }
        'SharedVoicemail' { return "Shared voicemail $(Get-CarillonName -Id $id -Names $Names)" }
        'ApplicationEndpoint' {
            if ($Endpoints.ContainsKey($key)) { return "$($Endpoints[$key]) (via $(Get-CarillonName -Id $id -Names $Names))" }
            return "Resource account $(Get-CarillonName -Id $id -Names $Names)"
        }
        'ConfigurationEndpoint' {
            if ($Configs.ContainsKey($key)) { return "$($Configs[$key])" }
            return "[unresolved queue / attendant $id]"
        }
        default { return "$($Target.Type) $(Get-CarillonName -Id $id -Names $Names)" }
    }
}

function Format-CarillonAction {
    # "Forward -> User Jane Doe" / "Disconnect". Queue actions name a target
    # only when they send the call somewhere.
    param([string]$Action, [string]$TargetText)
    if (-not $Action) { return '' }
    if ($TargetText -and $Action -notin $DisconnectActions) { return "$Action -> $TargetText" }
    return $Action
}

function Get-CarillonQueueIssue {
    # The finding codes a queue earns from its own settings and agent list.
    # $Queue carries AgentCount, OptedInCount, OverflowThreshold,
    # OverflowAction and TimeoutAction.
    param($Queue)
    $codes = [System.Collections.Generic.List[string]]::new()
    if ([int]$Queue.AgentCount -eq 0) {
        [void]$codes.Add('QueueNoAgents')
    } elseif ([int]$Queue.OptedInCount -eq 0) {
        [void]$codes.Add('QueueAllOptedOut')
    } elseif ([int]$Queue.OptedInCount -eq 1) {
        [void]$codes.Add('QueueFewAgentsOptedIn')
    }
    if ($null -ne $Queue.OverflowThreshold -and "$($Queue.OverflowThreshold)" -ne '' -and
        [int]$Queue.OverflowThreshold -eq 0 -and "$($Queue.OverflowAction)" -ne 'Queue') {
        [void]$codes.Add('OverflowImmediate')
    }
    if ("$($Queue.OverflowAction)" -in $DisconnectActions -or "$($Queue.TimeoutAction)" -in $DisconnectActions) {
        [void]$codes.Add('CallsDisconnected')
    }
    return $codes.ToArray()
}

function Get-CarillonAgentSource {
    # How an agent came to be on the queue. $GroupMembers maps a group ID to
    # the set of user IDs in it (empty when groups could not be expanded).
    param([string]$AgentId, [string[]]$DirectUsers, [string[]]$Groups, [hashtable]$GroupMembers, [hashtable]$Names, [bool]$Channel)
    $src = [System.Collections.Generic.List[string]]::new()
    if (@($DirectUsers) -contains $AgentId) { [void]$src.Add('Direct') }
    foreach ($g in @($Groups)) {
        $key = "$g".ToLowerInvariant()
        if ($GroupMembers -and $GroupMembers.ContainsKey($key) -and $GroupMembers[$key].Contains($AgentId.ToLowerInvariant())) {
            [void]$src.Add("Group: $(Get-CarillonName -Id $g -Names $Names)")
        }
    }
    if ($src.Count -eq 0) {
        if ($Channel)                   { [void]$src.Add('Teams channel') }
        elseif (@($Groups).Count -gt 0) { [void]$src.Add('Group') }
        else                            { [void]$src.Add('Direct') }
    }
    return ($src -join '; ')
}

function Test-CarillonReachable {
    # A queue or attendant is reachable when one of its resource accounts has a
    # number, or another attendant / queue routes to it.
    param([string[]]$ResourceAccountIds, [hashtable]$NumberedAccounts, [string[]]$ReachedFrom)
    if (@($ReachedFrom | Where-Object { $_ }).Count -gt 0) { return $true }
    foreach ($id in @($ResourceAccountIds)) {
        if ($NumberedAccounts -and $NumberedAccounts.ContainsKey("$id".ToLowerInvariant())) { return $true }
    }
    return $false
}

function Get-CarillonVerdict {
    param([object[]]$FindingList)
    $sev = @($FindingList | ForEach-Object { $_.Severity })
    if ($sev -contains 'Error')   { return [PSCustomObject]@{ Verdict = 'Broken';    Class = 'err'  } }
    if ($sev -contains 'Warning') { return [PSCustomObject]@{ Verdict = 'Attention'; Class = 'warn' } }
    return [PSCustomObject]@{ Verdict = 'Healthy'; Class = 'ok' }
}

# ─────────────────────────────────────────────────────────────────────────────
# MODULES + CONNECTION
# ─────────────────────────────────────────────────────────────────────────────

function Install-CarillonModule {
    param([string]$ModuleName)
    if (Get-Module -ListAvailable -Name $ModuleName) {
        Write-Ok "$ModuleName — installed"
    } else {
        Write-Warn "$ModuleName — NOT found"
        $doInstall = $Unattended
        if (-not $Unattended) {
            $ans = Read-Host "  Install $ModuleName for current user? [Y/N]"
            $doInstall = $ans -match '^[Yy]'
        }
        if (-not $doInstall) { Write-Fail "$ModuleName was not installed."; return $false }
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            Write-Ok "$ModuleName installed."
        } catch {
            Write-Fail "Install failed: $($_.Exception.Message)"
            Write-TKError -ScriptName 'carillon' -Message "$ModuleName install failed: $($_.Exception.Message)" -Category 'Module Install'
            return $false
        }
    }
    try { Import-Module $ModuleName -ErrorAction Stop; return $true }
    catch { Write-Fail "Could not import ${ModuleName}: $($_.Exception.Message)"; return $false }
}

function Connect-CarillonTeams {
    Write-Section "CONNECT TO MICROSOFT TEAMS"
    try {
        $tenant = Get-CsTenant -ErrorAction Stop
        if ($tenant) {
            Write-Ok "Reusing existing Teams session: $($tenant.DisplayName)"
            return [PSCustomObject]@{ Account = ''; Tenant = "$($tenant.DisplayName)" }
        }
    } catch {
        Write-Verbose "No cached Teams session: $($_.Exception.Message)"
    }
    Write-Step "Requesting interactive sign-in..."
    try {
        $conn = Connect-MicrosoftTeams -ErrorAction Stop
        $tenantName = ''
        try { $tenantName = "$((Get-CsTenant -ErrorAction Stop).DisplayName)" } catch { Write-Verbose "Get-CsTenant failed: $($_.Exception.Message)" }
        Write-Ok "Connected as: $($conn.Account)"
        return [PSCustomObject]@{ Account = "$($conn.Account)"; Tenant = $tenantName }
    } catch {
        Write-Fail "Connect-MicrosoftTeams failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'carillon' -Message "Connect-MicrosoftTeams failed: $($_.Exception.Message)" -Category 'Teams Auth'
        return $null
    }
}

function Connect-CarillonGraph {
    Write-Section "CONNECT TO MICROSOFT GRAPH"
    if (-not (Install-CarillonModule -ModuleName 'Microsoft.Graph.Authentication')) { return $false }
    try {
        $ctx = Get-MgContext -ErrorAction Stop
        if ($ctx -and $ctx.Account -and -not @($GraphScopes | Where-Object { $ctx.Scopes -notcontains $_ })) {
            Write-Ok "Reusing existing session: $($ctx.Account)"
            return $true
        }
    } catch {
        Write-Verbose "No cached Graph context: $($_.Exception.Message)"
    }
    Write-Step "Requesting interactive sign-in (used only to name groups)..."
    Write-Info "Scopes: $($GraphScopes -join ', ')"
    try {
        Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
        Write-Ok "Connected as: $((Get-MgContext).Account)"
        return $true
    } catch {
        Write-Warn "Connect-MgGraph failed: $($_.Exception.Message)"
        Write-TKError -ScriptName 'carillon' -Message "Connect-MgGraph failed: $($_.Exception.Message)" -Category 'Graph Auth'
        return $false
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTORS
# ─────────────────────────────────────────────────────────────────────────────

function Get-CarillonPaged {
    # Get-CsCallQueue and Get-CsAutoAttendant return at most 100 objects per
    # call; page with -First / -Skip until a short page comes back.
    param([string]$Command, [string]$Label)
    $items = [System.Collections.Generic.List[object]]::new()
    $skip  = 0
    $page  = 100
    try {
        do {
            $batch = @(& $Command -First $page -Skip $skip -ErrorAction Stop -WarningAction SilentlyContinue)
            foreach ($b in $batch) { [void]$items.Add($b) }
            $skip += $page
        } while ($batch.Count -eq $page)
    } catch {
        Add-CarillonFinding -Code 'QueryFailed' -Subject $Label -Detail $_.Exception.Message
        Write-Warn "$Label query failed: $($_.Exception.Message)"
        return $null
    }
    return , $items.ToArray()
}

function Resolve-CarillonGraphName {
    # directoryObjects/getByIds names users and groups in one call per 1000
    # IDs. Returns the set of IDs Graph was asked about, so the caller can tell
    # "does not exist" from "never asked".
    param([string[]]$Ids, [hashtable]$Names, [hashtable]$Kinds)
    $asked = @{}
    $ids   = @($Ids | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() } | Select-Object -Unique)
    for ($i = 0; $i -lt $ids.Count; $i += 1000) {
        $chunk = @($ids[$i..([math]::Min($i + 999, $ids.Count - 1))])
        try {
            $body = @{ ids = $chunk; types = @('user', 'group') } | ConvertTo-Json -Depth 3
            $resp = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/directoryObjects/getByIds' -Body $body -ContentType 'application/json' -ErrorAction Stop
            foreach ($o in @($resp['value'])) {
                $key = "$($o['id'])".ToLowerInvariant()
                $Names[$key] = [PSCustomObject]@{ DisplayName = "$($o['displayName'])"; Upn = "$($o['userPrincipalName'])" }
                $Kinds[$key] = if ("$($o['@odata.type'])" -match 'group$') { 'Group' } else { 'User' }
            }
            foreach ($c in $chunk) { $asked[$c] = $true }
        } catch {
            Add-CarillonFinding -Code 'QueryFailed' -Subject 'Graph directoryObjects/getByIds' -Detail $_.Exception.Message
            Write-Warn "Graph name lookup failed: $($_.Exception.Message)"
            return $asked
        }
    }
    return $asked
}

function Get-CarillonGroupMember {
    # Transitive user members of a group, as a lower-case ID set; $null when
    # the query fails.
    param([string]$GroupId)
    $set  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $next = "v1.0/groups/$GroupId/transitiveMembers/microsoft.graph.user" + '?$select=id&$top=999'
    try {
        while ($next) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
            foreach ($m in @($resp['value'])) { [void]$set.Add("$($m['id'])") }
            $next = $resp['@odata.nextLink']
        }
    } catch {
        Write-Verbose "Group $GroupId expansion failed: $($_.Exception.Message)"
        return $null
    }
    return , $set
}

function Get-CarillonUser {
    # Get-CsOnlineUser for one ID. Status is Found, Missing (the object does
    # not exist) or Failed (the lookup itself broke -- not evidence of anything).
    param([string]$Id)
    try {
        $u = Get-CsOnlineUser -Identity $Id -ErrorAction Stop -WarningAction SilentlyContinue
        return [PSCustomObject]@{
            Status        = 'Found'
            DisplayName   = "$($u.DisplayName)"
            Upn           = "$($u.UserPrincipalName)"
            VoiceEnabled  = [bool]$u.EnterpriseVoiceEnabled
            AccountEnabled = if ($null -eq $u.AccountEnabled) { $true } else { [bool]$u.AccountEnabled }
        }
    } catch {
        $status = if ("$($_.Exception.Message)" -match 'not found|NotFound|does not exist|couldn''t find|could not find') { 'Missing' } else { 'Failed' }
        return [PSCustomObject]@{ Status = $status; DisplayName = ''; Upn = ''; VoiceEnabled = $null; AccountEnabled = $null }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────

function Get-CarillonMenuTarget {
    # Every callable entity an attendant can send a call to: its operator and
    # each menu option of each call flow.
    param($Attendant)
    $targets = [System.Collections.Generic.List[object]]::new()
    if ($Attendant.Operator -and "$($Attendant.Operator.Id)") { [void]$targets.Add($Attendant.Operator) }
    foreach ($flow in @(@($Attendant.DefaultCallFlow) + @($Attendant.CallFlows) | Where-Object { $_ })) {
        foreach ($opt in @($flow.Menu.MenuOptions | Where-Object { $_ })) {
            if ($opt.CallTarget -and "$($opt.CallTarget.Id)") { [void]$targets.Add($opt.CallTarget) }
        }
    }
    return , $targets.ToArray()
}

function Get-CarillonQueueTarget {
    param($Queue)
    return , @(@($Queue.OverflowActionTarget, $Queue.TimeoutActionTarget, $Queue.NoAgentActionTarget) |
        Where-Object { $_ -and "$($_.Id)" })
}

function Format-CarillonMenu {
    # One line per menu option: "1: Transfer -> Call queue Sales".
    param($Flow, [hashtable]$Names, [hashtable]$Endpoints, [hashtable]$Configs)
    if (-not $Flow -or -not $Flow.Menu) { return @() }
    $lines = foreach ($opt in @($Flow.Menu.MenuOptions | Where-Object { $_ })) {
        $key = ("$($opt.DtmfResponse)" -replace '^Tone', '') -replace '^Star$', '*' -replace '^Pound$', '#'
        if ($key -eq 'Automatic') { $key = 'Auto' }
        $action = if ($MenuActionNames.ContainsKey("$($opt.Action)")) { $MenuActionNames["$($opt.Action)"] } else { "$($opt.Action)" }
        $target = Format-CarillonTarget -Target $opt.CallTarget -Names $Names -Endpoints $Endpoints -Configs $Configs
        if ($target -and "$($opt.Action)" -eq 'TransferCallToTarget') { "${key}: $action -> $target" } else { "${key}: $action" }
    }
    $lines = @($lines | Where-Object { $_ })
    if ($Flow.Menu.DialByNameEnabled) { $lines += 'Dial by name enabled' }
    return $lines
}

function Invoke-CarillonAudit {
    param([bool]$UseGraph)

    Write-Section "CALL QUEUES, AUTO ATTENDANTS AND RESOURCE ACCOUNTS"
    Write-Step "Reading call queues..."
    $queues = Get-CarillonPaged -Command 'Get-CsCallQueue' -Label 'Call queues'
    $queues = @($queues | Where-Object { $_ })
    Write-Step "Reading auto attendants..."
    $attendants = Get-CarillonPaged -Command 'Get-CsAutoAttendant' -Label 'Auto attendants'
    $attendants = @($attendants | Where-Object { $_ })
    Write-Step "Reading resource accounts..."
    $accounts = @()
    try {
        $accounts = @(Get-CsOnlineApplicationInstance -ErrorAction Stop -WarningAction SilentlyContinue | Where-Object { $_ })
    } catch {
        Add-CarillonFinding -Code 'QueryFailed' -Subject 'Resource accounts' -Detail $_.Exception.Message
        Write-Warn "Resource account query failed: $($_.Exception.Message)"
    }
    Write-Info ("Call queues: {0}  |  Auto attendants: {1}  |  Resource accounts: {2}" -f $queues.Count, $attendants.Count, $accounts.Count)

    # ── Names ──
    # Resource accounts come back named; everything else is resolved below.
    $names = @{}
    $kinds = @{}
    $numbered = @{}
    foreach ($ra in $accounts) {
        $key = "$($ra.ObjectId)".ToLowerInvariant()
        $names[$key] = [PSCustomObject]@{ DisplayName = "$($ra.DisplayName)"; Upn = "$($ra.UserPrincipalName)" }
        $kinds[$key] = 'ResourceAccount'
        if ("$($ra.PhoneNumber)") { $numbered[$key] = Format-CarillonNumber "$($ra.PhoneNumber)" }
    }

    $agentIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $groupIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $objectIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($q in $queues) {
        foreach ($a in @($q.Agents | Where-Object { $_ })) { [void]$agentIds.Add("$($a.ObjectId)") }
        foreach ($u in (Get-CarillonList $q.Users)) { [void]$agentIds.Add($u) }
        foreach ($g in (Get-CarillonList $q.DistributionLists)) { [void]$groupIds.Add($g) }
        foreach ($t in (Get-CarillonQueueTarget $q)) {
            if ("$($t.Type)" -in @('User', 'SharedVoicemail', 'ApplicationEndpoint')) { [void]$objectIds.Add("$($t.Id)") }
        }
    }
    foreach ($aa in $attendants) {
        foreach ($t in (Get-CarillonMenuTarget $aa)) {
            if ("$($t.Type)" -in @('User', 'SharedVoicemail', 'ApplicationEndpoint')) { [void]$objectIds.Add("$($t.Id)") }
        }
    }
    foreach ($id in $agentIds) { [void]$objectIds.Add($id) }
    foreach ($id in $groupIds) { [void]$objectIds.Add($id) }

    # $missing holds IDs confirmed not to exist; an ID that simply could not be
    # looked up is never counted as deleted.
    $missing = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $groupMembers = @{}
    if ($UseGraph -and $objectIds.Count -gt 0) {
        Write-Step "Resolving names through Microsoft Graph..."
        $asked = Resolve-CarillonGraphName -Ids @($objectIds) -Names $names -Kinds $kinds
        foreach ($id in $objectIds) {
            $key = $id.ToLowerInvariant()
            if ($asked.ContainsKey($key) -and -not $names.ContainsKey($key)) { [void]$missing.Add($key) }
        }
        foreach ($g in $groupIds) {
            $members = Get-CarillonGroupMember -GroupId $g
            if ($null -ne $members) { $groupMembers[$g.ToLowerInvariant()] = $members }
        }
    }

    # Every agent through Get-CsOnlineUser: names without Graph, and the voice
    # and account state Graph does not report.
    $agentInfo = @{}
    $i = 0
    foreach ($id in $agentIds) {
        $i++
        Write-Progress -Activity 'Reading agents' -Status "$i of $($agentIds.Count)" -PercentComplete ([math]::Min(100, 100 * $i / [math]::Max(1, $agentIds.Count)))
        $info = Get-CarillonUser -Id $id
        $key  = $id.ToLowerInvariant()
        $agentInfo[$key] = $info
        if ($info.Status -eq 'Found') {
            if (-not $names.ContainsKey($key)) { $names[$key] = [PSCustomObject]@{ DisplayName = $info.DisplayName; Upn = $info.Upn } }
            [void]$missing.Remove($key)
        } elseif ($info.Status -eq 'Missing' -and -not $names.ContainsKey($key)) {
            [void]$missing.Add($key)
        }
    }
    Write-Progress -Activity 'Reading agents' -Completed

    # Without Graph, the non-agent user targets still get a name.
    if (-not $UseGraph) {
        foreach ($id in $objectIds) {
            $key = $id.ToLowerInvariant()
            if ($names.ContainsKey($key) -or $agentInfo.ContainsKey($key) -or $groupIds.Contains($id)) { continue }
            $info = Get-CarillonUser -Id $id
            if ($info.Status -eq 'Found')       { $names[$key] = [PSCustomObject]@{ DisplayName = $info.DisplayName; Upn = $info.Upn } }
            elseif ($info.Status -eq 'Missing') { [void]$missing.Add($key) }
        }
    }

    $unresolvedGroups = @($groupIds | Where-Object { -not $names.ContainsKey($_.ToLowerInvariant()) -and -not $missing.Contains($_) })
    if ($unresolvedGroups.Count -gt 0) {
        Add-CarillonFinding -Code 'GroupsNotResolved' -Subject "$($unresolvedGroups.Count) group(s)" -Detail ($unresolvedGroups -join ', ')
    }

    # ── Who fronts what ──
    $endpoints = @{}   # resource account ID -> "Call queue Sales"
    $configs   = @{}   # queue / attendant identity -> "Call queue Sales"
    $raOwner   = @{}   # resource account ID -> list of labels
    foreach ($q in $queues) {
        $label = "Call queue $($q.Name)"
        $configs["$($q.Identity)".ToLowerInvariant()] = $label
        foreach ($ra in (Get-CarillonList $q.ApplicationInstances)) {
            $endpoints[$ra.ToLowerInvariant()] = $label
            if (-not $raOwner.ContainsKey($ra.ToLowerInvariant())) { $raOwner[$ra.ToLowerInvariant()] = [System.Collections.Generic.List[string]]::new() }
            [void]$raOwner[$ra.ToLowerInvariant()].Add($label)
        }
    }
    foreach ($aa in $attendants) {
        $label = "Auto attendant $($aa.Name)"
        $configs["$($aa.Identity)".ToLowerInvariant()] = $label
        foreach ($ra in (Get-CarillonList $aa.ApplicationInstances)) {
            $endpoints[$ra.ToLowerInvariant()] = $label
            if (-not $raOwner.ContainsKey($ra.ToLowerInvariant())) { $raOwner[$ra.ToLowerInvariant()] = [System.Collections.Generic.List[string]]::new() }
            [void]$raOwner[$ra.ToLowerInvariant()].Add($label)
        }
    }

    # ── Who routes to what ──
    $reachedFrom = @{}   # queue / attendant identity -> list of source labels
    function _link { param([string]$ToIdentity, [string]$From)
        $k = $ToIdentity.ToLowerInvariant()
        if (-not $reachedFrom.ContainsKey($k)) { $reachedFrom[$k] = [System.Collections.Generic.List[string]]::new() }
        if (-not $reachedFrom[$k].Contains($From)) { [void]$reachedFrom[$k].Add($From) }
    }
    $sources = @(@($queues | ForEach-Object { [PSCustomObject]@{ Label = "Call queue $($_.Name)"; Targets = (Get-CarillonQueueTarget $_) } }) +
                 @($attendants | ForEach-Object { [PSCustomObject]@{ Label = "Auto attendant $($_.Name)"; Targets = (Get-CarillonMenuTarget $_) } }))
    foreach ($s in $sources) {
        foreach ($t in @($s.Targets)) {
            $tid = "$($t.Id)".ToLowerInvariant()
            if ("$($t.Type)" -eq 'ConfigurationEndpoint') { _link $tid $s.Label; continue }
            if ("$($t.Type)" -ne 'ApplicationEndpoint') { continue }
            foreach ($q in $queues)     { if ((Get-CarillonList $q.ApplicationInstances)  -contains $tid) { _link "$($q.Identity)"  $s.Label } }
            foreach ($aa in $attendants) { if ((Get-CarillonList $aa.ApplicationInstances) -contains $tid) { _link "$($aa.Identity)" $s.Label } }
        }
    }
    function _from { param([string]$Identity) $k = $Identity.ToLowerInvariant(); if ($reachedFrom.ContainsKey($k)) { @($reachedFrom[$k]) } else { @() } }
    function _stale { param([object[]]$Targets)
        @($Targets | Where-Object { "$($_.Type)" -in @('User', 'SharedVoicemail', 'ApplicationEndpoint') -and $missing.Contains("$($_.Id)") } |
            ForEach-Object { "$($_.Type) $($_.Id)" })
    }

    $inScope = { param([string]$ItemName) (-not $Name) -or ($ItemName -like $Name) }

    # ── Call queues ──
    $queueRows = [System.Collections.Generic.List[object]]::new()
    $agentRows = [System.Collections.Generic.List[object]]::new()
    foreach ($q in ($queues | Sort-Object Name)) {
        if (-not (& $inScope "$($q.Name)")) { continue }
        $qName   = "$($q.Name)"
        $direct  = @(Get-CarillonList $q.Users | ForEach-Object { $_.ToLowerInvariant() })
        $groups  = @(Get-CarillonList $q.DistributionLists)
        $channel = [bool]"$($q.ChannelId)"

        # Agents is the effective list (direct users plus expanded groups and
        # channel members) with opt-in state. Older module builds lack it.
        $agentList = @($q.Agents | Where-Object { $_ } | ForEach-Object { [PSCustomObject]@{ Id = "$($_.ObjectId)"; OptIn = $_.OptIn } })
        if ($agentList.Count -eq 0 -and $direct.Count -gt 0) {
            $agentList = @($direct | ForEach-Object { [PSCustomObject]@{ Id = $_; OptIn = $null } })
        }

        $rows = foreach ($a in $agentList) {
            $key  = $a.Id.ToLowerInvariant()
            $info = $agentInfo[$key]
            [PSCustomObject]@{
                Queue          = $qName
                Agent          = if ($names.ContainsKey($key)) { $names[$key].DisplayName } elseif ($missing.Contains($key)) { "[deleted user $($a.Id)]" } else { "[unresolved $($a.Id)]" }
                Upn            = if ($names.ContainsKey($key)) { $names[$key].Upn } else { '' }
                OptedIn        = if ($null -eq $a.OptIn) { 'Unknown' } elseif ($a.OptIn) { 'Yes' } else { 'No' }
                Source         = Get-CarillonAgentSource -AgentId $key -DirectUsers $direct -Groups $groups -GroupMembers $groupMembers -Names $names -Channel $channel
                VoiceEnabled   = if ($info -and $info.Status -eq 'Found') { if ($info.VoiceEnabled) { 'Yes' } else { 'No' } } else { 'Unknown' }
                AccountEnabled = if ($info -and $info.Status -eq 'Found') { if ($info.AccountEnabled) { 'Yes' } else { 'No' } } else { 'Unknown' }
                Missing        = $missing.Contains($key)
                ObjectId       = $a.Id
            }
        }
        $rows = @($rows | Sort-Object Missing, Agent)
        foreach ($r in $rows) { [void]$agentRows.Add($r) }

        $optedIn = @($rows | Where-Object { $_.OptedIn -ne 'No' -and -not $_.Missing }).Count
        $raIds   = @(Get-CarillonList $q.ApplicationInstances)
        $numbers = @($raIds | Where-Object { $numbered.ContainsKey($_.ToLowerInvariant()) } | ForEach-Object { $numbered[$_.ToLowerInvariant()] })
        $from    = @(_from "$($q.Identity)")

        $row = [PSCustomObject]@{
            Name               = $qName
            Numbers            = $numbers -join ', '
            ResourceAccounts   = (@($raIds | ForEach-Object { Get-CarillonName -Id $_ -Names $names }) -join ', ')
            Routing            = if ($RoutingMethodNames.ContainsKey("$($q.RoutingMethod)")) { $RoutingMethodNames["$($q.RoutingMethod)"] } else { "$($q.RoutingMethod)" }
            PresenceBased      = [bool]$q.PresenceBasedRouting
            ConferenceMode     = [bool]$q.ConferenceMode
            AllowOptOut        = [bool]$q.AllowOptOut
            AlertTime          = "$($q.AgentAlertTime)"
            AgentSource        = @(@(if ($direct.Count) { "$($direct.Count) direct" }) + @($groups | ForEach-Object { "group $(Get-CarillonName -Id $_ -Names $names)" }) + @(if ($channel) { 'Teams channel' })) -join ', '
            AgentCount         = $rows.Count
            OptedInCount       = $optedIn
            OverflowThreshold  = $q.OverflowThreshold
            OverflowAction     = "$($q.OverflowAction)"
            Overflow           = Format-CarillonAction -Action "$($q.OverflowAction)" -TargetText (Format-CarillonTarget -Target $q.OverflowActionTarget -Names $names -Endpoints $endpoints -Configs $configs)
            TimeoutThreshold   = "$($q.TimeoutThreshold)"
            TimeoutAction      = "$($q.TimeoutAction)"
            Timeout            = Format-CarillonAction -Action "$($q.TimeoutAction)" -TargetText (Format-CarillonTarget -Target $q.TimeoutActionTarget -Names $names -Endpoints $endpoints -Configs $configs)
            NoAgent            = Format-CarillonAction -Action "$($q.NoAgentAction)" -TargetText (Format-CarillonTarget -Target $q.NoAgentActionTarget -Names $names -Endpoints $endpoints -Configs $configs)
            ReachedFrom        = $from -join ', '
            Agents             = $rows
        }
        [void]$queueRows.Add($row)

        foreach ($code in (Get-CarillonQueueIssue $row)) {
            switch ($code) {
                'QueueNoAgents'         { Add-CarillonFinding -Code 'QueueNoAgents' -Subject $qName }
                'QueueAllOptedOut'      { Add-CarillonFinding -Code 'QueueAllOptedOut' -Subject $qName -Detail "$($rows.Count) agent(s), all opted out" }
                'QueueFewAgentsOptedIn' { Add-CarillonFinding -Code 'QueueFewAgentsOptedIn' -Subject $qName -Detail (@($rows | Where-Object { $_.OptedIn -ne 'No' -and -not $_.Missing } | ForEach-Object { $_.Agent }) -join ', ') }
                'OverflowImmediate'     { Add-CarillonFinding -Code 'OverflowImmediate' -Subject $qName -Detail $row.Overflow }
                'CallsDisconnected'     { Add-CarillonFinding -Code 'CallsDisconnected' -Subject $qName -Detail ("Overflow: {0}; Timeout: {1}" -f $row.Overflow, $row.Timeout) }
            }
        }
        $noVoice = @($rows | Where-Object { $_.VoiceEnabled -eq 'No' })
        if ($noVoice.Count) { Add-CarillonFinding -Code 'AgentNotVoiceEnabled' -Subject $qName -Detail (@($noVoice | ForEach-Object { $_.Agent }) -join ', ') }
        $disabled = @($rows | Where-Object { $_.AccountEnabled -eq 'No' })
        if ($disabled.Count) { Add-CarillonFinding -Code 'AgentDisabled' -Subject $qName -Detail (@($disabled | ForEach-Object { $_.Agent }) -join ', ') }
        $stale = @(@($rows | Where-Object { $_.Missing } | ForEach-Object { "agent $($_.ObjectId)" }) +
                   @($groups | Where-Object { $missing.Contains($_) } | ForEach-Object { "group $_" }) +
                   @(_stale (Get-CarillonQueueTarget $q)))
        if ($stale.Count) { Add-CarillonFinding -Code 'StaleReference' -Subject "Call queue $qName" -Detail ($stale -join ', ') }
        if (-not (Test-CarillonReachable -ResourceAccountIds $raIds -NumberedAccounts $numbered -ReachedFrom $from)) {
            Add-CarillonFinding -Code 'Unreachable' -Subject "Call queue $qName" -Detail $(if ($raIds.Count) { 'Resource account has no phone number' } else { 'No resource account assigned' })
        }
    }

    # ── Auto attendants ──
    $attendantRows = [System.Collections.Generic.List[object]]::new()
    foreach ($aa in ($attendants | Sort-Object Name)) {
        if (-not (& $inScope "$($aa.Name)")) { continue }
        $aName   = "$($aa.Name)"
        $raIds   = @(Get-CarillonList $aa.ApplicationInstances)
        $numbers = @($raIds | Where-Object { $numbered.ContainsKey($_.ToLowerInvariant()) } | ForEach-Object { $numbered[$_.ToLowerInvariant()] })
        $from    = @(_from "$($aa.Identity)")

        # CallHandlingAssociations tie each non-default call flow to a schedule.
        $schedules = @{}
        foreach ($s in @($aa.Schedules | Where-Object { $_ })) { $schedules["$($s.Id)"] = "$($s.Name)" }
        $flows = @{}
        foreach ($f in @($aa.CallFlows | Where-Object { $_ })) { $flows["$($f.Id)"] = $f }
        $otherFlows = foreach ($assoc in @($aa.CallHandlingAssociations | Where-Object { $_ })) {
            $flow  = $flows["$($assoc.CallFlowId)"]
            $sched = if ($schedules.ContainsKey("$($assoc.ScheduleId)")) { $schedules["$($assoc.ScheduleId)"] } else { "$($assoc.ScheduleId)" }
            [PSCustomObject]@{
                Type     = "$($assoc.Type)"
                Schedule = $sched
                Enabled  = if ($null -eq $assoc.Enabled) { $true } else { [bool]$assoc.Enabled }
                Menu     = @(Format-CarillonMenu -Flow $flow -Names $names -Endpoints $endpoints -Configs $configs)
            }
        }
        $otherFlows = @($otherFlows | Where-Object { $_ })

        $row = [PSCustomObject]@{
            Name             = $aName
            Numbers          = $numbers -join ', '
            ResourceAccounts = (@($raIds | ForEach-Object { Get-CarillonName -Id $_ -Names $names }) -join ', ')
            Language         = "$($aa.LanguageId)"
            TimeZone         = "$($aa.TimeZoneId)"
            Operator         = Format-CarillonTarget -Target $aa.Operator -Names $names -Endpoints $endpoints -Configs $configs
            VoiceInput       = [bool]$aa.VoiceResponseEnabled
            BusinessHours    = @(Format-CarillonMenu -Flow $aa.DefaultCallFlow -Names $names -Endpoints $endpoints -Configs $configs)
            OtherFlows       = $otherFlows
            ReachedFrom      = $from -join ', '
        }
        [void]$attendantRows.Add($row)

        if (@($otherFlows | Where-Object { $_.Type -eq 'AfterHours' }).Count -eq 0) { Add-CarillonFinding -Code 'NoAfterHours' -Subject $aName }
        $stale = @(_stale (Get-CarillonMenuTarget $aa))
        if ($stale.Count) { Add-CarillonFinding -Code 'StaleReference' -Subject "Auto attendant $aName" -Detail ($stale -join ', ') }
        if (-not (Test-CarillonReachable -ResourceAccountIds $raIds -NumberedAccounts $numbered -ReachedFrom $from)) {
            Add-CarillonFinding -Code 'Unreachable' -Subject "Auto attendant $aName" -Detail $(if ($raIds.Count) { 'Resource account has no phone number' } else { 'No resource account assigned' })
        }
    }

    # ── Resource accounts ──
    $accountRows = foreach ($ra in ($accounts | Sort-Object DisplayName)) {
        $key   = "$($ra.ObjectId)".ToLowerInvariant()
        $labels = if ($raOwner.ContainsKey($key)) { @($raOwner[$key]) } else { @() }
        if ($Name) {
            # Filtered run: only the accounts fronting a queue or attendant in scope.
            if (@($labels | Where-Object { ($_ -replace '^(Call queue|Auto attendant) ', '') -like $Name }).Count -eq 0) { continue }
        } elseif ($labels.Count -eq 0) {
            Add-CarillonFinding -Code 'ResourceAccountUnassigned' -Subject "$($ra.DisplayName)" -Detail $(if ("$($ra.PhoneNumber)") { "Holds $(Format-CarillonNumber "$($ra.PhoneNumber)")" } else { '' })
        }
        [PSCustomObject]@{
            Name       = "$($ra.DisplayName)"
            Upn        = "$($ra.UserPrincipalName)"
            Number     = Format-CarillonNumber "$($ra.PhoneNumber)"
            Type       = if ($ResourceAccountApps.ContainsKey("$($ra.ApplicationId)".ToLowerInvariant())) { $ResourceAccountApps["$($ra.ApplicationId)".ToLowerInvariant()] } else { 'Other' }
            AssignedTo = $labels -join ', '
        }
    }

    return [PSCustomObject]@{
        Queues           = $queueRows.ToArray()
        Agents           = $agentRows.ToArray()
        Attendants       = $attendantRows.ToArray()
        Accounts         = @($accountRows | Where-Object { $_ })
        QueueTotal       = $queues.Count
        AttendantTotal   = $attendants.Count
        AccountTotal     = $accounts.Count
        GraphUsed        = $UseGraph
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# CONSOLE + HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Show-CarillonInventory {
    param([object]$Audit)

    Write-Section "CALL QUEUES"
    if ($Audit.Queues.Count -eq 0) { Write-Info "No call queues$(if ($Name) { " matching '$Name'" })." }
    foreach ($q in $Audit.Queues) {
        Write-Host ""
        Write-Host ("  {0}" -f $q.Name) -ForegroundColor $ColorSchema.Header -NoNewline
        Write-Host ("   {0}  |  {1}  |  {2} of {3} agent(s) opted in" -f $(if ($q.Numbers) { $q.Numbers } else { 'no number' }), $q.Routing, $q.OptedInCount, $q.AgentCount) -ForegroundColor $ColorSchema.Info
        if ($q.AgentCount -eq 0) { Write-Fail "No agents"; continue }
        foreach ($a in $q.Agents) {
            $color = if ($a.Missing -or $a.AccountEnabled -eq 'No') { $ColorSchema.Error } elseif ($a.OptedIn -eq 'No' -or $a.VoiceEnabled -eq 'No') { $ColorSchema.Warning } else { $ColorSchema.Success }
            $state = if ($a.OptedIn -eq 'No') { 'opted out' } elseif ($a.OptedIn -eq 'Yes') { 'opted in' } else { 'opt-in unknown' }
            Write-Host ("      {0,-34} {1,-40} {2,-15} {3}" -f $a.Agent, $a.Upn, $state, $a.Source) -ForegroundColor $color
        }
    }

    Write-Section "AUTO ATTENDANTS"
    if ($Audit.Attendants.Count -eq 0) { Write-Info "No auto attendants$(if ($Name) { " matching '$Name'" })." }
    foreach ($aa in $Audit.Attendants) {
        Write-Host ""
        Write-Host ("  {0}" -f $aa.Name) -ForegroundColor $ColorSchema.Header -NoNewline
        Write-Host ("   {0}  |  {1}" -f $(if ($aa.Numbers) { $aa.Numbers } else { 'no number' }), $aa.TimeZone) -ForegroundColor $ColorSchema.Info
        if ($aa.Operator) { Write-Host "      Operator: $($aa.Operator)" -ForegroundColor $ColorSchema.Info }
        foreach ($line in $aa.BusinessHours) { Write-Host "      $line" -ForegroundColor $ColorSchema.Info }
        foreach ($f in $aa.OtherFlows) {
            Write-Host ("      [{0}: {1}]" -f $f.Type, $f.Schedule) -ForegroundColor $ColorSchema.Accent
            foreach ($line in $f.Menu) { Write-Host "        $line" -ForegroundColor $ColorSchema.Info }
        }
    }
}

function Show-CarillonFindings {
    Write-Section "FINDINGS"
    if ($Findings.Count -eq 0) { Write-Ok "Nothing to report."; return }
    $order = @{ 'Error' = 0; 'Warning' = 1; 'Info' = 2 }
    foreach ($f in ($Findings | Sort-Object { $order[$_.Severity] })) {
        $line = "$($f.Title)" + $(if ($f.Subject) { " -- $($f.Subject)" } else { '' })
        switch ($f.Severity) { 'Error' { Write-Fail $line } 'Warning' { Write-Warn $line } default { Write-Info $line } }
    }
}

function Build-CarillonReport {
    param([object]$Audit, [object]$Verdict, [object]$Connection)

    $cfg        = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($cfg.OrgName)) { "$($cfg.OrgName) -- " } else { '' }
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant     = if ($Connection.Tenant) { $Connection.Tenant } elseif ($Connection.Account -match '@(.+)$') { $Matches[1] } else { 'Tenant' }
    $order      = @{ 'Error' = 0; 'Warning' = 1; 'Info' = 2 }
    function _yn { param([bool]$Value) if ($Value) { 'Yes' } else { 'No' } }
    function _cell { param([string]$Text) if ($Text) { EscHtml $Text } else { '-' } }

    $fRows = [System.Text.StringBuilder]::new()
    if ($Findings.Count -eq 0) { [void]$fRows.Append("<tr><td colspan='4'>Nothing to report.</td></tr>") }
    foreach ($f in ($Findings | Sort-Object { $order[$_.Severity] })) {
        [void]$fRows.Append(
            "<tr><td><span class='tk-badge-$(Get-SeverityClass $f.Severity)'>$(EscHtml $f.Severity)</span></td>" +
            "<td><strong>$(EscHtml $f.Title)</strong><br/>$(EscHtml $f.Summary)</td>" +
            "<td>$(EscHtml $f.Subject)" + $(if ($f.Detail) { "<br/><span class='tk-mono'>$(EscHtml $f.Detail)</span>" } else { '' }) + "</td>" +
            "<td>$(EscHtml $f.Remedy)</td></tr>")
    }

    $qRows = [System.Text.StringBuilder]::new()
    if ($Audit.Queues.Count -eq 0) { [void]$qRows.Append("<tr><td colspan='7'>No call queues.</td></tr>") }
    foreach ($q in $Audit.Queues) {
        $cls = if ($q.AgentCount -eq 0 -or $q.OptedInCount -eq 0) { 'err' } elseif ($q.OptedInCount -eq 1) { 'warn' } else { 'ok' }
        $opts = @(@(if ($q.PresenceBased) { 'presence-based' }) + @(if ($q.ConferenceMode) { 'conference mode' }) + @(if ($q.AllowOptOut) { 'opt-out allowed' })) -join ', '
        [void]$qRows.Append(
            "<tr><td><strong>$(EscHtml $q.Name)</strong><br/><span class='tk-mono'>$(_cell $q.ResourceAccounts)</span></td>" +
            "<td class='tk-mono'>$(_cell $q.Numbers)</td>" +
            "<td>$(EscHtml $q.Routing)$(if ($opts) { "<br/><span class='tk-mono'>$(EscHtml $opts)</span>" })<br/><span class='tk-mono'>ring $(EscHtml $q.AlertTime) s</span></td>" +
            "<td><span class='tk-badge-$cls'>$($q.OptedInCount) / $($q.AgentCount)</span><br/><span class='tk-mono'>$(_cell $q.AgentSource)</span></td>" +
            "<td>after $(EscHtml "$($q.OverflowThreshold)") call(s): $(_cell $q.Overflow)</td>" +
            "<td>after $(EscHtml $q.TimeoutThreshold) s: $(_cell $q.Timeout)$(if ($q.NoAgent) { "<br/><span class='tk-mono'>no agents: $(EscHtml $q.NoAgent)</span>" })</td>" +
            "<td>$(_cell $q.ReachedFrom)</td></tr>")
    }

    $agentCards = [System.Text.StringBuilder]::new()
    if ($Audit.Queues.Count -eq 0) { [void]$agentCards.Append("<div class='tk-card'>No call queues.</div>") }
    foreach ($q in $Audit.Queues) {
        [void]$agentCards.Append("<div class='tk-card'><div class='tk-card-header'><span class='tk-card-label'>$(EscHtml $q.Name)</span><span class='tk-mono'>$($q.OptedInCount) of $($q.AgentCount) opted in</span></div>")
        if ($q.AgentCount -eq 0) {
            [void]$agentCards.Append("<div class='tk-info-box'><span class='tk-info-label'>No agents</span> Nobody is connected to this queue.</div></div>")
            continue
        }
        [void]$agentCards.Append("<table class='tk-table'><thead><tr><th>Agent</th><th>Sign-in name</th><th>Opted in</th><th>Voice enabled</th><th>Account</th><th>Added through</th></tr></thead><tbody>")
        foreach ($a in $q.Agents) {
            $opt   = switch ($a.OptedIn) { 'Yes' { "<span class='tk-badge-ok'>Yes</span>" } 'No' { "<span class='tk-badge-warn'>No</span>" } default { "<span class='tk-badge-info'>Unknown</span>" } }
            $voice = switch ($a.VoiceEnabled) { 'Yes' { "<span class='tk-badge-ok'>Yes</span>" } 'No' { "<span class='tk-badge-warn'>No</span>" } default { "<span class='tk-badge-info'>Unknown</span>" } }
            $acct  = if ($a.Missing) { "<span class='tk-badge-err'>Deleted</span>" } else { switch ($a.AccountEnabled) { 'Yes' { "<span class='tk-badge-ok'>Enabled</span>" } 'No' { "<span class='tk-badge-err'>Disabled</span>" } default { "<span class='tk-badge-info'>Unknown</span>" } } }
            [void]$agentCards.Append("<tr><td><strong>$(EscHtml $a.Agent)</strong></td><td class='tk-mono'>$(_cell $a.Upn)</td><td>$opt</td><td>$voice</td><td>$acct</td><td>$(EscHtml $a.Source)</td></tr>")
        }
        [void]$agentCards.Append("</tbody></table></div>")
    }

    $aRows = [System.Text.StringBuilder]::new()
    if ($Audit.Attendants.Count -eq 0) { [void]$aRows.Append("<tr><td colspan='5'>No auto attendants.</td></tr>") }
    foreach ($aa in $Audit.Attendants) {
        $bh = if ($aa.BusinessHours.Count) { (@($aa.BusinessHours | ForEach-Object { EscHtml $_ }) -join '<br/>') } else { '-' }
        $other = if ($aa.OtherFlows.Count) {
            (@($aa.OtherFlows | ForEach-Object {
                "<strong>$(EscHtml $_.Type)</strong> <span class='tk-mono'>($(EscHtml $_.Schedule)$(if (-not $_.Enabled) { ', disabled' }))</span><br/>" +
                (@($_.Menu | ForEach-Object { EscHtml $_ }) -join '<br/>')
            }) -join '<br/><br/>')
        } else { "<span class='tk-badge-info'>None -- business hours 24/7</span>" }
        [void]$aRows.Append(
            "<tr><td><strong>$(EscHtml $aa.Name)</strong><br/><span class='tk-mono'>$(_cell $aa.ResourceAccounts)</span><br/><span class='tk-mono'>$(EscHtml $aa.Language) | $(EscHtml $aa.TimeZone)$(if ($aa.VoiceInput) { ' | voice input' })</span></td>" +
            "<td class='tk-mono'>$(_cell $aa.Numbers)</td>" +
            "<td>$bh$(if ($aa.Operator) { "<br/><span class='tk-mono'>Operator: $(EscHtml $aa.Operator)</span>" })</td>" +
            "<td>$other</td><td>$(_cell $aa.ReachedFrom)</td></tr>")
    }

    $rRows = [System.Text.StringBuilder]::new()
    if (@($Audit.Accounts).Count -eq 0) { [void]$rRows.Append("<tr><td colspan='5'>No resource accounts.</td></tr>") }
    foreach ($r in $Audit.Accounts) {
        $assigned = if ($r.AssignedTo) { EscHtml $r.AssignedTo } else { "<span class='tk-badge-info'>Unassigned</span>" }
        [void]$rRows.Append("<tr><td><strong>$(EscHtml $r.Name)</strong></td><td class='tk-mono'>$(_cell $r.Upn)</td><td class='tk-mono'>$(_cell $r.Number)</td><td>$(EscHtml $r.Type)</td><td>$assigned</td></tr>")
    }

    $errCount   = @($Findings | Where-Object { $_.Severity -eq 'Error' }).Count
    $warnCount  = @($Findings | Where-Object { $_.Severity -eq 'Warning' }).Count
    $agentTotal = @($Audit.Agents | Select-Object -ExpandProperty ObjectId -Unique).Count

    $meta = [ordered]@{
        'Tenant'    = $tenant
        'Generated' = $reportDate
        'Verdict'   = $Verdict.Verdict
    }
    if ($Connection.Account) { $meta['Signed in as'] = $Connection.Account }
    if ($Name) { $meta['Name filter'] = $Name }

    $htmlHead = Get-TKHtmlHead `
        -Title      'C.A.R.I.L.L.O.N. Call Queue & Auto Attendant Report' `
        -ScriptName 'C.A.R.I.L.L.O.N.' `
        -Subtitle   "${orgPrefix}Teams Phone Call Queues & Auto Attendants -- $tenant" `
        -MetaItems  $meta `
        -NavItems   @('Findings', 'Call Queues', 'Queue Agents', 'Auto Attendants', 'Resource Accounts')

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Verdict</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Audit.Queues.Count)</div><div class="tk-summary-lbl">Call Queues</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$agentTotal</div><div class="tk-summary-lbl">Distinct Agents</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Audit.Attendants.Count)</div><div class="tk-summary-lbl">Auto Attendants</div></div>
    <div class="tk-summary-card $(if ($errCount) { 'err' } else { 'ok' })"><div class="tk-summary-num">$errCount</div><div class="tk-summary-lbl">Errors</div></div>
    <div class="tk-summary-card $(if ($warnCount) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$warnCount</div><div class="tk-summary-lbl">Warnings</div></div>
  </div>

  <div class="tk-section" id="s01">
    <div class="tk-section-title"><span class="tk-section-num">01</span> Findings</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Severity</th><th>Finding</th><th>Where</th><th>Remedy</th></tr></thead>
      <tbody>$($fRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s02">
    <div class="tk-section-title"><span class="tk-section-num">02</span> Call Queues</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Queue</th><th>Number</th><th>Routing</th><th>Agents (opted in / total)</th><th>Overflow</th><th>Timeout</th><th>Reached from</th></tr></thead>
      <tbody>$($qRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s03">
    <div class="tk-section-title"><span class="tk-section-num">03</span> Queue Agents</div>
    $($agentCards.ToString())
  </div>

  <div class="tk-section" id="s04">
    <div class="tk-section-title"><span class="tk-section-num">04</span> Auto Attendants</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Auto attendant</th><th>Number</th><th>Business hours</th><th>After hours / holidays</th><th>Reached from</th></tr></thead>
      <tbody>$($aRows.ToString())</tbody></table></div>
  </div>

  <div class="tk-section" id="s05">
    <div class="tk-section-title"><span class="tk-section-num">05</span> Resource Accounts</div>
    <div class="tk-card"><table class="tk-table">
      <thead><tr><th>Name</th><th>Sign-in name</th><th>Number</th><th>Type</th><th>Assigned to</th></tr></thead>
      <tbody>$($rRows.ToString())</tbody></table></div>
  </div>

"@ + (Get-TKHtmlFoot -ScriptName 'C.A.R.I.L.L.O.N. v5.1')

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-CarillonRun {
    $Findings.Clear()

    Write-Section "MODULE CHECK"
    if (-not (Install-CarillonModule -ModuleName 'MicrosoftTeams')) { if ($Unattended) { exit 1 }; return }
    $conn = Connect-CarillonTeams
    if (-not $conn) { if ($Unattended) { exit 1 }; return }

    $useGraph = $false
    if (-not $SkipGraph) {
        $useGraph = Connect-CarillonGraph
        if (-not $useGraph) { Write-Warn "Continuing without Graph -- users are named through Teams, groups are not." }
    }

    $audit = Invoke-CarillonAudit -UseGraph $useGraph
    Show-CarillonInventory -Audit $audit
    Show-CarillonFindings

    $verdict = Get-CarillonVerdict -FindingList $Findings.ToArray()
    Write-Section "VERDICT"
    switch ($verdict.Class) { 'err' { Write-Fail $verdict.Verdict } 'warn' { Write-Warn $verdict.Verdict } default { Write-Ok $verdict.Verdict } }

    Add-TKNote -Text ("CARILLON Teams Phone audit: verdict {0}; {1} call queue(s), {2} auto attendant(s); {3} finding(s)." -f $verdict.Verdict, $audit.Queues.Count, $audit.Attendants.Count, $Findings.Count) -Category 'Info' -ScriptName 'carillon'

    $logDir = Resolve-LogDirectory -FallbackPath $ScriptPath
    $stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'

    Write-Step "Writing queue agent CSV..."
    try {
        $csvPath = Join-Path $logDir "CARILLON_Agents_$stamp.csv"
        @($audit.Agents | Select-Object Queue, Agent, @{ n = 'UserPrincipalName'; e = { $_.Upn } }, OptedIn, VoiceEnabled, AccountEnabled, Source, ObjectId) |
            Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
        Write-Ok "Agent CSV saved: $csvPath"
    } catch {
        Write-Fail "Could not save CSV: $($_.Exception.Message)"
        Write-TKError -ScriptName 'carillon' -Message "CSV save failed: $($_.Exception.Message)" -Category 'Report'
    }

    Write-Step "Generating HTML report..."
    $html    = Build-CarillonReport -Audit $audit -Verdict $verdict -Connection $conn
    $outPath = Join-Path $logDir "CARILLON_$stamp.html"
    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Show-TKReportResult -Path $outPath -Unattended:$Unattended
    } catch {
        Write-Fail "Could not save report: $($_.Exception.Message)"
        Write-TKError -ScriptName 'carillon' -Message "Report save failed: $($_.Exception.Message)" -Category 'Report'
    }
}

Show-CarillonBanner
Invoke-CarillonRun

if (-not $Unattended) { Read-Host "  Press Enter to exit" | Out-Null }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath -and -not (Test-Path (Join-Path $PSScriptRoot '.git'))) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
