<#
.SYNOPSIS
    P.Y.R.E. — Power-Yield Reliability Evaluator
    Laptop Battery Health Audit Tool for PowerShell 5.1+

.DESCRIPTION
    Surfaces the three numbers that matter for a laptop battery's
    retirement decision: design capacity, current full-charge capacity,
    and cycle count. Pulls values from the ROOT\WMI battery classes --
    BatteryStaticData, BatteryFullChargedCapacity, BatteryCycleCount,
    BatteryStatus -- and correlates against Win32_Battery for the user-
    friendly device name. Reports health as a percentage of design,
    applies industry thresholds (80 / 60 for capacity, 300 / 500 for
    cycles), and emits a dark-themed HTML report with a red / yellow /
    green replacement recommendation.

.USAGE
    PS C:\> .\pyre.ps1                    # Interactive run
    PS C:\> .\pyre.ps1 -Unattended        # Silent: export HTML and exit

.NOTES
    Version : 3.0

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

function Show-PyreBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  P.Y.R.E. — Power-Yield Reliability Evaluator" -ForegroundColor Cyan
    Write-Host "  Laptop Battery Health Audit Tool  v3.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Collectors, verdict, and report builders appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# BATTERY DATA COLLECTION
# ─────────────────────────────────────────────────────────────────────────────

# ROOT\WMI battery classes key off InstanceName so multiple batteries (rare
# but possible on gaming laptops) get the same joining key. Win32_Battery
# provides the device caption that a technician recognises; the ROOT\WMI
# classes provide the energy-accounting figures powercfg uses internally.
function Get-BatteryHealthRecords {
    $win32 = @()
    try { $win32 = @(Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue) } catch {
        # Desktops without a battery will naturally return nothing; surface that as "no battery" later.
    }

    $static = @()
    $full   = @()
    $cycle  = @()
    $status = @()
    try { $static = @(Get-CimInstance -Namespace 'ROOT\WMI' -ClassName BatteryStaticData           -ErrorAction SilentlyContinue) } catch { }
    try { $full   = @(Get-CimInstance -Namespace 'ROOT\WMI' -ClassName BatteryFullChargedCapacity  -ErrorAction SilentlyContinue) } catch { }
    try { $cycle  = @(Get-CimInstance -Namespace 'ROOT\WMI' -ClassName BatteryCycleCount           -ErrorAction SilentlyContinue) } catch { }
    try { $status = @(Get-CimInstance -Namespace 'ROOT\WMI' -ClassName BatteryStatus               -ErrorAction SilentlyContinue) } catch { }

    if ($win32.Count -eq 0 -and $static.Count -eq 0) {
        return @()
    }

    # The union of all instance names tells us how many battery records to emit.
    $instances = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($set in @($static, $full, $cycle, $status)) {
        foreach ($item in $set) {
            if ($item.InstanceName) { [void]$instances.Add($item.InstanceName) }
        }
    }

    # Fallback: no ROOT\WMI data at all. Emit the Win32_Battery record(s) with
    # nulls for the detail fields so the tool still produces a report.
    if ($instances.Count -eq 0) {
        return @($win32 | ForEach-Object {
            [PSCustomObject]@{
                InstanceName   = $_.DeviceID
                Name           = $_.Name
                Caption        = $_.Caption
                Chemistry      = Convert-BatteryChemistry $_.Chemistry
                DesignCapacity = $null
                FullCapacity   = $null
                CycleCount     = $null
                HealthPercent  = $null
                ChargeRate     = $null
                DischargeRate  = $null
                Voltage        = $null
                Charging       = $null
                Discharging    = $null
                StatusText     = Convert-BatteryStatus $_.BatteryStatus
            }
        })
    }

    $rows = foreach ($inst in $instances) {
        $s = $static | Where-Object { $_.InstanceName -eq $inst } | Select-Object -First 1
        $f = $full   | Where-Object { $_.InstanceName -eq $inst } | Select-Object -First 1
        $c = $cycle  | Where-Object { $_.InstanceName -eq $inst } | Select-Object -First 1
        $t = $status | Where-Object { $_.InstanceName -eq $inst } | Select-Object -First 1

        # Match the instance to a Win32_Battery device for the friendly caption.
        # Instance names look like "PNP0C0A\1_0"; Win32_Battery.DeviceID looks like
        # "1_0" on most machines. Use a suffix match.
        $wi = $win32 | Where-Object { $inst -like "*$($_.DeviceID)*" } | Select-Object -First 1

        $design = if ($s) { [int64]$s.DesignedCapacity } else { $null }
        $fullCap = if ($f) { [int64]$f.FullChargedCapacity } else { $null }
        $cycles = if ($c) { [int]$c.CycleCount } else { $null }

        $health = $null
        if ($design -and $fullCap -and $design -gt 0) {
            $health = [math]::Round(($fullCap / $design) * 100, 1)
        }

        [PSCustomObject]@{
            InstanceName   = $inst
            Name           = if ($wi) { $wi.Name } else { $inst }
            Caption        = if ($wi) { $wi.Caption } else { '' }
            Chemistry      = if ($wi) { Convert-BatteryChemistry $wi.Chemistry } else { 'Unknown' }
            DesignCapacity = $design
            FullCapacity   = $fullCap
            CycleCount     = $cycles
            HealthPercent  = $health
            ChargeRate     = if ($t) { [int]$t.ChargeRate } else { $null }
            DischargeRate  = if ($t) { [int]$t.DischargeRate } else { $null }
            Voltage        = if ($t) { [int]$t.Voltage } else { $null }
            Charging       = if ($t) { [bool]$t.Charging } else { $null }
            Discharging    = if ($t) { [bool]$t.Discharging } else { $null }
            StatusText     = if ($wi) { Convert-BatteryStatus $wi.BatteryStatus } else { 'Unknown' }
        }
    }

    return @($rows)
}

function Convert-BatteryChemistry {
    param($code)
    switch ([int]$code) {
        1 { 'Other' }       2 { 'Unknown' }  3 { 'Lead Acid' }  4 { 'Nickel Cadmium' }
        5 { 'Nickel Metal Hydride' }  6 { 'Lithium-ion' }  7 { 'Zinc air' }  8 { 'Lithium Polymer' }
        default { 'Unknown' }
    }
}

function Convert-BatteryStatus {
    param($code)
    switch ([int]$code) {
        1 { 'Discharging' }       2 { 'AC Power (charging/charged)' }  3 { 'Fully Charged' }
        4 { 'Low' }               5 { 'Critical' }                     6 { 'Charging' }
        7 { 'Charging + High' }   8 { 'Charging + Low' }               9 { 'Charging + Critical' }
        10 { 'Undefined' }        11 { 'Partially Charged' }
        default { 'Unknown' }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# VERDICT
# ─────────────────────────────────────────────────────────────────────────────

# Industry-consensus thresholds. Manufacturers warrant ~80% capacity at 300
# cycles; below 60% the laptop's unplugged runtime is noticeably short.
$script:CapOkPct   = 80.0
$script:CapWarnPct = 60.0
$script:CycOk      = 300
$script:CycWarn    = 500

function Get-BatteryClass {
    param([PSCustomObject]$B)
    if (-not $B.HealthPercent) { return 'info' }
    if ($B.HealthPercent -ge $script:CapOkPct)    { return 'ok' }
    elseif ($B.HealthPercent -ge $script:CapWarnPct) { return 'warn' }
    else { return 'err' }
}

function Get-CycleClass {
    param([PSCustomObject]$B)
    if ($null -eq $B.CycleCount) { return 'info' }
    if ($B.CycleCount -lt $script:CycOk)   { return 'ok' }
    elseif ($B.CycleCount -lt $script:CycWarn) { return 'warn' }
    else { return 'err' }
}

function Get-Verdict {
    param([array]$Batteries)

    if ($Batteries.Count -eq 0) {
        return [PSCustomObject]@{
            Verdict = 'NO BATTERY'
            Class   = 'info'
            Issues  = @()
            Warns   = @('No battery hardware detected — this is expected on desktops and servers.')
            Worst   = $null
        }
    }

    $issues = [System.Collections.Generic.List[string]]::new()
    $warns  = [System.Collections.Generic.List[string]]::new()
    $worstClass = 'ok'

    foreach ($b in $Batteries) {
        $capClass = Get-BatteryClass -B $b
        $cycClass = Get-CycleClass   -B $b

        if ($capClass -eq 'err') {
            $issues.Add("Battery '$($b.Name)' is at $($b.HealthPercent)% of design capacity — below the $script:CapWarnPct% replacement threshold.")
            $worstClass = 'err'
        } elseif ($capClass -eq 'warn' -and $worstClass -ne 'err') {
            $warns.Add("Battery '$($b.Name)' is at $($b.HealthPercent)% of design capacity — plan replacement within the next upgrade cycle.")
            $worstClass = 'warn'
        }

        if ($cycClass -eq 'err') {
            $issues.Add("Battery '$($b.Name)' has $($b.CycleCount) charge cycles — past the $script:CycWarn-cycle comfort limit.")
            $worstClass = 'err'
        } elseif ($cycClass -eq 'warn' -and $worstClass -ne 'err') {
            $warns.Add("Battery '$($b.Name)' has $($b.CycleCount) charge cycles — approaching the $script:CycWarn-cycle limit.")
            if ($worstClass -ne 'err') { $worstClass = 'warn' }
        }

        if ($null -eq $b.HealthPercent -and $null -eq $b.CycleCount) {
            $warns.Add("Battery '$($b.Name)' reports no design/cycle data — the firmware exposes only basic status. Replacement decisions require powercfg /batteryreport.")
        }
    }

    $verdict = switch ($worstClass) {
        'err'  { 'REPLACE NOW' }
        'warn' { 'REPLACEMENT SOON' }
        'info' { 'DATA INCOMPLETE' }
        default { 'HEALTHY' }
    }

    return [PSCustomObject]@{
        Verdict = $verdict
        Class   = $worstClass
        Issues  = @($issues)
        Warns   = @($warns)
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-HtmlReport {
    param([array]$Batteries, $Verdict)

    $reportDate = Get-Date -Format 'MMMM d, yyyy HH:mm'
    $machine    = $env:COMPUTERNAME
    $tkCfg      = Get-TKConfig
    $orgPrefix  = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $bestHealth = if ($Batteries.Count -gt 0) {
        ($Batteries | Where-Object { $_.HealthPercent } | Measure-Object -Property HealthPercent -Maximum).Maximum
    } else { $null }

    $worstHealth = if ($Batteries.Count -gt 0) {
        ($Batteries | Where-Object { $_.HealthPercent } | Measure-Object -Property HealthPercent -Minimum).Minimum
    } else { $null }

    $maxCycles = if ($Batteries.Count -gt 0) {
        ($Batteries | Where-Object { $null -ne $_.CycleCount } | Measure-Object -Property CycleCount -Maximum).Maximum
    } else { $null }

    # Battery detail rows
    $rows = [System.Text.StringBuilder]::new()
    if ($Batteries.Count -eq 0) {
        [void]$rows.Append("<tr><td colspan='8' class='tk-badge-info' style='text-align:center;'>No battery hardware detected.</td></tr>")
    } else {
        foreach ($b in $Batteries) {
            $capClass = Get-BatteryClass -B $b
            $cycClass = Get-CycleClass   -B $b
            $capBadge = if ($null -eq $b.HealthPercent) { "<span class='tk-badge-info'>n/a</span>" } else { "<span class='tk-badge-$capClass'>$($b.HealthPercent)%</span>" }
            $cycBadge = if ($null -eq $b.CycleCount) { "<span class='tk-badge-info'>n/a</span>" } else { "<span class='tk-badge-$cycClass'>$($b.CycleCount)</span>" }
            $design   = if ($b.DesignCapacity) { "$([math]::Round($b.DesignCapacity / 1000, 1)) Wh" } else { 'n/a' }
            $full     = if ($b.FullCapacity)   { "$([math]::Round($b.FullCapacity / 1000, 1)) Wh" } else { 'n/a' }
            $voltage  = if ($b.Voltage)        { "$([math]::Round($b.Voltage / 1000, 2)) V" } else { 'n/a' }
            [void]$rows.Append("<tr><td>$(EscHtml $b.Name)</td><td>$(EscHtml $b.Chemistry)</td><td>$design</td><td>$full</td><td>$capBadge</td><td>$cycBadge</td><td>$voltage</td><td>$(EscHtml $b.StatusText)</td></tr>`n")
        }
    }

    $verdictBlock = [System.Text.StringBuilder]::new()
    foreach ($i in $Verdict.Issues) { [void]$verdictBlock.Append("<li class='tk-badge-err'>$(EscHtml $i)</li>`n") }
    foreach ($w in $Verdict.Warns)  { [void]$verdictBlock.Append("<li class='tk-badge-warn'>$(EscHtml $w)</li>`n") }
    if ($Verdict.Issues.Count -eq 0 -and $Verdict.Warns.Count -eq 0) {
        [void]$verdictBlock.Append("<li class='tk-badge-ok'>All batteries are within healthy thresholds.</li>")
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'P.Y.R.E. Battery Health Report' `
        -ScriptName 'P.Y.R.E.' `
        -Subtitle   "${orgPrefix}Laptop Battery Health -- $machine" `
        -MetaItems  ([ordered]@{
            'Machine'   = $machine
            'Generated' = $reportDate
            'Verdict'   = $Verdict.Verdict
            'Batteries' = $Batteries.Count
        }) `
        -NavItems   @('Verdict', 'Batteries')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'P.Y.R.E. v3.0'

    $bestCard  = if ($null -ne $bestHealth)  { "$bestHealth%"  } else { 'n/a' }
    $worstCard = if ($null -ne $worstHealth) { "$worstHealth%" } else { 'n/a' }
    $cycCard   = if ($null -ne $maxCycles)   { "$maxCycles" } else { 'n/a' }

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $($Verdict.Class)"><div class="tk-summary-num">$(EscHtml $Verdict.Verdict)</div><div class="tk-summary-lbl">Battery Verdict</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Batteries.Count)</div><div class="tk-summary-lbl">Batteries</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$bestCard</div><div class="tk-summary-lbl">Best Health</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$worstCard</div><div class="tk-summary-lbl">Worst Health</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$cycCard</div><div class="tk-summary-lbl">Max Cycle Count</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$script:CapOkPct% / $script:CycOk</div><div class="tk-summary-lbl">OK Thresholds</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Verdict &amp; Findings</span><span class="tk-section-num">$(EscHtml $Verdict.Verdict)</span></div>
    <div class="tk-card"><ul class="tk-info-box" style="list-style:none;padding-left:0;">$($verdictBlock.ToString())</ul></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Battery Details</span><span class="tk-section-num">$($Batteries.Count) battery/batteries</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>Name</th><th>Chemistry</th><th>Design</th><th>Full Charge</th><th>Health %</th><th>Cycles</th><th>Voltage</th><th>Status</th></tr></thead>
        <tbody>$($rows.ToString())</tbody>
      </table>
      <div class="tk-info-box" style="margin-top:18px;">
        <span class="tk-info-label">Thresholds</span>
        Capacity: &ge; $script:CapOkPct% healthy, $script:CapWarnPct-$script:CapOkPct% plan replacement, &lt; $script:CapWarnPct% replace now.&nbsp;&nbsp;
        Cycles: &lt; $script:CycOk healthy, $script:CycOk-$script:CycWarn plan replacement, &gt;= $script:CycWarn replace now.
      </div>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Show-PyreBanner

Write-Section "BATTERY INVENTORY"
$batteries = Get-BatteryHealthRecords

if ($batteries.Count -eq 0) {
    Write-Warn "No battery hardware detected (desktop / server / VM)."
} else {
    foreach ($b in $batteries) {
        Write-Host "  $($b.Name)  ($($b.Chemistry))" -ForegroundColor $C.Header
        $design = if ($b.DesignCapacity) { "$([math]::Round($b.DesignCapacity / 1000, 1)) Wh" } else { 'n/a' }
        $full   = if ($b.FullCapacity)   { "$([math]::Round($b.FullCapacity / 1000, 1)) Wh" } else { 'n/a' }
        Write-Host ("    Design     : {0}" -f $design) -ForegroundColor $C.Info
        Write-Host ("    Full charge: {0}" -f $full)   -ForegroundColor $C.Info
        if ($null -ne $b.HealthPercent) {
            $cls = Get-BatteryClass -B $b
            $col = switch ($cls) { 'ok' { $C.Success } 'warn' { $C.Warning } 'err' { $C.Error } default { $C.Info } }
            Write-Host ("    Health     : {0}%" -f $b.HealthPercent) -ForegroundColor $col
        }
        if ($null -ne $b.CycleCount) {
            $cls = Get-CycleClass -B $b
            $col = switch ($cls) { 'ok' { $C.Success } 'warn' { $C.Warning } 'err' { $C.Error } default { $C.Info } }
            Write-Host ("    Cycles     : {0}" -f $b.CycleCount) -ForegroundColor $col
        }
        Write-Host ("    Status     : {0}" -f $b.StatusText) -ForegroundColor $C.Info
    }
}
Write-Host ""

$verdict = Get-Verdict -Batteries $batteries

Write-Section "BATTERY VERDICT"
$vColor = switch ($verdict.Class) { 'ok' { $C.Success } 'warn' { $C.Warning } 'err' { $C.Error } default { $C.Info } }
Write-Host "  $($verdict.Verdict)" -ForegroundColor $vColor
foreach ($i in $verdict.Issues) { Write-Host "    [!!] $i" -ForegroundColor $C.Error }
foreach ($w in $verdict.Warns)  { Write-Host "    [~ ] $w" -ForegroundColor $C.Warning }
if ($verdict.Issues.Count -eq 0 -and $verdict.Warns.Count -eq 0) {
    Write-Host "    [+ ] All batteries within healthy thresholds." -ForegroundColor $C.Success
}
Write-Host ""

Write-Step "Generating HTML report..."
$html      = Build-HtmlReport -Batteries $batteries -Verdict $verdict
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "PYRE_${timestamp}.html"

try {
    [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
    Write-Ok "Report saved: $outPath"
    if (-not $Unattended) {
        Write-Step "Opening in default browser..."
        Start-Process $outPath
    }
} catch {
    Write-Fail "Could not save report: $($_.Exception.Message)"
}

if (-not $Unattended) { Read-Host "  Press Enter to exit" }
if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
