<#
.SYNOPSIS
    R.I.T.U.A.L. — Runs Integrated Tool Usage in Automation Loops
    Workflow Orchestrator for the Technician Toolkit (PowerShell 5.1+)

.DESCRIPTION
    Runs an ordered sequence of toolkit scripts as a single named recipe.
    Ships with built-in recipes for common scenarios (Onboard, Retire,
    HealthCheck, TenantSweep) and accepts custom PSD1 recipe files. Each
    step is invoked with -Unattended so the whole recipe runs without
    prompts; step duration, exit code, and any new files written into the
    log directory are captured and rolled up into a single HTML summary.

    Built-in recipes:
      Onboard      -- New machine bring-up:
                      COVENANT -> SIGIL -> CONJURE -> CIPHER -> AUSPEX -> ARTIFACT
      Retire       -- Pre-reimage / pre-disposal:
                      TETHER -> EXHUME -> ARCHIVE -> CLEANSE
      HealthCheck  -- Quarterly machine review:
                      AUSPEX -> WARD -> THRESHOLD -> AUGUR -> GARGOYLE -> ARTIFACT
      TenantSweep  -- Cloud tenant posture (one sign-in, four reports):
                      TALISMAN -> RELIQUARY -> GOLEM -> WRAITH

.USAGE
    PS C:\> .\ritual.ps1                                    # Interactive menu
    PS C:\> .\ritual.ps1 -Recipe HealthCheck                # Run a named recipe
    PS C:\> .\ritual.ps1 -RecipeFile .\custom.psd1          # Run a custom recipe file
    PS C:\> .\ritual.ps1 -Recipe Retire -ContinueOnError    # Ignore per-step failures

.NOTES
    Version : 3.0

#>

param(
    [switch]$Unattended,
    [switch]$Transcript,
    [ValidateSet('Onboard','Retire','HealthCheck','TenantSweep')]
    [string]$Recipe = '',
    [ValidateScript({ [string]::IsNullOrWhiteSpace($_) -or (Test-Path -LiteralPath $_) })]
    [string]$RecipeFile = '',
    [switch]$ContinueOnError
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

function Show-RitualBanner {
    if (-not $Unattended) { Clear-Host }
    Write-Host ""
    Write-Host "  R.I.T.U.A.L. — Runs Integrated Tool Usage in Automation Loops" -ForegroundColor Cyan
    Write-Host "  Workflow Orchestrator for the Technician Toolkit  v3.0" -ForegroundColor Cyan
    Write-Host ""
}

# ─── Recipes, execution, and rollup appended below ───

# ─────────────────────────────────────────────────────────────────────────────
# BUILT-IN RECIPES
# ─────────────────────────────────────────────────────────────────────────────

# Each recipe is an ordered list of step hashtables. Fields:
#   Tool         -- file name of the script, resolved against $ScriptPath
#   Args         -- array of arguments forwarded to the script
#   StopOnError  -- override: $true means abort the recipe on step failure
#                   regardless of the run-wide -ContinueOnError flag
#   Label        -- pretty name shown in console / HTML
$script:BuiltInRecipes = @{
    'Onboard' = @{
        Name        = 'New Machine Onboard'
        Description = 'New-machine bring-up: identity, hardening, software, encryption, diagnostics, certificates.'
        Steps       = @(
            @{ Label = 'Machine onboarding';        Tool = 'covenant.ps1';   Args = @('-Unattended'); StopOnError = $true  }
            @{ Label = 'Security baseline';         Tool = 'sigil.ps1';      Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Core software install';     Tool = 'conjure.ps1';    Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'BitLocker enable';          Tool = 'cipher.ps1';     Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'System diagnostics';        Tool = 'auspex.ps1';     Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Certificate health';        Tool = 'artifact.ps1';   Args = @('-Unattended'); StopOnError = $false }
        )
    }
    'Retire' = @{
        Name        = 'Pre-Reimage / Disposal'
        Description = 'Pre-reimage workflow: validate cloud data before local cleanup.'
        Steps       = @(
            @{ Label = 'OneDrive KFM validator';    Tool = 'tether.ps1';  Args = @('-Unattended'); StopOnError = $true  }
            @{ Label = 'Outlook PST discovery';     Tool = 'exhume.ps1';  Args = @('-Unattended'); StopOnError = $true  }
            @{ Label = 'Profile backup (ZIP)';      Tool = 'archive.ps1'; Args = @('-Unattended'); StopOnError = $true  }
            @{ Label = 'Temp / cache cleanup';      Tool = 'cleanse.ps1'; Args = @('-Unattended'); StopOnError = $false }
        )
    }
    'HealthCheck' = @{
        Name        = 'Quarterly Machine Review'
        Description = 'Read-only sweep for a point-in-time machine health snapshot.'
        Steps       = @(
            @{ Label = 'System diagnostics';        Tool = 'auspex.ps1';    Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Account audit';             Tool = 'ward.ps1';      Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Disk space';                Tool = 'threshold.ps1'; Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Disk hardware (SMART)';     Tool = 'augur.ps1';     Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Services & tasks';          Tool = 'gargoyle.ps1';  Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Certificate health';        Tool = 'artifact.ps1';  Args = @('-Unattended'); StopOnError = $false }
        )
    }
    'TenantSweep' = @{
        Name        = 'Cloud Tenant Posture'
        Description = 'Full tenant posture in one sign-in sequence: Azure, M365 licensing, Intune, Entra ID hygiene.'
        Steps       = @(
            @{ Label = 'Azure assessment';          Tool = 'talisman.ps1';  Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'M365 license audit';        Tool = 'reliquary.ps1'; Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Intune / MDM compliance';   Tool = 'golem.ps1';     Args = @('-Unattended'); StopOnError = $false }
            @{ Label = 'Entra ID identity hygiene'; Tool = 'wraith.ps1';    Args = @('-Unattended'); StopOnError = $false }
        )
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# RECIPE RESOLVER
# ─────────────────────────────────────────────────────────────────────────────

function Resolve-Recipe {
    param(
        [string]$NamedRecipe,
        [string]$PathToFile
    )

    if ($PathToFile) {
        try {
            $data = Import-PowerShellDataFile -Path $PathToFile -ErrorAction Stop
        } catch {
            Write-Fail "Could not parse recipe file '$PathToFile': $($_.Exception.Message)"
            Write-TKError -ScriptName 'ritual' -Message "Recipe file parse failure: $($_.Exception.Message)" -Category 'Recipe Load'
            return $null
        }

        # A recipe file is just a hashtable with the same shape as the built-ins.
        if (-not $data.Steps -or $data.Steps.Count -eq 0) {
            Write-Fail "Recipe file '$PathToFile' has no Steps defined."
            return $null
        }

        if (-not $data.Name)        { $data.Name        = (Split-Path -Leaf $PathToFile) }
        if (-not $data.Description) { $data.Description = 'Custom recipe' }
        return $data
    }

    if ($NamedRecipe -and $script:BuiltInRecipes.ContainsKey($NamedRecipe)) {
        return $script:BuiltInRecipes[$NamedRecipe]
    }

    return $null
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

# Snapshot the log directory so we can diff it after a step runs and attribute
# any new files (HTML reports, CSV logs, etc.) to the step that produced them.
function Get-LogDirSnapshot {
    param([string]$LogRoot)
    if (-not (Test-Path $LogRoot)) { return @{} }
    $map = @{}
    Get-ChildItem -LiteralPath $LogRoot -File -ErrorAction SilentlyContinue |
        ForEach-Object { $map[$_.FullName] = $_.LastWriteTimeUtc.Ticks }
    return $map
}

function Get-NewLogFiles {
    param([hashtable]$Before, [string]$LogRoot)
    $after = Get-LogDirSnapshot -LogRoot $LogRoot
    $new = foreach ($path in $after.Keys) {
        if (-not $Before.ContainsKey($path) -or $Before[$path] -ne $after[$path]) {
            Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
        }
    }
    return @($new | Where-Object { $_ })
}

function Invoke-RecipeStep {
    param(
        [int]$StepNumber,
        [hashtable]$Step,
        [string]$LogRoot
    )

    $toolPath = Join-Path $ScriptPath $Step.Tool
    $label    = if ($Step.Label) { $Step.Label } else { $Step.Tool }

    Write-Section ("STEP {0} - {1}" -f $StepNumber, $label)
    Write-Info "Tool    : $($Step.Tool)"
    Write-Info "Args    : $($Step.Args -join ' ')"
    Write-Host ""

    if (-not (Test-Path $toolPath)) {
        Write-Fail "Tool file not found: $toolPath"
        return [PSCustomObject]@{
            StepNumber = $StepNumber
            Label      = $label
            Tool       = $Step.Tool
            Status     = 'Missing'
            DurationSec = 0
            ExitCode   = $null
            NewFiles   = @()
            Error      = "Tool file not found at $toolPath"
        }
    }

    $beforeSnap = Get-LogDirSnapshot -LogRoot $LogRoot
    $start      = Get-Date
    $errorMsg   = $null
    $status     = 'Succeeded'
    $exitCode   = 0

    try {
        # Use & with the script path so each step runs in its own variable scope.
        # The child script's -Unattended flag lives in $Step.Args; RITUAL itself is
        # always unattended toward children regardless of its own -Unattended state.
        & $toolPath @($Step.Args)
        if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) {
            $exitCode = $LASTEXITCODE
            $status   = 'Failed'
            $errorMsg = "Child script exited with code $exitCode"
        }
    } catch {
        $status   = 'Failed'
        $errorMsg = $_.Exception.Message
        Write-Fail "Step threw: $errorMsg"
        Write-TKError -ScriptName 'ritual' -Message "Step '$label' ($($Step.Tool)) failed: $errorMsg" -Category 'Recipe Step'
    }

    $duration = [math]::Round(((Get-Date) - $start).TotalSeconds, 1)
    $newFiles = Get-NewLogFiles -Before $beforeSnap -LogRoot $LogRoot

    if ($status -eq 'Succeeded') {
        Write-Ok ("Step finished in {0}s." -f $duration)
        if ($newFiles.Count -gt 0) {
            Write-Info ("Produced {0} log/report file(s) in the log directory." -f $newFiles.Count)
        }
    } else {
        Write-Fail ("Step failed after {0}s: {1}" -f $duration, $errorMsg)
    }

    Write-Host ""

    return [PSCustomObject]@{
        StepNumber  = $StepNumber
        Label       = $label
        Tool        = $Step.Tool
        Status      = $status
        DurationSec = $duration
        ExitCode    = $exitCode
        NewFiles    = @($newFiles)
        Error       = $errorMsg
        StopOnError = [bool]$Step.StopOnError
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# HTML ROLLUP REPORT
# ─────────────────────────────────────────────────────────────────────────────

function Build-RollupHtml {
    param(
        [hashtable]$RecipeData,
        [array]$Results,
        [datetime]$StartTime,
        [datetime]$EndTime
    )

    $tkCfg     = Get-TKConfig
    $orgPrefix = if (-not [string]::IsNullOrWhiteSpace($tkCfg.OrgName)) { "$(EscHtml $tkCfg.OrgName) -- " } else { '' }

    $totalSec   = [math]::Round(($EndTime - $StartTime).TotalSeconds, 1)
    $succeeded  = @($Results | Where-Object { $_.Status -eq 'Succeeded' }).Count
    $failed     = @($Results | Where-Object { $_.Status -eq 'Failed' }).Count
    $skipped    = @($Results | Where-Object { $_.Status -eq 'Skipped' }).Count
    $missing    = @($Results | Where-Object { $_.Status -eq 'Missing' }).Count

    $overallClass = if ($failed -gt 0 -or $missing -gt 0) { 'err' }
                    elseif ($skipped -gt 0) { 'warn' }
                    else { 'ok' }
    $overallLabel = if ($failed -gt 0 -or $missing -gt 0) { 'RECIPE FAILED' }
                    elseif ($skipped -gt 0) { 'RECIPE PARTIAL' }
                    else { 'RECIPE SUCCEEDED' }

    $stepRows = [System.Text.StringBuilder]::new()
    foreach ($r in $Results) {
        $badge = switch ($r.Status) {
            'Succeeded' { "<span class='tk-badge-ok'>Succeeded</span>" }
            'Failed'    { "<span class='tk-badge-err'>Failed</span>" }
            'Skipped'   { "<span class='tk-badge-warn'>Skipped</span>" }
            'Missing'   { "<span class='tk-badge-err'>Missing</span>" }
            default     { "<span class='tk-badge-info'>$(EscHtml $r.Status)</span>" }
        }

        $artifactList = if ($r.NewFiles.Count -eq 0) {
            "<span class='tk-badge-info'>(none)</span>"
        } else {
            ($r.NewFiles | ForEach-Object {
                # Link out to the file with a relative file:// URI so the rollup
                # HTML can open each child report in the browser.
                $uri = [System.Uri]::new($_.FullName).AbsoluteUri
                "<a href='$uri'>$(EscHtml $_.Name)</a>"
            }) -join '<br>'
        }

        $errCell = if ($r.Error) { "<span class='tk-badge-err'>$(EscHtml $r.Error)</span>" } else { '' }

        [void]$stepRows.Append("<tr><td>$($r.StepNumber)</td><td>$(EscHtml $r.Label)</td><td><code>$(EscHtml $r.Tool)</code></td><td>$badge</td><td>$($r.DurationSec)s</td><td>$artifactList</td><td>$errCell</td></tr>`n")
    }

    $htmlHead = Get-TKHtmlHead `
        -Title      'R.I.T.U.A.L. Recipe Execution Report' `
        -ScriptName 'R.I.T.U.A.L.' `
        -Subtitle   "${orgPrefix}$(EscHtml $RecipeData.Name)" `
        -MetaItems  ([ordered]@{
            'Machine'     = $env:COMPUTERNAME
            'Run As'      = "$env:USERDOMAIN\$env:USERNAME"
            'Recipe'      = $RecipeData.Name
            'Started'     = $StartTime.ToString('yyyy-MM-dd HH:mm:ss')
            'Finished'    = $EndTime.ToString('yyyy-MM-dd HH:mm:ss')
            'Total'       = "${totalSec}s"
        }) `
        -NavItems   @('Overall', 'Steps')

    $htmlFoot = Get-TKHtmlFoot -ScriptName 'R.I.T.U.A.L. v3.0'

    $html = $htmlHead + @"

  <div class="tk-summary-row">
    <div class="tk-summary-card $overallClass"><div class="tk-summary-num">$overallLabel</div><div class="tk-summary-lbl">Outcome</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">$($Results.Count)</div><div class="tk-summary-lbl">Total Steps</div></div>
    <div class="tk-summary-card ok"><div class="tk-summary-num">$succeeded</div><div class="tk-summary-lbl">Succeeded</div></div>
    <div class="tk-summary-card $(if ($failed -gt 0) { 'err' } else { 'ok' })"><div class="tk-summary-num">$failed</div><div class="tk-summary-lbl">Failed</div></div>
    <div class="tk-summary-card $(if ($skipped -gt 0) { 'warn' } else { 'ok' })"><div class="tk-summary-num">$skipped</div><div class="tk-summary-lbl">Skipped</div></div>
    <div class="tk-summary-card info"><div class="tk-summary-num">${totalSec}s</div><div class="tk-summary-lbl">Total Duration</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Recipe -- $(EscHtml $RecipeData.Name)</span></div>
    <div class="tk-card"><div class="tk-info-box">$(EscHtml $RecipeData.Description)</div></div>
  </div>

  <div class="tk-section">
    <div class="tk-card-header"><span class="tk-section-title">Step Results</span><span class="tk-section-num">$($Results.Count) step(s)</span></div>
    <div class="tk-card">
      <table class="tk-table">
        <thead><tr><th>#</th><th>Label</th><th>Tool</th><th>Status</th><th>Duration</th><th>Artifacts</th><th>Error</th></tr></thead>
        <tbody>$($stepRows.ToString())</tbody>
      </table>
    </div>
  </div>

"@ + $htmlFoot

    return $html
}

function Export-RollupHtml {
    param([hashtable]$RecipeData, [array]$Results, [datetime]$StartTime, [datetime]$EndTime)

    $html      = Build-RollupHtml -RecipeData $RecipeData -Results $Results -StartTime $StartTime -EndTime $EndTime
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outPath   = Join-Path (Resolve-LogDirectory -FallbackPath $ScriptPath) "RITUAL_${timestamp}.html"

    try {
        [System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
        Write-Ok "Rollup report saved: $outPath"
        if (-not $Unattended) {
            Write-Step "Opening in default browser..."
            Start-Process $outPath
        }
    } catch {
        Write-Fail "Could not save rollup report: $($_.Exception.Message)"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# RECIPE DRIVER
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-Recipe {
    param([hashtable]$RecipeData)

    Show-RitualBanner
    Write-Section ("RECIPE: " + $RecipeData.Name)
    Write-Info $RecipeData.Description
    Write-Host ""

    $logRoot  = Resolve-LogDirectory -FallbackPath $ScriptPath
    $results  = [System.Collections.Generic.List[object]]::new()
    $abort    = $false
    $stepIdx  = 0
    $start    = Get-Date

    foreach ($step in $RecipeData.Steps) {
        $stepIdx++

        if ($abort) {
            Write-Section ("STEP {0} - {1}" -f $stepIdx, $step.Label)
            Write-Warn "Skipped (prior step aborted the recipe)."
            $results.Add([PSCustomObject]@{
                StepNumber  = $stepIdx
                Label       = $step.Label
                Tool        = $step.Tool
                Status      = 'Skipped'
                DurationSec = 0
                ExitCode    = $null
                NewFiles    = @()
                Error       = 'Prior step aborted the recipe'
                StopOnError = [bool]$step.StopOnError
            })
            continue
        }

        $result = Invoke-RecipeStep -StepNumber $stepIdx -Step $step -LogRoot $logRoot
        $results.Add($result) | Out-Null

        if ($result.Status -ne 'Succeeded') {
            $shouldAbort = if ($result.StopOnError) { $true } else { -not $ContinueOnError }
            if ($shouldAbort) {
                Write-Fail "Aborting remaining steps (step failed and continue-on-error is off)."
                Write-Host ""
                $abort = $true
            }
        }
    }

    $end = Get-Date

    # Final console summary
    $total   = $results.Count
    $ok      = @($results | Where-Object { $_.Status -eq 'Succeeded' }).Count
    $bad     = @($results | Where-Object { $_.Status -eq 'Failed' -or $_.Status -eq 'Missing' }).Count
    $skipped = @($results | Where-Object { $_.Status -eq 'Skipped' }).Count

    Write-Section "RECIPE SUMMARY"
    $color = if ($bad -gt 0) { $C.Error } elseif ($skipped -gt 0) { $C.Warning } else { $C.Success }
    Write-Host ("  Recipe   : {0}" -f $RecipeData.Name) -ForegroundColor $C.Info
    Write-Host ("  Total    : {0} steps" -f $total) -ForegroundColor $C.Info
    Write-Host ("  Succeeded: {0}" -f $ok)      -ForegroundColor $C.Success
    Write-Host ("  Failed   : {0}" -f $bad)     -ForegroundColor $(if ($bad -gt 0) { $C.Error } else { $C.Info })
    Write-Host ("  Skipped  : {0}" -f $skipped) -ForegroundColor $(if ($skipped -gt 0) { $C.Warning } else { $C.Info })
    Write-Host ("  Duration : {0}s" -f [math]::Round(($end - $start).TotalSeconds, 1)) -ForegroundColor $C.Info
    Write-Host ""

    Export-RollupHtml -RecipeData $RecipeData -Results @($results) -StartTime $start -EndTime $end

    return (@($results | Where-Object { $_.Status -eq 'Failed' -or $_.Status -eq 'Missing' }).Count -eq 0)
}

# ─────────────────────────────────────────────────────────────────────────────
# MENU
# ─────────────────────────────────────────────────────────────────────────────

function Show-Menu {
    Show-RitualBanner
    Write-Host "  Built-in recipes:" -ForegroundColor $C.Header
    Write-Host ""
    $i = 0
    $menuMap = @{}
    foreach ($name in $script:BuiltInRecipes.Keys) {
        $i++
        $menuMap["$i"] = $name
        $r = $script:BuiltInRecipes[$name]
        Write-Host ("  [{0}]  {1}" -f $i, $r.Name) -ForegroundColor $C.Info
        Write-Host ("        $($r.Description)") -ForegroundColor $C.Info
        Write-Host ("        Steps: " + (($r.Steps | ForEach-Object { $_.Tool -replace '\.ps1$','' }) -join ' -> ')) -ForegroundColor $C.Info
        Write-Host ""
    }
    Write-Host "  [F]  Load a custom recipe file (.psd1)" -ForegroundColor $C.Info
    Write-Host "  [Q]  Quit" -ForegroundColor $C.Info
    Write-Host ""
    return $menuMap
}

# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if ($Unattended -or $Recipe -or $RecipeFile) {
    $chosen = Resolve-Recipe -NamedRecipe $Recipe -PathToFile $RecipeFile
    if (-not $chosen) {
        if ($Unattended -and -not $Recipe -and -not $RecipeFile) {
            Write-Fail "Unattended mode requires -Recipe or -RecipeFile."
        } else {
            Write-Fail "Could not resolve the requested recipe."
        }
        if ($Transcript) { Stop-TKTranscript }
        exit 1
    }
    $ok = Invoke-Recipe -RecipeData $chosen
    if ($Transcript) { Stop-TKTranscript }
    if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
    exit $(if ($ok) { 0 } else { 1 })
}

do {
    $menuMap = Show-Menu
    $choice  = (Read-Host "  Select recipe").Trim().ToUpper()

    if ($choice -eq 'Q') { break }

    if ($choice -eq 'F') {
        $path = Read-Host "  Path to .psd1 recipe file"
        if (-not (Test-Path -LiteralPath $path)) {
            Write-Warn "File not found: $path"
            Start-Sleep -Seconds 1
            continue
        }
        $chosen = Resolve-Recipe -PathToFile $path
    } elseif ($menuMap.ContainsKey($choice)) {
        $chosen = Resolve-Recipe -NamedRecipe $menuMap[$choice]
    } else {
        Write-Warn "Invalid selection."
        Start-Sleep -Milliseconds 800
        continue
    }

    if ($chosen) {
        Invoke-Recipe -RecipeData $chosen | Out-Null
        Read-Host "  Press Enter to return to menu"
    }

} while ($true)

if ($Transcript) { Stop-TKTranscript }
if ($PSCommandPath) { Remove-Item -Path $PSCommandPath -Force -ErrorAction SilentlyContinue }
