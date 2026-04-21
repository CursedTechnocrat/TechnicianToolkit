#Requires -Modules Pester
<#
.SYNOPSIS
    Pester tests for the TechnicianToolkit shared module (TechnicianToolkit.psm1).
    Tests cover the pure utility functions that do not require admin rights or
    live Windows APIs, so they can run in CI without elevated privileges.
#>

BeforeAll {
    $ModulePath = Join-Path $PSScriptRoot '..\TechnicianToolkit.psm1'
    Import-Module $ModulePath -Force
}

# ─────────────────────────────────────────────────────────────────────────────
# EscHtml
# ─────────────────────────────────────────────────────────────────────────────
Describe 'EscHtml' {
    It 'escapes ampersands' {
        EscHtml 'a & b' | Should -Be 'a &amp; b'
    }
    It 'escapes less-than' {
        EscHtml '<script>' | Should -Be '&lt;script&gt;'
    }
    It 'escapes double quotes' {
        EscHtml '"hello"' | Should -Be '&quot;hello&quot;'
    }
    It 'returns empty string for null input' {
        EscHtml $null | Should -Be ''
    }
    It 'returns empty string for empty input' {
        EscHtml '' | Should -Be ''
    }
    It 'passes through plain text unchanged' {
        EscHtml 'hello world' | Should -Be 'hello world'
    }
    It 'handles multiple special chars in one string' {
        EscHtml '<b>me & you</b>' | Should -Be '&lt;b&gt;me &amp; you&lt;/b&gt;'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Get-TKConfig
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Get-TKConfig' {
    BeforeAll {
        # Point the module's config path at a temp directory
        $script:TempDir = Join-Path $TestDrive 'TKConfig'
        New-Item -ItemType Directory -Path $script:TempDir -Force | Out-Null
    }

    Context 'when config.json does not exist' {
        BeforeAll {
            # Import module with a fresh PSScriptRoot pointing at TempDir (no config.json there)
            # We test via the exported function directly; the real config path uses $PSScriptRoot.
            # To isolate, we check that defaults have the expected shape.
            $cfg = Get-TKConfig
        }

        It 'returns an object' {
            $cfg | Should -Not -BeNullOrEmpty
        }
        It 'has OrgName property' {
            $cfg.PSObject.Properties.Name | Should -Contain 'OrgName'
        }
        It 'has LogDirectory property' {
            $cfg.PSObject.Properties.Name | Should -Contain 'LogDirectory'
        }
        It 'has TeamsWebhook property' {
            $cfg.PSObject.Properties.Name | Should -Contain 'TeamsWebhook'
        }
        It 'has Archive section' {
            $cfg.Archive | Should -Not -BeNullOrEmpty
        }
        It 'has Phantom section' {
            $cfg.Phantom | Should -Not -BeNullOrEmpty
        }
        It 'has Covenant section' {
            $cfg.Covenant | Should -Not -BeNullOrEmpty
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Set-TKConfig / Get-TKConfig round-trip
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Set-TKConfig and Get-TKConfig round-trip' {
    BeforeAll {
        # We cannot easily redirect $PSScriptRoot inside the module, so these
        # tests verify the JSON serialisation logic using a local temp file.
        $script:ConfigFile = Join-Path $TestDrive 'config.json'
    }

    It 'writes and reads a top-level key' {
        # Write a minimal config directly to simulate what Set-TKConfig would produce
        [PSCustomObject]@{ OrgName = 'Contoso' } | ConvertTo-Json | Set-Content $script:ConfigFile -Encoding UTF8
        $raw = Get-Content $script:ConfigFile | ConvertFrom-Json
        $raw.OrgName | Should -Be 'Contoso'
    }

    It 'produces valid JSON' {
        { Get-Content $script:ConfigFile | ConvertFrom-Json } | Should -Not -Throw
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Test-IsAdmin
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Test-IsAdmin' {
    It 'returns a boolean' {
        $result = Test-IsAdmin
        $result | Should -BeOfType [bool]
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Write-TKError — smoke tests (no real network/filesystem side effects in CI)
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Write-TKError' {
    It 'does not throw when LogDirectory is not configured' {
        { Write-TKError -ScriptName 'test' -Message 'unit test error' -Category 'Test' } |
            Should -Not -Throw
    }

    It 'does not throw with a blank Teams webhook' {
        { Write-TKError -ScriptName 'test' -Message 'webhook test' } |
            Should -Not -Throw
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Module exports
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Module exports' {
    # Pester 5 data binding: -ForEach hashtable keys become variables in the
    # test body. A plain `foreach { It { ... $fn } }` leaks the discovery-time
    # loop variable out of scope by run phase — use -ForEach so $fn resolves.
    It 'exports <fn>' -ForEach @(
        @{ fn = 'Write-Section' }
        @{ fn = 'Write-Step' }
        @{ fn = 'Write-Ok' }
        @{ fn = 'Write-Warn' }
        @{ fn = 'Write-Fail' }
        @{ fn = 'Write-Info' }
        @{ fn = 'EscHtml' }
        @{ fn = 'Get-TKHtmlCss' }
        @{ fn = 'Get-TKHtmlHead' }
        @{ fn = 'Get-TKHtmlFoot' }
        @{ fn = 'Test-IsAdmin' }
        @{ fn = 'Assert-AdminPrivilege' }
        @{ fn = 'Invoke-AdminElevation' }
        @{ fn = 'Get-TKConfig' }
        @{ fn = 'Set-TKConfig' }
        @{ fn = 'Resolve-LogDirectory' }
        @{ fn = 'Start-TKTranscript' }
        @{ fn = 'Stop-TKTranscript' }
        @{ fn = 'Write-TKError' }
    ) {
        Get-Command $fn -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Script syntax validation — all .ps1 files must parse without errors
# ─────────────────────────────────────────────────────────────────────────────
Describe 'PowerShell syntax — all scripts' {
    $scriptCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -File |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It '<Name> has no parse errors' -ForEach $scriptCases {
        $errors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile(
            $FullName, [ref]$null, [ref]$errors
        )
        $errors.Count | Should -Be 0
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Module bootstrap compliance — every tool script must use the shared-module
# bootstrap block so the module is auto-downloaded when missing and imports
# fail loudly (-ErrorAction Stop) rather than silently partially-executing.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Module bootstrap compliance — all tool scripts' {
    $scriptCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -File |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It '<Name> defines $TKModulePath next to $PSScriptRoot' -ForEach $scriptCases {
        $content = Get-Content $FullName -Raw
        $content | Should -Match '\$TKModulePath\s*=\s*Join-Path\s+\$PSScriptRoot\s+''TechnicianToolkit\.psm1'''
    }

    It '<Name> imports via $TKModulePath with -ErrorAction Stop' -ForEach $scriptCases {
        $content = Get-Content $FullName -Raw
        $content | Should -Match 'Import-Module\s+\$TKModulePath\s+-Force\s+-ErrorAction\s+Stop'
    }

    It '<Name> no longer uses the silent-fail import' -ForEach $scriptCases {
        $content = Get-Content $FullName -Raw
        # The old pattern (quoted path, no -ErrorAction) must be gone.
        $content | Should -Not -Match 'Import-Module\s+"\$PSScriptRoot\\TechnicianToolkit\.psm1"\s+-Force\s*$'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Param block compliance — interactive tool scripts must declare -Unattended
# Excludes the two launcher-style scripts that don't have sensible defaults
# for their required inputs (grimoire needs a tool choice; specter needs a
# target machine + credentials).
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Param block compliance — -Unattended switch' {
    $scriptCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -File |
        Where-Object { $_.Name -notin @('grimoire.ps1', 'specter.ps1') } |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It '<Name> declares -Unattended' -ForEach $scriptCases {
        $errors = $null
        $ast    = [System.Management.Automation.Language.Parser]::ParseFile(
            $FullName, [ref]$null, [ref]$errors
        )
        $params = $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.ParameterAst]
        }, $true)
        $paramNames = $params | ForEach-Object { $_.Name.VariablePath.UserPath }
        $paramNames | Should -Contain 'Unattended'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# GRIMOIRE registry integrity — every File entry must exist on disk
# ─────────────────────────────────────────────────────────────────────────────
Describe 'GRIMOIRE registry integrity' {
    BeforeAll {
        # Pester 5: BeforeAll-scoped variables are accessible to It bodies
        # without $script: prefix. Discovery-scope vars (below) are not.
        $grimoirePath = Join-Path $PSScriptRoot '..\grimoire.ps1'
    }

    It 'grimoire.ps1 exists' {
        $grimoirePath | Should -Exist
    }

    # -ForEach cases must be built at discovery time.
    $GrimoirePath = Join-Path $PSScriptRoot '..\grimoire.ps1'
    $ToolkitRoot  = Join-Path $PSScriptRoot '..'
    $registryContent = Get-Content $GrimoirePath -Raw
    $registryCases   = [regex]::Matches($registryContent, "File\s*=\s*'([^']+)'") |
        ForEach-Object { $_.Groups[1].Value } |
        Select-Object -Unique |
        ForEach-Object { @{ FileName = $_; FullPath = (Join-Path $ToolkitRoot $_) } }

    It "registered tool '<FileName>' exists on disk" -ForEach $registryCases {
        $FullPath | Should -Exist
    }
}
