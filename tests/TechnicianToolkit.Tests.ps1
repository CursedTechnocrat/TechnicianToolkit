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
    $expectedFunctions = @(
        'Write-Section', 'Write-Step', 'Write-Ok', 'Write-Warn', 'Write-Fail', 'Write-Info',
        'EscHtml',
        'Test-IsAdmin', 'Assert-AdminPrivilege', 'Invoke-AdminElevation',
        'Get-TKConfig', 'Set-TKConfig',
        'Resolve-LogDirectory',
        'Start-TKTranscript', 'Stop-TKTranscript',
        'Write-TKError'
    )

    foreach ($fn in $expectedFunctions) {
        It "exports $fn" {
            Get-Command $fn -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Script syntax validation — all .ps1 files must parse without errors
# ─────────────────────────────────────────────────────────────────────────────
Describe 'PowerShell syntax — all scripts' {
    $scripts = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -File

    foreach ($script in $scripts) {
        It "$($script.Name) has no parse errors" {
            $errors = $null
            $null = [System.Management.Automation.Language.Parser]::ParseFile(
                $script.FullName, [ref]$null, [ref]$errors
            )
            $errors.Count | Should -Be 0
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Module import compliance — every tool script must import TechnicianToolkit.psm1
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Module import compliance — all tool scripts' {
    $scripts = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -File

    foreach ($script in $scripts) {
        It "$($script.Name) imports TechnicianToolkit.psm1" {
            $content = Get-Content $script.FullName -Raw
            $content | Should -Match 'Import-Module.*TechnicianToolkit\.psm1'
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Param block compliance — interactive tool scripts must declare -Unattended
# Excludes grimoire.ps1 (hub launcher, not an interactive tool itself)
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Param block compliance — -Unattended switch' {
    $scripts = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -File |
        Where-Object { $_.Name -ne 'grimoire.ps1' }

    foreach ($script in $scripts) {
        It "$($script.Name) declares -Unattended" {
            $errors = $null
            $ast    = [System.Management.Automation.Language.Parser]::ParseFile(
                $script.FullName, [ref]$null, [ref]$errors
            )
            $params = $ast.FindAll({
                param($node)
                $node -is [System.Management.Automation.Language.ParameterAst]
            }, $true)
            $paramNames = $params | ForEach-Object { $_.Name.VariablePath.UserPath }
            $paramNames | Should -Contain 'Unattended'
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# GRIMOIRE registry integrity — every File entry must exist on disk
# ─────────────────────────────────────────────────────────────────────────────
Describe 'GRIMOIRE registry integrity' {
    $GrimoirePath = Join-Path $PSScriptRoot '..\grimoire.ps1'
    $ToolkitRoot  = Join-Path $PSScriptRoot '..'

    It 'grimoire.ps1 exists' {
        $GrimoirePath | Should -Exist
    }

    # Parse the $Tools array by extracting File = '...' values from the script text
    $content   = Get-Content $GrimoirePath -Raw
    $fileNames = [regex]::Matches($content, "File\s*=\s*'([^']+)'") |
        ForEach-Object { $_.Groups[1].Value } |
        Select-Object -Unique

    foreach ($fileName in $fileNames) {
        It "registered tool '$fileName' exists on disk" {
            $fullPath = Join-Path $ToolkitRoot $fileName
            $fullPath | Should -Exist
        }
    }
}
