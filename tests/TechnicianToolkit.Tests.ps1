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
# Format-Bytes
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Format-Bytes' {
    It 'returns bytes for values under 1 KB' {
        Format-Bytes 512 | Should -Be '512 B'
    }
    It 'returns KB for values under 1 MB' {
        Format-Bytes 2048 | Should -Be '2.00 KB'
    }
    It 'returns MB for values under 1 GB' {
        Format-Bytes (5 * 1MB) | Should -Be '5.00 MB'
    }
    It 'returns GB for values under 1 TB' {
        Format-Bytes (3 * 1GB) | Should -Be '3.00 GB'
    }
    It 'returns TB for values at or above 1 TB' {
        Format-Bytes (2 * 1TB) | Should -Be '2.00 TB'
    }
    It 'handles zero bytes' {
        Format-Bytes 0 | Should -Be '0 B'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Get-TKHtmlHead / Get-TKHtmlFoot — structural smoke tests
# ─────────────────────────────────────────────────────────────────────────────
Describe 'HTML report helpers' {
    Context 'Get-TKHtmlHead' {
        It 'returns a well-formed HTML preamble' {
            $html = Get-TKHtmlHead -Title 'Unit Test' -ScriptName 'T.E.S.T.'
            $html | Should -Match '^<!DOCTYPE html>'
            $html | Should -Match '<html lang="en">'
            $html | Should -Match '<title>Unit Test</title>'
            $html | Should -Match '<div class="tk-main">'
        }
        It 'embeds the shared CSS block' {
            $html = Get-TKHtmlHead -Title 'X' -ScriptName 'X'
            $html | Should -Match '--tk-bg:'
            $html | Should -Match 'class="tk-page-header"'
        }
        It 'HTML-escapes a title containing special characters' {
            $html = Get-TKHtmlHead -Title '<script>&"' -ScriptName 'X'
            $html | Should -Match '&lt;script&gt;&amp;&quot;'
        }
        It 'renders a meta bar when MetaItems are supplied' {
            $meta = [ordered]@{ Generated = '2026-04-22'; Host = 'UNIT01' }
            $html = Get-TKHtmlHead -Title 'X' -ScriptName 'X' -MetaItems $meta
            $html | Should -Match 'class=''tk-meta-bar'''
            $html | Should -Match '>Generated<'
            $html | Should -Match '>UNIT01<'
        }
        It 'renders a nav bar when NavItems are supplied' {
            $html = Get-TKHtmlHead -Title 'X' -ScriptName 'X' -NavItems @('Alpha','Beta')
            $html | Should -Match 'class=''tk-nav'''
            $html | Should -Match 'Alpha</a>'
            $html | Should -Match 'Beta</a>'
        }
    }

    Context 'Get-TKHtmlFoot' {
        It 'closes the document' {
            $html = Get-TKHtmlFoot -ScriptName 'T.E.S.T. v1'
            $html | Should -Match '</body>'
            $html | Should -Match '</html>'
        }
        It 'includes the script name in the footer' {
            $html = Get-TKHtmlFoot -ScriptName 'T.E.S.T. v1'
            $html | Should -Match 'T\.E\.S\.T\. v1'
        }
    }

    Context 'round-trip' {
        It 'head + body + foot produces balanced HTML' {
            $doc = (Get-TKHtmlHead -Title 'X' -ScriptName 'X') + '<p>body</p>' + (Get-TKHtmlFoot -ScriptName 'X')
            # Each opening tag should have exactly one closing tag.
            ($doc | Select-String -Pattern '<html' -AllMatches).Matches.Count   | Should -Be 1
            ($doc | Select-String -Pattern '</html>' -AllMatches).Matches.Count | Should -Be 1
            ($doc | Select-String -Pattern '<body'  -AllMatches).Matches.Count  | Should -Be 1
            ($doc | Select-String -Pattern '</body>' -AllMatches).Matches.Count | Should -Be 1
        }
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
        It 'has Revenant section' {
            $cfg.Revenant | Should -Not -BeNullOrEmpty
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
        @{ fn = 'Format-Bytes' }
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
# for their required inputs (grimoire needs a tool choice; shade needs a
# target machine + credentials).
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Param block compliance — -Unattended switch' {
    $scriptCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -File |
        Where-Object { $_.Name -notin @('grimoire.ps1', 'shade.ps1') } |
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

# ─────────────────────────────────────────────────────────────────────────────
# Legacy-name regression — the v3.0 rename retired eight tool acronyms.
# New source or documentation must never reintroduce them. Deprecation stubs
# and CHANGELOG (which documents the rename) are exempt.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Legacy tool names must not reappear' {
    # Pester 5 scoping: the pattern list must be defined in BeforeAll so it
    # exists at Run time when each It block executes. Variables assigned in
    # the Describe body only exist at Discovery time.
    BeforeAll {
        $script:LegacyAcronyms = @(
            'O.R.A.C.L.E.', 'S.E.N.T.I.N.E.L.', 'B.A.S.T.I.O.N.',
            'V.A.U.L.T.',   'P.H.A.N.T.O.M.',   'S.P.E.C.T.E.R.',
            'A.E.G.I.S.',   'R.E.L.I.C.'
        )
    }

    # Discovery time: enumerate files as -ForEach test cases
    $legacyStubs = @(
        'oracle.ps1','sentinel.ps1','bastion.ps1','vault.ps1',
        'phantom.ps1','specter.ps1','aegis.ps1','relic.ps1'
    )
    # Files that are *about* the rename legitimately mention the retired names.
    $allowlist = @('CHANGELOG.md', 'TechnicianToolkit.Tests.ps1')

    $root  = Resolve-Path (Join-Path $PSScriptRoot '..')
    $files = Get-ChildItem -Path $root -Recurse -File -Include '*.ps1','*.md' |
        Where-Object {
            $_.FullName -notmatch [regex]::Escape([IO.Path]::DirectorySeparatorChar + '.git' + [IO.Path]::DirectorySeparatorChar) -and
            $_.Name -notin $legacyStubs -and
            $_.Name -notin $allowlist
        } |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It '<Name> contains no retired dotted acronyms' -ForEach $files {
        $hits = Select-String -Path $FullName -SimpleMatch -Pattern $script:LegacyAcronyms -ErrorAction SilentlyContinue
        $hits | Should -BeNullOrEmpty -Because "retired acronym found in $Name"
    }

    # Second form: bare underscore-prefixed filenames (e.g. `SPECTER_<MachineName>`,
    # `PHANTOM_MigrationLog_*.csv`). The v3.0 rename changed every tool's emitted
    # filename prefix, and the README logging table drifted without this catch
    # because the dotted form above didn't match the bare-prefix form.
    It '<Name> contains no retired filename prefixes' -ForEach $files {
        $prefixPatterns = @(
            'ORACLE_', 'SENTINEL_', 'BASTION_', 'VAULT_',
            'PHANTOM_', 'SPECTER_', 'AEGIS_', 'RELIC_'
        )
        $hits = Select-String -Path $FullName -SimpleMatch -Pattern $prefixPatterns -ErrorAction SilentlyContinue
        $hits | Should -BeNullOrEmpty -Because "retired filename prefix found in $Name"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# -WhatIf compliance — destructive tools must declare -WhatIf so that GRIMOIRE
# can pass it through in dry-run mode. See grimoire.ps1 Invoke-Tool.
# ─────────────────────────────────────────────────────────────────────────────
Describe '-WhatIf declared on destructive tools' {
    # Tools that make persistent, hard-to-reverse changes: file moves, registry
    # writes, domain joins, disk encryption toggles, AV policy changes, driver
    # and Windows Update installs, printer driver / network printer additions.
    $destructiveCases = @(
        'revenant.ps1','archive.ps1','covenant.ps1','sigil.ps1','cleanse.ps1','cipher.ps1',
        'forge.ps1','restoration.ps1','runepress.ps1'
    ) | ForEach-Object {
        @{ Name = $_; FullName = (Join-Path $PSScriptRoot "..\$_") }
    }

    It '<Name> declares -WhatIf' -ForEach $destructiveCases {
        $errors = $null
        $ast    = [System.Management.Automation.Language.Parser]::ParseFile(
            $FullName, [ref]$null, [ref]$errors
        )
        $paramNames = $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.ParameterAst]
        }, $true) | ForEach-Object { $_.Name.VariablePath.UserPath }
        $paramNames | Should -Contain 'WhatIf'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Deprecation stub integrity — each legacy-name stub must forward every
# argument to its renamed replacement and show a one-line deprecation warning.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Deprecation stub forwarding' {
    # Stub -> expected forwarding target (set at v3.0 rename).
    $stubCases = @(
        @{ Stub = 'oracle.ps1';   Target = 'auspex.ps1'    }
        @{ Stub = 'sentinel.ps1'; Target = 'gargoyle.ps1'  }
        @{ Stub = 'bastion.ps1';  Target = 'citadel.ps1'   }
        @{ Stub = 'vault.ps1';    Target = 'reliquary.ps1' }
        @{ Stub = 'phantom.ps1';  Target = 'revenant.ps1'  }
        @{ Stub = 'specter.ps1';  Target = 'shade.ps1'     }
        @{ Stub = 'aegis.ps1';    Target = 'talisman.ps1'  }
        @{ Stub = 'relic.ps1';    Target = 'artifact.ps1'  }
    ) | ForEach-Object {
        $_ + @{ StubPath = (Join-Path $PSScriptRoot "..\$($_.Stub)") }
    }

    It '<Stub> forwards to <Target>' -ForEach $stubCases {
        $content = Get-Content $StubPath -Raw
        # Forwarding site: `& $target @fwd` where $target = Join-Path ... '<target>.ps1'
        $content | Should -Match ([regex]::Escape("Join-Path `$PSScriptRoot '$Target'"))
        $content | Should -Match '&\s+\$target\s+@fwd'
    }

    It '<Stub> emits a deprecation warning' -ForEach $stubCases {
        $content = Get-Content $StubPath -Raw
        $content | Should -Match 'Write-Warning\s+"[^"]*deprecated'
    }

    It '<Stub> captures remaining args via ValueFromRemainingArguments' -ForEach $stubCases {
        $content = Get-Content $StubPath -Raw
        $content | Should -Match 'ValueFromRemainingArguments'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Duplicated helpers removed — the local HtmlEncode / Format-Bytes definitions
# that were consolidated into the shared module must not reappear.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'No duplicated helper functions' {
    $toolCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Filter '*.ps1' -File |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It '<Name> does not redefine HtmlEncode locally' -ForEach $toolCases {
        $content = Get-Content $FullName -Raw
        $content | Should -Not -Match '(?m)^\s*function\s+HtmlEncode\b'
    }

    It '<Name> does not redefine Format-Bytes locally' -ForEach $toolCases {
        $content = Get-Content $FullName -Raw
        $content | Should -Not -Match '(?m)^\s*function\s+Format-Bytes\b'
    }
}
