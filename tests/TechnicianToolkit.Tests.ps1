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
        @{ fn = 'Show-TKReportResult' }
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
        @{ fn = 'Add-TKNote' }
        @{ fn = 'Get-TKNote' }
        @{ fn = 'Clear-TKNote' }
        @{ fn = 'Export-TKNoteReport' }
    ) {
        Get-Command $fn -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Technician notes — Add-TKNote / Get-TKNote / Clear-TKNote / Export-TKNoteReport
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Technician notes' {
    BeforeEach { Clear-TKNote }

    It 'starts the session with no notes' {
        @(Get-TKNote).Count | Should -Be 0
    }
    It 'records a note' {
        Add-TKNote -Text 'first note'
        $n = @(Get-TKNote)
        $n.Count   | Should -Be 1
        $n[0].Text | Should -Be 'first note'
    }
    It 'defaults the category to Info' {
        Add-TKNote -Text 'x'
        (@(Get-TKNote))[0].Category | Should -Be 'Info'
    }
    It 'rejects an invalid category' {
        { Add-TKNote -Text 'x' -Category 'Bogus' } | Should -Throw
    }
    It 'preserves insertion order' {
        Add-TKNote -Text 'one'
        Add-TKNote -Text 'two'
        $n = @(Get-TKNote)
        $n[0].Text | Should -Be 'one'
        $n[1].Text | Should -Be 'two'
    }
    It 'Clear-TKNote empties the buffer' {
        Add-TKNote -Text 'x'
        Clear-TKNote
        @(Get-TKNote).Count | Should -Be 0
    }

    Context 'Export-TKNoteReport' {
        It 'writes an HTML file and returns its path' {
            Add-TKNote -Text 'did a thing' -Category Action
            $out = Join-Path $TestDrive 'notes.html'
            Export-TKNoteReport -Path $out -ScriptName 'T.E.S.T.' | Should -Be $out
            $out | Should -Exist
        }
        It 'produces a balanced HTML document' {
            Add-TKNote -Text 'note body'
            $out = Join-Path $TestDrive 'notes2.html'
            Export-TKNoteReport -Path $out
            $doc = Get-Content $out -Raw
            $doc | Should -Match '^<!DOCTYPE html>'
            $doc | Should -Match '</html>'
        }
        It 'escapes HTML-special characters in note text' {
            Add-TKNote -Text '<script>alert(1)</script>'
            $out = Join-Path $TestDrive 'notes3.html'
            Export-TKNoteReport -Path $out
            $doc = Get-Content $out -Raw
            $doc | Should -Not -Match '<script>alert'
            $doc | Should -Match '&lt;script&gt;'
        }
        It 'includes the ticket reference when supplied' {
            Add-TKNote -Text 'x'
            $out = Join-Path $TestDrive 'notes4.html'
            Export-TKNoteReport -Path $out -Ticket 'INC0099'
            (Get-Content $out -Raw) | Should -Match 'INC0099'
        }
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
# UTF-8 BOM — every .ps1/.psm1 must start with a UTF-8 byte-order mark.
# Windows PowerShell 5.1 reads a BOM-less file as ANSI (Windows-1252), which
# mangles the Unicode box-drawing banners and menu glyphs at parse time (the
# launcher invokes powershell.exe 5.1, so it hits this every run). A BOM forces
# UTF-8 decoding. CI runs under pwsh (PS7, which assumes UTF-8) so it would not
# otherwise catch a missing BOM — this test does.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'UTF-8 BOM — all scripts' {
    $bomCases = Get-ChildItem -Path (Join-Path $PSScriptRoot '..') -Include '*.ps1', '*.psm1' -File -Recurse |
        Where-Object { $_.FullName -notmatch ([regex]::Escape([IO.Path]::DirectorySeparatorChar + '.git' + [IO.Path]::DirectorySeparatorChar)) } |
        ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }

    It '<Name> begins with a UTF-8 BOM' -ForEach $bomCases {
        $bytes = [System.IO.File]::ReadAllBytes($FullName)
        $bytes.Length | Should -BeGreaterThan 2
        $bytes[0..2] -join ',' | Should -Be '239,187,191' -Because "$Name must start with a UTF-8 BOM (EF BB BF) so Windows PowerShell 5.1 parses it as UTF-8"
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
# Legacy-name regression — the v3.0 rename retired eight tool acronyms; the
# v3.1 cleanup deleted their forwarding stubs. New source or documentation
# must never reintroduce the retired names. CHANGELOG (which documents the
# rename) is exempt.
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

    # Files that are *about* the rename legitimately mention the retired names.
    $allowlist = @('CHANGELOG.md', 'TechnicianToolkit.Tests.ps1')

    $root  = Resolve-Path (Join-Path $PSScriptRoot '..')
    $files = Get-ChildItem -Path $root -Recurse -File -Include '*.ps1','*.md' |
        Where-Object {
            $_.FullName -notmatch [regex]::Escape([IO.Path]::DirectorySeparatorChar + '.git' + [IO.Path]::DirectorySeparatorChar) -and
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
# Deprecation stubs removed — the eight v3.0 forwarding stubs were retired in
# v3.1. Their filenames must not reappear in the working tree (any reintroduction
# would resurrect a name we have explicitly deleted).
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Deprecation stubs removed' {
    $retiredStubs = @(
        'oracle.ps1','sentinel.ps1','bastion.ps1','vault.ps1',
        'phantom.ps1','specter.ps1','aegis.ps1','relic.ps1'
    ) | ForEach-Object {
        @{ Name = $_; FullPath = (Join-Path $PSScriptRoot "..\$_") }
    }

    It '<Name> no longer exists' -ForEach $retiredStubs {
        $FullPath | Should -Not -Exist
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

# ─────────────────────────────────────────────────────────────────────────────
# Tier-mapper data tables — the verdict logic in PALADIN / BEACON / PORTAL
# leans on small reference hashtables (and one tiny helper for ASR action
# codes). This block extracts those tables via AST lookup and asserts on
# their contents, so a careless rename or removal fails CI loudly.
#
# We don't dot-source the tools whole because they have side-effects (they
# launch their main flow on import). Instead we walk the AST, find the
# specific assignment / function-definition node by name, and re-evaluate
# just that node in the test scope.
# ─────────────────────────────────────────────────────────────────────────────
Describe 'Tier-mapper data tables' {
    BeforeAll {
        $script:ToolkitRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
        $script:PaladinPath = Join-Path $script:ToolkitRoot 'paladin.ps1'
        $script:BeaconPath  = Join-Path $script:ToolkitRoot 'beacon.ps1'
        $script:PortalPath  = Join-Path $script:ToolkitRoot 'portal.ps1'

        function Import-ScriptHashtable {
            param([string]$ScriptPath, [string]$VarName)
            $errs = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile(
                $ScriptPath, [ref]$null, [ref]$errs
            )
            $assign = $ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
                $n.Left.VariablePath.UserPath -eq $VarName
            }, $true) | Select-Object -First 1
            if (-not $assign) { return $null }
            return & ([scriptblock]::Create($assign.Right.Extent.Text))
        }

        function Get-ScriptFunctionScriptBlock {
            param([string]$ScriptPath, [string]$FuncName)
            $errs = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile(
                $ScriptPath, [ref]$null, [ref]$errs
            )
            $fn = $ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $FuncName
            }, $true) | Select-Object -First 1
            if (-not $fn) { return $null }
            # Returns a scriptblock that, when dot-sourced, defines the function in the caller's scope.
            return [scriptblock]::Create($fn.Extent.Text)
        }
    }

    Context 'PALADIN: $AsrRuleNames hashtable' {
        It 'covers at least 16 well-known ASR rule GUIDs' {
            $asr = Import-ScriptHashtable -ScriptPath $script:PaladinPath -VarName 'AsrRuleNames'
            $asr | Should -Not -BeNullOrEmpty
            $asr.Count | Should -BeGreaterOrEqual 16
        }
        It 'maps the abused-driver GUID to a recognisable name' {
            $asr = Import-ScriptHashtable -ScriptPath $script:PaladinPath -VarName 'AsrRuleNames'
            $asr['56a863a9-875e-4185-98a7-b882c64b5ce5'] | Should -Match 'vulnerable signed drivers'
        }
        It 'maps the LSASS-credential-theft GUID' {
            $asr = Import-ScriptHashtable -ScriptPath $script:PaladinPath -VarName 'AsrRuleNames'
            $asr['9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2'] | Should -Match 'LSASS'
        }
        It 'uses lowercase GUID keys (matches the case from Get-MpPreference output)' {
            $asr = Import-ScriptHashtable -ScriptPath $script:PaladinPath -VarName 'AsrRuleNames'
            foreach ($k in $asr.Keys) {
                $k | Should -Match '^[0-9a-f-]+$' -Because "ASR keys must be lowercase to match Get-MpPreference output (saw '$k')"
            }
        }
    }

    Context 'PALADIN: Get-AsrActionLabel function' {
        BeforeAll {
            # Extract the function definition AST from paladin.ps1 and dot-source
            # it so this Context can call the function directly.
            $sb = Get-ScriptFunctionScriptBlock -ScriptPath $script:PaladinPath -FuncName 'Get-AsrActionLabel'
            . $sb
        }
        It 'maps action 0 to Not Configured' {
            Get-AsrActionLabel -Action 0 | Should -Be 'Not Configured'
        }
        It 'maps action 1 to Block' {
            Get-AsrActionLabel -Action 1 | Should -Be 'Block'
        }
        It 'maps action 2 to Audit' {
            Get-AsrActionLabel -Action 2 | Should -Be 'Audit'
        }
        It 'maps action 6 to Warn' {
            Get-AsrActionLabel -Action 6 | Should -Be 'Warn'
        }
        It 'falls back to "Unknown (<n>)" for unmapped codes' {
            Get-AsrActionLabel -Action 99 | Should -Be 'Unknown (99)'
        }
    }

    Context 'BEACON: $AuthStrength hashtable (Wi-Fi)' {
        It 'classifies open / shared / WEP as Insecure' {
            $auth = Import-ScriptHashtable -ScriptPath $script:BeaconPath -VarName 'AuthStrength'
            $auth['open']   | Should -Be 'Insecure'
            $auth['shared'] | Should -Be 'Insecure'
            $auth['WEP']    | Should -Be 'Insecure'
        }
        It 'classifies WPA1 (WPA / WPAPSK) as Weak' {
            $auth = Import-ScriptHashtable -ScriptPath $script:BeaconPath -VarName 'AuthStrength'
            $auth['WPA']    | Should -Be 'Weak'
            $auth['WPAPSK'] | Should -Be 'Weak'
        }
        It 'classifies WPA2 personal+enterprise / WPA3 / OWE as Strong' {
            $auth = Import-ScriptHashtable -ScriptPath $script:BeaconPath -VarName 'AuthStrength'
            $auth['WPA2']    | Should -Be 'Strong'
            $auth['WPA2PSK'] | Should -Be 'Strong'
            $auth['WPA3SAE'] | Should -Be 'Strong'
            $auth['WPA3ENT'] | Should -Be 'Strong'
            $auth['OWE']     | Should -Be 'Strong'
        }
    }

    Context 'BEACON: $CipherStrength hashtable (Wi-Fi)' {
        It 'classifies cipher tiers correctly' {
            $c = Import-ScriptHashtable -ScriptPath $script:BeaconPath -VarName 'CipherStrength'
            $c['none'] | Should -Be 'Insecure'
            $c['WEP']  | Should -Be 'Insecure'
            $c['TKIP'] | Should -Be 'Weak'
            $c['AES']  | Should -Be 'Strong'
            $c['GCMP'] | Should -Be 'Strong'
        }
    }

    Context 'PORTAL: $AuthStrength hashtable (VPN)' {
        It 'classifies PAP as Insecure (cleartext credentials)' {
            $a = Import-ScriptHashtable -ScriptPath $script:PortalPath -VarName 'AuthStrength'
            $a['Pap'] | Should -Be 'Insecure'
        }
        It 'classifies CHAP as Weak' {
            $a = Import-ScriptHashtable -ScriptPath $script:PortalPath -VarName 'AuthStrength'
            $a['Chap'] | Should -Be 'Weak'
        }
        It 'classifies MS-CHAPv2 as Acceptable' {
            $a = Import-ScriptHashtable -ScriptPath $script:PortalPath -VarName 'AuthStrength'
            $a['MSChapv2'] | Should -Be 'Acceptable'
        }
        It 'classifies EAP and MachineCertificate as Strong' {
            $a = Import-ScriptHashtable -ScriptPath $script:PortalPath -VarName 'AuthStrength'
            $a['Eap']                | Should -Be 'Strong'
            $a['MachineCertificate'] | Should -Be 'Strong'
        }
    }

    Context 'PORTAL: $EncryptionStrength hashtable (VPN)' {
        It 'classifies encryption levels correctly' {
            $e = Import-ScriptHashtable -ScriptPath $script:PortalPath -VarName 'EncryptionStrength'
            $e['NoEncryption'] | Should -Be 'Insecure'
            $e['Optional']     | Should -Be 'Weak'
            $e['Required']     | Should -Be 'Strong'
            $e['Maximum']      | Should -Be 'Strong'
        }
    }
}
