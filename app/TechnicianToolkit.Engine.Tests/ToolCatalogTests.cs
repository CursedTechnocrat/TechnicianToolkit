// ToolCatalogTests.cs - The GRIMOIRE registry reader.
// Part of the Technician Toolkit - https://github.com/CursedTechnocrat/TechnicianToolkit
//
// Copyright (C) 2026 John Joseph Bejarana (CursedTechnocrat) and the Technician Toolkit contributors
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.
//
// SPDX-License-Identifier: GPL-3.0-or-later

using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace TechnicianToolkit.Engine.Tests
{
    public sealed class ToolCatalogTests
    {
        private const string Registry = @"
$CategoryOrder = @(
    'Diagnostics & Reporting'
    'Deployment & Onboarding'
)

$Tools = @(
    [PSCustomObject]@{
        Key         = '1'
        Name        = 'W.A.R.D.'
        File        = 'ward.ps1'
        Version     = '5.0'
        Description = 'Local account roster'
        Color       = 'Cyan'
        Category    = 'Diagnostics & Reporting'
    },
    [PSCustomObject]@{
        Key         = '2'
        Name        = 'C.O.N.J.U.R.E.'
        File        = 'conjure.ps1'
        Version     = '5.0'
        Description = 'Software deployment'
        Color       = 'Blue'
        Category    = 'Deployment & Onboarding'
    }
)
";

        private static IReadOnlyList<ToolEntry> LoadRegistry(ScriptFixture fixture, string source = Registry)
            => ToolCatalog.Load(fixture.Write("grimoire.ps1", source));

        [Fact]
        public void Load_reads_every_entry_in_declaration_order()
        {
            using var fixture = new ScriptFixture();

            IReadOnlyList<ToolEntry> catalog = LoadRegistry(fixture);

            Assert.Equal(2, catalog.Count);
            Assert.Equal("ward.ps1", catalog[0].File);
            Assert.Equal("conjure.ps1", catalog[1].File);
        }

        [Fact]
        public void Load_reads_every_field_of_an_entry()
        {
            using var fixture = new ScriptFixture();

            ToolEntry ward = LoadRegistry(fixture)[0];

            Assert.Equal("1", ward.Key);
            Assert.Equal("W.A.R.D.", ward.Name);
            Assert.Equal("5.0", ward.Version);
            Assert.Equal("Local account roster", ward.Description);
            Assert.Equal("Cyan", ward.Color);
            Assert.Equal("Diagnostics & Reporting", ward.Category);
        }

        [Fact]
        public void ShortName_strips_the_acronym_dots()
        {
            using var fixture = new ScriptFixture();

            Assert.Equal("WARD", LoadRegistry(fixture)[0].ShortName);
            Assert.Equal("CONJURE", LoadRegistry(fixture)[1].ShortName);
        }

        // The sweep collects every hashtable under the assignment, so anything
        // without a File key has to be dropped or an unrelated literal in the
        // registry block becomes a phantom tool.
        [Fact]
        public void Load_ignores_hashtables_that_are_not_tool_entries()
        {
            using var fixture = new ScriptFixture();
            const string withExtras = @"
$Tools = @(
    [PSCustomObject]@{ Key = '1'; Name = 'W.A.R.D.'; File = 'ward.ps1' },
    @{ SomethingElse = 'not a tool'; Color = 'Red' }
)
";

            IReadOnlyList<ToolEntry> catalog = LoadRegistry(fixture, withExtras);

            Assert.Single(catalog);
            Assert.Equal("ward.ps1", catalog[0].File);
        }

        [Theory]
        [InlineData("1", "ward.ps1")]              // registry key
        [InlineData("WARD", "ward.ps1")]           // short name
        [InlineData("ward", "ward.ps1")]           // short name, wrong case
        [InlineData("W.A.R.D.", "ward.ps1")]       // dotted acronym
        [InlineData("ward.ps1", "ward.ps1")]       // file name
        [InlineData("WARD.PS1", "ward.ps1")]       // file name, wrong case
        [InlineData("conjure", "conjure.ps1")]     // stem only
        public void Resolve_accepts_every_form_a_technician_might_type(string token, string expectedFile)
        {
            using var fixture = new ScriptFixture();

            ToolEntry? found = ToolCatalog.Resolve(LoadRegistry(fixture), token);

            Assert.NotNull(found);
            Assert.Equal(expectedFile, found!.File);
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("nosuchtool")]
        [InlineData("99")]
        public void Resolve_returns_null_rather_than_guessing(string token)
        {
            using var fixture = new ScriptFixture();

            Assert.Null(ToolCatalog.Resolve(LoadRegistry(fixture), token));
        }

        [Fact]
        public void LoadCategoryOrder_returns_the_declared_order()
        {
            using var fixture = new ScriptFixture();

            IReadOnlyList<string> order = ToolCatalog.LoadCategoryOrder(fixture.Write("grimoire.ps1", Registry));

            Assert.Equal(new[] { "Diagnostics & Reporting", "Deployment & Onboarding" }, order);
        }

        [Fact]
        public void LoadCategoryOrder_is_empty_when_the_hub_declares_none()
        {
            using var fixture = new ScriptFixture();

            IReadOnlyList<string> order = ToolCatalog.LoadCategoryOrder(
                fixture.Write("grimoire.ps1", "$Tools = @()"));

            Assert.Empty(order);
        }

        [Fact]
        public void Load_throws_when_the_hub_is_missing()
        {
            using var fixture = new ScriptFixture();

            Assert.Throws<FileNotFoundException>(
                () => ToolCatalog.Load(Path.Combine(fixture.Directory, "absent.ps1")));
        }

        // A renamed or restructured registry must fail loudly. Returning an empty
        // catalog would present the operator with a hub that simply has no tools.
        [Fact]
        public void Load_throws_when_the_registry_shape_has_changed()
        {
            using var fixture = new ScriptFixture();

            InvalidOperationException error = Assert.Throws<InvalidOperationException>(
                () => LoadRegistry(fixture, "$SomethingElse = @( 1, 2, 3 )"));

            Assert.Contains("$Tools", error.Message, StringComparison.Ordinal);
        }

        // A tool that will not parse would not run either, so the reader surfaces
        // the parse error instead of silently skipping the file.
        [Fact]
        public void Load_reports_a_parse_error_with_its_line_number()
        {
            using var fixture = new ScriptFixture();

            InvalidOperationException error = Assert.Throws<InvalidOperationException>(
                () => LoadRegistry(fixture, "$Tools = @(\n    [PSCustomObject]@{ File = 'x.ps1'\n"));

            Assert.Contains("grimoire.ps1", error.Message, StringComparison.Ordinal);
            Assert.Contains("line", error.Message, StringComparison.OrdinalIgnoreCase);
        }
    }
}
