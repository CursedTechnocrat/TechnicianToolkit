// ToolTraitsTests.cs - What the front end needs to know before it runs a tool.
// Part of the Technician Toolkit - https://github.com/CursedTechnocrat/TechnicianToolkit
//
// Copyright (C) 2026 CursedTechnocrat and the Technician Toolkit contributors
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

using System.IO;
using Xunit;

namespace TechnicianToolkit.Engine.Tests
{
    public sealed class ToolTraitsTests
    {
        private static ToolTrait Inspect(ScriptFixture fixture, string source)
            => ToolTraits.Inspect(fixture.Write("tool.ps1", source));

        // Both admin gates count. Invoke-AdminElevation relaunches elevated and
        // Assert-AdminPrivilege error-exits, but either one means the tool cannot
        // do its job without Administrator.
        [Theory]
        [InlineData("Invoke-AdminElevation -ScriptFile $PSCommandPath")]
        [InlineData("Assert-AdminPrivilege")]
        public void A_tool_with_an_admin_gate_is_reported_as_needing_one(string gate)
        {
            using var fixture = new ScriptFixture();

            Assert.True(Inspect(fixture, "param()\n" + gate).RequiresAdmin);
        }

        [Fact]
        public void A_tool_with_no_admin_gate_is_not_reported_as_needing_one()
        {
            using var fixture = new ScriptFixture();

            Assert.False(Inspect(fixture, "param()\nWrite-Host 'read-only'").RequiresAdmin);
        }

        // -WhatIf is the toolkit's marker for a tool that changes the machine;
        // the Pester suite enforces the same list from the other side.
        [Fact]
        public void Declaring_WhatIf_marks_a_tool_destructive()
        {
            using var fixture = new ScriptFixture();

            Assert.True(Inspect(fixture, "param([switch]$Unattended, [switch]$WhatIf)").IsDestructive);
        }

        [Fact]
        public void A_tool_without_WhatIf_is_not_marked_destructive()
        {
            using var fixture = new ScriptFixture();

            Assert.False(Inspect(fixture, "param([switch]$Unattended)").IsDestructive);
        }

        [Fact]
        public void Unattended_support_is_detected()
        {
            using var fixture = new ScriptFixture();

            Assert.True(Inspect(fixture, "param([switch]$Unattended)").SupportsUnattended);
            Assert.False(Inspect(fixture, "param([switch]$Transcript)").SupportsUnattended);
        }

        [Fact]
        public void Traits_are_read_case_insensitively_from_the_param_block()
        {
            using var fixture = new ScriptFixture();

            ToolTrait trait = Inspect(fixture, "param([switch]$unattended, [switch]$whatif)");

            Assert.True(trait.SupportsUnattended);
            Assert.True(trait.IsDestructive);
        }

        [Fact]
        public void Inspecting_a_missing_tool_throws()
        {
            using var fixture = new ScriptFixture();

            Assert.Throws<FileNotFoundException>(
                () => ToolTraits.Inspect(Path.Combine(fixture.Directory, "absent.ps1")));
        }
    }
}
