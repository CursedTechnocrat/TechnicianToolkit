// ToolParametersTests.cs - The param() block reader that generates the run form.
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

using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace TechnicianToolkit.Engine.Tests
{
    public sealed class ToolParametersTests
    {
        private static IReadOnlyList<ToolParameter> Read(ScriptFixture fixture, string source)
            => ToolParameters.Load(fixture.Write("tool.ps1", source));

        private static ToolParameter Single(ScriptFixture fixture, string source, string name)
            => Read(fixture, source).First(p => p.Name == name);

        [Fact]
        public void A_switch_becomes_a_checkbox()
        {
            using var fixture = new ScriptFixture();

            ToolParameter p = Single(fixture, "param([switch]$Unattended)", "Unattended");

            Assert.True(p.IsSwitch);
            Assert.Equal("switch", p.TypeName);
            Assert.Equal(ParameterControl.Checkbox, p.Control);
        }

        [Fact]
        public void A_securestring_becomes_a_masked_field()
        {
            using var fixture = new ScriptFixture();

            ToolParameter p = Single(fixture, "param([securestring]$Password)", "Password");

            Assert.True(p.IsSecureString);
            Assert.Equal(ParameterControl.MaskedText, p.Control);
        }

        [Fact]
        public void A_ValidateSet_becomes_a_dropdown_carrying_its_values()
        {
            using var fixture = new ScriptFixture();

            ToolParameter p = Single(
                fixture,
                "param([ValidateSet('Fast','Full','Deep')][string]$Mode)",
                "Mode");

            Assert.Equal(new[] { "Fast", "Full", "Deep" }, p.ValidValues);
            Assert.Equal(ParameterControl.Dropdown, p.Control);
        }

        [Fact]
        public void A_ValidateScript_becomes_a_path_picker()
        {
            using var fixture = new ScriptFixture();

            ToolParameter p = Single(
                fixture,
                "param([ValidateScript({ Test-Path $_ })][string]$Destination)",
                "Destination");

            Assert.True(p.HasValidateScript);
            Assert.Equal(ParameterControl.PathPicker, p.Control);
        }

        [Fact]
        public void A_ValidatePattern_is_carried_through_for_the_form_to_enforce()
        {
            using var fixture = new ScriptFixture();

            ToolParameter p = Single(
                fixture,
                @"param([ValidatePattern('^\d{1,3}(\.\d{1,3}){3}$')][string]$IPAddress)",
                "IPAddress");

            Assert.Equal(@"^\d{1,3}(\.\d{1,3}){3}$", p.ValidationPattern);
            Assert.Equal(ParameterControl.Text, p.Control);
        }

        [Fact]
        public void A_plain_string_is_a_text_field()
        {
            using var fixture = new ScriptFixture();

            Assert.Equal(ParameterControl.Text, Single(fixture, "param([string]$Name)", "Name").Control);
        }

        // Mandatory has three spellings in the wild and all of them mean the same
        // thing. A bare 'Mandatory' carries an implicit true that the AST models
        // as an omitted expression, which is the one easy to miss.
        [Theory]
        [InlineData("param([Parameter(Mandatory)][string]$Target)")]
        [InlineData("param([Parameter(Mandatory=$true)][string]$Target)")]
        [InlineData("param([Parameter(Mandatory = $True)][string]$Target)")]
        public void Mandatory_is_recognised_however_it_is_written(string source)
        {
            using var fixture = new ScriptFixture();

            Assert.True(Single(fixture, source, "Target").IsMandatory);
        }

        [Theory]
        [InlineData("param([Parameter(Mandatory=$false)][string]$Target)")]
        [InlineData("param([Parameter(Position=0)][string]$Target)")]
        [InlineData("param([string]$Target)")]
        public void Optional_parameters_are_not_reported_as_mandatory(string source)
        {
            using var fixture = new ScriptFixture();

            Assert.False(Single(fixture, source, "Target").IsMandatory);
        }

        [Fact]
        public void A_default_value_is_carried_through_as_written()
        {
            using var fixture = new ScriptFixture();

            Assert.Equal("'C:\\Reports'", Single(fixture, @"param([string]$Path = 'C:\Reports')", "Path").DefaultValue);
        }

        [Fact]
        public void A_parameter_without_a_default_reports_none()
        {
            using var fixture = new ScriptFixture();

            Assert.Null(Single(fixture, "param([string]$Path)", "Path").DefaultValue);
        }

        [Fact]
        public void A_script_with_no_param_block_yields_no_parameters()
        {
            using var fixture = new ScriptFixture();

            Assert.Empty(Read(fixture, "Write-Host 'no parameters here'"));
        }

        // Only the script's own param() drives the form. A nested function's
        // parameters are internal detail and would be nonsense to prompt for.
        [Fact]
        public void Nested_function_parameters_are_not_mistaken_for_the_scripts_own()
        {
            using var fixture = new ScriptFixture();
            const string source = @"
param([switch]$Unattended)

function Get-Thing {
    param([string]$InternalOnly, [switch]$AlsoInternal)
    Write-Host $InternalOnly
}
";

            IReadOnlyList<ToolParameter> parameters = Read(fixture, source);

            Assert.Single(parameters);
            Assert.Equal("Unattended", parameters[0].Name);
        }

        [Fact]
        public void Every_parameter_is_returned_in_declaration_order()
        {
            using var fixture = new ScriptFixture();
            const string source = @"
param(
    [string[]]$DirectUrl,
    [switch]$Unattended,
    [switch]$Transcript,
    [switch]$WhatIf
)
";

            IReadOnlyList<ToolParameter> parameters = Read(fixture, source);

            Assert.Equal(
                new[] { "DirectUrl", "Unattended", "Transcript", "WhatIf" },
                parameters.Select(p => p.Name));
        }
    }
}
