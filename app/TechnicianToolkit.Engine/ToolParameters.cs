// ToolParameters.cs - Reads a tool param() block by AST so a form can build itself.
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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation.Language;

namespace TechnicianToolkit.Engine
{
    /// <summary>
    /// One declared parameter, described richly enough that a form control can
    /// be chosen without any per-tool knowledge.
    /// </summary>
    public sealed class ToolParameter
    {
        public string Name { get; init; } = string.Empty;

        /// <summary>The declared type name, e.g. string, switch, securestring.</summary>
        public string TypeName { get; init; } = "object";

        public bool IsSwitch { get; init; }
        public bool IsSecureString { get; init; }
        public bool IsMandatory { get; init; }

        /// <summary>Values from [ValidateSet(...)]; empty when none was declared.</summary>
        public IReadOnlyList<string> ValidValues { get; init; } = Array.Empty<string>();

        /// <summary>The regex from [ValidatePattern(...)], or null.</summary>
        public string? ValidationPattern { get; init; }

        /// <summary>
        /// True when the parameter carries a [ValidateScript(...)]. In this suite
        /// that reliably means a path check, which is the cue for a browse button.
        /// </summary>
        public bool HasValidateScript { get; init; }

        /// <summary>The literal default as written in the script, or null.</summary>
        public string? DefaultValue { get; init; }

        /// <summary>
        /// The control a form should render. Kept as a computed property so the
        /// mapping lives in one place.
        /// </summary>
        public ParameterControl Control =>
            IsSwitch ? ParameterControl.Checkbox
            : IsSecureString ? ParameterControl.MaskedText
            : ValidValues.Count > 0 ? ParameterControl.Dropdown
            : HasValidateScript ? ParameterControl.PathPicker
            : ParameterControl.Text;

        public override string ToString() => Name + " [" + TypeName + "]";
    }

    public enum ParameterControl
    {
        Text,
        MaskedText,
        Checkbox,
        Dropdown,
        PathPicker,
    }

    /// <summary>
    /// The tools are already richly typed -- ValidateSet, ValidatePattern,
    /// ValidateScript, switch, securestring. Parsing the param() block yields a
    /// form for free, and one that can never drift from the script it describes.
    /// </summary>
    public static class ToolParameters
    {
        /// <summary>
        /// Read the top-level param() block of a script file.
        /// </summary>
        public static IReadOnlyList<ToolParameter> Load(string scriptPath)
        {
            Ast root = ToolCatalog.ParseFile(scriptPath);

            // The top-level param() block, not one belonging to a nested function.
            ParamBlockAst? paramBlock = (root as ScriptBlockAst)?.ParamBlock;
            if (paramBlock == null)
            {
                return Array.Empty<ToolParameter>();
            }

            return paramBlock.Parameters.Select(Read).ToList();
        }

        private static ToolParameter Read(ParameterAst parameter)
        {
            string typeName = parameter.StaticType?.Name ?? "object";
            bool isSwitch = string.Equals(typeName, "SwitchParameter", StringComparison.OrdinalIgnoreCase);
            bool isSecure = string.Equals(typeName, "SecureString", StringComparison.OrdinalIgnoreCase);

            var validValues = new List<string>();
            string? pattern = null;
            bool hasValidateScript = false;
            bool mandatory = false;

            foreach (AttributeBaseAst attributeBase in parameter.Attributes)
            {
                if (attributeBase is not AttributeAst attribute)
                {
                    continue;
                }

                string attributeName = attribute.TypeName.Name;

                if (attributeName.Equals("ValidateSet", StringComparison.OrdinalIgnoreCase))
                {
                    validValues.AddRange(attribute.PositionalArguments
                        .OfType<StringConstantExpressionAst>()
                        .Select(s => s.Value));
                }
                else if (attributeName.Equals("ValidatePattern", StringComparison.OrdinalIgnoreCase))
                {
                    pattern = attribute.PositionalArguments
                        .OfType<StringConstantExpressionAst>()
                        .Select(s => s.Value)
                        .FirstOrDefault();
                }
                else if (attributeName.Equals("ValidateScript", StringComparison.OrdinalIgnoreCase))
                {
                    hasValidateScript = true;
                }
                else if (attributeName.Equals("Parameter", StringComparison.OrdinalIgnoreCase))
                {
                    mandatory = attribute.NamedArguments.Any(n =>
                        n.ArgumentName.Equals("Mandatory", StringComparison.OrdinalIgnoreCase)
                        && IsTrue(n));
                }
            }

            return new ToolParameter
            {
                Name = parameter.Name.VariablePath.UserPath,
                TypeName = isSwitch ? "switch" : typeName,
                IsSwitch = isSwitch,
                IsSecureString = isSecure,
                IsMandatory = mandatory,
                ValidValues = validValues,
                ValidationPattern = pattern,
                HasValidateScript = hasValidateScript,
                DefaultValue = parameter.DefaultValue?.Extent.Text,
            };
        }

        /// <summary>
        /// Mandatory can be written as Mandatory, Mandatory=$true, or
        /// Mandatory = $True. A bare named argument carries an implicit true,
        /// which the AST represents with ExpressionOmitted.
        /// </summary>
        private static bool IsTrue(NamedAttributeArgumentAst argument)
        {
            if (argument.ExpressionOmitted)
            {
                return true;
            }

            return argument.Argument is VariableExpressionAst v
                && v.VariablePath.UserPath.Equals("true", StringComparison.OrdinalIgnoreCase);
        }
    }
}
