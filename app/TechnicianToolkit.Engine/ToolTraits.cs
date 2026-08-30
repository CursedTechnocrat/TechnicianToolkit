// ToolTraits.cs - What a front end needs to know about a tool before running it.
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
using System.IO;
using System.Linq;

namespace TechnicianToolkit.Engine
{
    /// <summary>
    /// The facts a catalog badge or a form needs, read from the script itself so
    /// nothing has to be maintained in a second place.
    /// </summary>
    public sealed class ToolTrait
    {
        /// <summary>
        /// The tool calls Assert-AdminPrivilege or Invoke-AdminElevation, so it
        /// will refuse to run without Administrator.
        /// </summary>
        public bool RequiresAdmin { get; init; }

        /// <summary>
        /// The tool declares -WhatIf, which in this suite is how a tool says it
        /// changes the machine. Read-only tools do not offer it.
        /// </summary>
        public bool IsDestructive { get; init; }

        /// <summary>The tool declares -Unattended, so a form can drive it without prompts.</summary>
        public bool SupportsUnattended { get; init; }
    }

    public static class ToolTraits
    {
        /// <summary>
        /// Inspect one tool. The admin check is a source scan rather than an AST
        /// walk because the call sits at the top level of the script, before any
        /// structure worth parsing for.
        /// </summary>
        public static ToolTrait Inspect(string scriptPath)
        {
            if (!File.Exists(scriptPath))
            {
                throw new FileNotFoundException("Tool script not found.", scriptPath);
            }

            string source = File.ReadAllText(scriptPath);
            bool requiresAdmin =
                source.Contains("Assert-AdminPrivilege", StringComparison.Ordinal)
                || source.Contains("Invoke-AdminElevation", StringComparison.Ordinal);

            IReadOnlyList<ToolParameter> parameters = ToolParameters.Load(scriptPath);

            return new ToolTrait
            {
                RequiresAdmin = requiresAdmin,
                IsDestructive = parameters.Any(p => p.Name.Equals("WhatIf", StringComparison.OrdinalIgnoreCase)),
                SupportsUnattended = parameters.Any(p => p.Name.Equals("Unattended", StringComparison.OrdinalIgnoreCase)),
            };
        }
    }
}
