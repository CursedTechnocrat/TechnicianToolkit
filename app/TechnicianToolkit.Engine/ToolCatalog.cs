// ToolCatalog.cs - Reads the GRIMOIRE tool registry out of grimoire.ps1 by AST.
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
using System.Linq;
using System.Management.Automation.Language;

namespace TechnicianToolkit.Engine
{
    /// <summary>One entry in the GRIMOIRE registry.</summary>
    public sealed class ToolEntry
    {
        public string Key { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string File { get; init; } = string.Empty;
        public string Version { get; init; } = string.Empty;
        public string Description { get; init; } = string.Empty;
        public string Color { get; init; } = string.Empty;
        public string Category { get; init; } = string.Empty;

        /// <summary>
        /// The name without its dots, e.g. WARD for W.A.R.D. -- what a person
        /// types on the command line rather than the display form.
        /// </summary>
        public string ShortName => Name.Replace(".", string.Empty).Trim();

        public override string ToString() => ShortName + " (" + File + ")";
    }

    /// <summary>
    /// grimoire.ps1 already holds the single source of truth for the suite:
    /// names, files, versions, colours and categories. Reading it by AST rather
    /// than duplicating it in C# means the catalog can never drift from the hub.
    ///
    /// The Pester suite already extracts data tables from these scripts the same
    /// way, so the technique has precedent in this repository.
    /// </summary>
    public static class ToolCatalog
    {
        /// <summary>
        /// Parse grimoire.ps1 and return its registry in declaration order.
        /// </summary>
        public static IReadOnlyList<ToolEntry> Load(string grimoirePath)
        {
            if (!System.IO.File.Exists(grimoirePath))
            {
                throw new FileNotFoundException("grimoire.ps1 was not found.", grimoirePath);
            }

            Ast root = ParseFile(grimoirePath);
            Ast? toolsValue = FindAssignedValue(root, "Tools");
            if (toolsValue == null)
            {
                throw new InvalidOperationException(
                    "No $Tools assignment found in " + grimoirePath + ". The registry shape has changed.");
            }

            // Each registry entry is a [PSCustomObject]@{ ... } literal, so every
            // hashtable under the assignment is one tool.
            var entries = new List<ToolEntry>();
            foreach (HashtableAst table in toolsValue.FindAll(a => a is HashtableAst, searchNestedScriptBlocks: true).Cast<HashtableAst>())
            {
                Dictionary<string, string> fields = ReadStringFields(table);
                if (!fields.ContainsKey("File"))
                {
                    // Not a tool entry -- some other hashtable caught by the sweep.
                    continue;
                }

                entries.Add(new ToolEntry
                {
                    Key = Get(fields, "Key"),
                    Name = Get(fields, "Name"),
                    File = Get(fields, "File"),
                    Version = Get(fields, "Version"),
                    Description = Get(fields, "Description"),
                    Color = Get(fields, "Color"),
                    Category = Get(fields, "Category"),
                });
            }

            return entries;
        }

        /// <summary>
        /// The category display order, also declared in grimoire.ps1. Returns an
        /// empty list if the hub stops declaring one.
        /// </summary>
        public static IReadOnlyList<string> LoadCategoryOrder(string grimoirePath)
        {
            Ast root = ParseFile(grimoirePath);
            Ast? value = FindAssignedValue(root, "CategoryOrder");
            if (value == null)
            {
                return Array.Empty<string>();
            }

            return value
                .FindAll(a => a is StringConstantExpressionAst, searchNestedScriptBlocks: true)
                .Cast<StringConstantExpressionAst>()
                .Select(s => s.Value)
                .ToList();
        }

        /// <summary>
        /// Resolve a tool by registry key (the menu number), by short name
        /// (WARD), or by file name (ward.ps1). Matching is case-insensitive
        /// because nobody types W.A.R.D. with the right case under pressure.
        /// </summary>
        public static ToolEntry? Resolve(IReadOnlyList<ToolEntry> catalog, string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return null;
            }

            token = token.Trim();

            return catalog.FirstOrDefault(t =>
                       string.Equals(t.Key, token, StringComparison.OrdinalIgnoreCase))
                ?? catalog.FirstOrDefault(t =>
                       string.Equals(t.ShortName, token.Replace(".", string.Empty), StringComparison.OrdinalIgnoreCase))
                ?? catalog.FirstOrDefault(t =>
                       string.Equals(t.File, token, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(Path.GetFileNameWithoutExtension(t.File), token, StringComparison.OrdinalIgnoreCase));
        }

        internal static Ast ParseFile(string path)
        {
            Token[] tokens;
            ParseError[] errors;
            Ast ast = Parser.ParseFile(path, out tokens, out errors);

            // A parse error in a tool is worth surfacing loudly: it means the
            // script would not run either.
            if (errors != null && errors.Length > 0)
            {
                throw new InvalidOperationException(
                    "Could not parse " + Path.GetFileName(path) + ": " + errors[0].Message
                    + " (line " + errors[0].Extent.StartLineNumber + ")");
            }

            return ast;
        }

        /// <summary>
        /// Find the right-hand side of the first top-level assignment to a named
        /// variable.
        /// </summary>
        private static Ast? FindAssignedValue(Ast root, string variableName)
        {
            AssignmentStatementAst? assignment = root
                .FindAll(a => a is AssignmentStatementAst, searchNestedScriptBlocks: true)
                .Cast<AssignmentStatementAst>()
                .FirstOrDefault(a =>
                    a.Left is VariableExpressionAst v
                    && string.Equals(v.VariablePath.UserPath, variableName, StringComparison.OrdinalIgnoreCase));

            return assignment?.Right;
        }

        /// <summary>
        /// Read a hashtable literal into a string dictionary, keeping only the
        /// entries whose value is a plain string or number.
        /// </summary>
        private static Dictionary<string, string> ReadStringFields(HashtableAst table)
        {
            var fields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var pair in table.KeyValuePairs)
            {
                if (pair.Item1 is not StringConstantExpressionAst key)
                {
                    continue;
                }

                // A hashtable value is a statement, not a bare expression: the
                // parser wraps 'Key = value' as a one-element pipeline.
                // GetPureExpression unwraps exactly that case and returns null
                // for anything with real pipeline machinery in it.
                ExpressionAst? expression = (pair.Item2 as PipelineBaseAst)?.GetPureExpression();

                string? value = expression switch
                {
                    StringConstantExpressionAst s => s.Value,
                    ConstantExpressionAst c => c.Value?.ToString(),
                    _ => null,
                };

                if (value != null)
                {
                    fields[key.Value] = value;
                }
            }

            return fields;
        }

        private static string Get(Dictionary<string, string> fields, string key) =>
            fields.TryGetValue(key, out string? value) ? value : string.Empty;
    }
}
