// ToolkitConfig.cs - Read and write config.json through the toolkit's own functions.
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
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text.Json;

namespace TechnicianToolkit.Engine
{
    /// <summary>One editable setting, described so a form can build itself.</summary>
    public sealed class ConfigField
    {
        /// <summary>The key Set-TKConfig takes.</summary>
        public string Key { get; init; } = string.Empty;

        /// <summary>The section, or empty for a top-level key.</summary>
        public string Section { get; init; } = string.Empty;

        public string Label { get; init; } = string.Empty;
        public string Description { get; init; } = string.Empty;
        public bool IsPath { get; init; }

        public string Value { get; set; } = string.Empty;

        public string DisplayName =>
            string.IsNullOrEmpty(Section) ? Key : Section + "." + Key;
    }

    /// <summary>
    /// The settings surface, expressed as the toolkit already expresses it.
    ///
    /// Reading and writing both go through Get-TKConfig and Set-TKConfig in a
    /// runspace rather than through the JSON file directly. That is slower, and
    /// it is the point: hearth.ps1 stays the console way to do this, and two
    /// writers with two different opinions about the file's shape is how
    /// configuration formats rot. Set-TKConfig also owns details worth not
    /// re-deriving, like creating a missing section.
    /// </summary>
    public static class ToolkitConfig
    {
        /// <summary>
        /// The fields the window offers. Deliberately not everything in
        /// config.json: the per-tool defaults are better set from the tool's own
        /// form, where their meaning is obvious.
        /// </summary>
        public static IReadOnlyList<ConfigField> Describe() => new[]
        {
            new ConfigField
            {
                Key = "OrgName",
                Label = "Organisation name",
                Description = "Stamped onto the HTML reports every tool generates.",
            },
            new ConfigField
            {
                Key = "LogDirectory",
                Label = "Report directory",
                Description = "Where tools write reports, CSVs and transcripts. The window watches this folder.",
                IsPath = true,
            },
            new ConfigField
            {
                Key = "TeamsWebhook",
                Label = "Teams webhook",
                Description = "Optional. Tools that can post a summary will use it.",
            },
            new ConfigField
            {
                Key = "DefaultDestination",
                Section = "Archive",
                Label = "ARCHIVE destination",
                Description = "Default path for pre-reimaging profile backups.",
                IsPath = true,
            },
            new ConfigField
            {
                Key = "DefaultDestination",
                Section = "Revenant",
                Label = "REVENANT destination",
                Description = "Default target for profile migration.",
                IsPath = true,
            },
            new ConfigField
            {
                Key = "DefaultTimezone",
                Section = "Covenant",
                Label = "COVENANT timezone",
                Description = "Default timezone applied during onboarding.",
            },
            new ConfigField
            {
                Key = "DefaultLocalAdminUser",
                Section = "Covenant",
                Label = "COVENANT local admin",
                Description = "Default local administrator account name.",
            },
        };

        /// <summary>
        /// Read the current values via Get-TKConfig, so whatever HEARTH last
        /// wrote is what the window shows.
        /// </summary>
        public static IReadOnlyList<ConfigField> Read(string suiteDirectory)
        {
            List<ConfigField> fields = Describe().ToList();

            string json = Invoke(suiteDirectory, @"
Import-Module (Join-Path $TkSuite 'TechnicianToolkit.psm1') -Force -ErrorAction Stop
Get-TKConfig | ConvertTo-Json -Depth 6 -Compress
");

            if (string.IsNullOrWhiteSpace(json))
            {
                return fields;
            }

            try
            {
                using JsonDocument document = JsonDocument.Parse(json);
                JsonElement root = document.RootElement;

                foreach (ConfigField field in fields)
                {
                    JsonElement holder = root;

                    if (!string.IsNullOrEmpty(field.Section))
                    {
                        if (!root.TryGetProperty(field.Section, out holder))
                        {
                            continue;
                        }
                    }

                    if (holder.ValueKind == JsonValueKind.Object
                        && holder.TryGetProperty(field.Key, out JsonElement value)
                        && value.ValueKind == JsonValueKind.String)
                    {
                        field.Value = value.GetString() ?? string.Empty;
                    }
                }
            }
            catch (JsonException)
            {
                // Get-TKConfig fills missing keys with defaults, so the only way
                // here is a genuinely unreadable file. The form falls back to
                // empty values, which Set-TKConfig will then repair on save.
            }

            return fields;
        }

        /// <summary>
        /// Write the changed fields, one Set-TKConfig call each, which is the
        /// granularity the function offers.
        /// </summary>
        public static void Write(string suiteDirectory, IEnumerable<ConfigField> changed)
        {
            List<ConfigField> list = changed.ToList();
            if (list.Count == 0)
            {
                return;
            }

            var script = new System.Text.StringBuilder();
            script.AppendLine("Import-Module (Join-Path $TkSuite 'TechnicianToolkit.psm1') -Force -ErrorAction Stop");

            for (int i = 0; i < list.Count; i++)
            {
                // Values are passed as parameters rather than interpolated, so a
                // path with a quote in it cannot rewrite the script.
                script.Append("Set-TKConfig -Key $Keys[").Append(i).Append("] -Value $Values[").Append(i).Append(']');

                if (!string.IsNullOrEmpty(list[i].Section))
                {
                    script.Append(" -Section $Sections[").Append(i).Append(']');
                }

                script.AppendLine();
            }

            Invoke(
                suiteDirectory,
                script.ToString(),
                new Dictionary<string, object?>
                {
                    ["Keys"] = list.Select(f => f.Key).ToArray(),
                    ["Values"] = list.Select(f => f.Value ?? string.Empty).ToArray(),
                    ["Sections"] = list.Select(f => f.Section).ToArray(),
                });
        }

        private static string Invoke(
            string suiteDirectory, string body, IDictionary<string, object?>? parameters = null)
        {
            string modulePath = Path.Combine(suiteDirectory, "TechnicianToolkit.psm1");
            if (!File.Exists(modulePath))
            {
                throw new FileNotFoundException("The shared module is not where it was extracted.", modulePath);
            }

            var initialState = InitialSessionState.CreateDefault2();
            initialState.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;

            using Runspace runspace = RunspaceFactory.CreateRunspace(initialState);
            runspace.Open();
            runspace.SessionStateProxy.Path.SetLocation(suiteDirectory);

            using var shell = PowerShell.Create();
            shell.Runspace = runspace;

            // The suite path is passed as a parameter rather than read from
            // $PSScriptRoot: a script built from a string has no script root, so
            // the automatic variable is null and Join-Path fails with a binding
            // error that names neither. The module's own $PSScriptRoot still
            // resolves, because it is imported from a real file on disk, which is
            // what lets Get-TKConfig find config.json beside it.
            var arguments = new Dictionary<string, object?>(parameters ?? new Dictionary<string, object?>())
            {
                ["TkSuite"] = suiteDirectory,
            };

            var declarations = new System.Text.StringBuilder();
            declarations.Append("param(")
                        .Append(string.Join(", ", arguments.Keys.Select(k => "$" + k)))
                        .AppendLine(")");

            shell.AddScript(declarations + body);

            foreach (KeyValuePair<string, object?> argument in arguments)
            {
                shell.AddParameter(argument.Key, argument.Value);
            }

            Collection<PSObject> output = shell.Invoke();

            if (shell.Streams.Error.Count > 0)
            {
                throw new InvalidOperationException(
                    "Configuration call failed: " + shell.Streams.Error[0]);
            }

            return output.Count > 0 ? output[0]?.ToString() ?? string.Empty : string.Empty;
        }
    }
}
