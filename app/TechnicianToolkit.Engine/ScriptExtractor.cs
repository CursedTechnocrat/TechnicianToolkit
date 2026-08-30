// ScriptExtractor.cs - Writes the embedded toolkit suite to a working directory.
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
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace TechnicianToolkit.Engine
{
    /// <summary>Where the prepared toolkit lives on disk.</summary>
    public sealed class ToolkitLayout
    {
        /// <summary>
        /// The folder holding the suite and the reports. Application state that
        /// is neither an extracted script nor a technician's artifact -- the run
        /// history, for one -- belongs here, beside them rather than among them.
        /// </summary>
        public string RootDirectory { get; init; } = string.Empty;

        /// <summary>The extracted suite: 42 tools, the module, config.json, the licence.</summary>
        public string SuiteDirectory { get; init; } = string.Empty;

        /// <summary>
        /// Where tools write their HTML reports, CSVs and transcripts. Deliberately
        /// NOT the suite directory: that one is rewritten from the embedded copy on
        /// every launch, and something that watches for new reports should not have
        /// to ignore 45 files reappearing underneath it each time.
        /// </summary>
        public string ReportDirectory { get; init; } = string.Empty;

        public int FilesWritten { get; init; }
    }

    /// <summary>
    /// Every toolkit script, the shared module, config.json and the licence are
    /// embedded in this assembly. They are written to a working directory before
    /// anything runs.
    ///
    /// Lifted from the launcher prototype (launcher/Program.cs, added in #22),
    /// which the application superseded.
    ///
    /// Two properties of this are load-bearing and must not be optimised away:
    ///
    /// 1. Scripts run BY PATH, never from a string. Every tool bootstrap block
    ///    depends on $PSScriptRoot and $PSCommandPath resolving to a real
    ///    location, and on the shared module sitting beside it.
    /// 2. Because the whole suite is present locally, neither the GRIMOIRE hub
    ///    nor any tool ever reaches out to GitHub for a missing file. That is
    ///    the offline guarantee, and it holds exactly as it does today.
    /// </summary>
    public static class ScriptExtractor
    {
        /// <summary>Prefix applied to every embedded resource by the .csproj.</summary>
        private const string ResourcePrefix = "TKScripts.";

        private const string RootFolderName = "TechnicianToolkit";
        private const string SuiteFolderName = "suite";
        private const string ReportFolderName = "reports";
        private const string ConfigFileName = "config.json";

        /// <summary>
        /// Extract the suite, settle where reports go, and point the toolkit's own
        /// configuration at that directory.
        /// </summary>
        public static ToolkitLayout Prepare()
        {
            string root = ResolveRoot();
            string suite = Path.Combine(root, SuiteFolderName);
            string reports = Path.Combine(root, ReportFolderName);

            Directory.CreateDirectory(suite);
            Directory.CreateDirectory(reports);

            int written = Extract(suite);
            PointConfigAtReports(suite, reports);

            return new ToolkitLayout
            {
                RootDirectory = root,
                SuiteDirectory = suite,
                ReportDirectory = reports,
                FilesWritten = written,
            };
        }

        /// <summary>
        /// Extract the suite into a specific directory. Files are overwritten on
        /// every run so what executes is always the known-good embedded version --
        /// that is what makes the build reproducible in the field.
        ///
        /// config.json is the one exception: it is seeded when missing and never
        /// overwritten, because HEARTH and the settings screen write to it and a
        /// technician's org name and log path must survive the next launch.
        /// </summary>
        public static int Extract(string suiteDir)
        {
            Directory.CreateDirectory(suiteDir);

            // The resources live beside this type, so this works identically in
            // the console harness and in the window.
            Assembly assembly = typeof(ScriptExtractor).Assembly;

            List<string> resources = assembly
                .GetManifestResourceNames()
                .Where(n => n.StartsWith(ResourcePrefix, StringComparison.Ordinal))
                .ToList();

            int count = 0;
            foreach (string resource in resources)
            {
                string fileName = resource.Substring(ResourcePrefix.Length);
                string target = Path.Combine(suiteDir, fileName);

                if (string.Equals(fileName, ConfigFileName, StringComparison.OrdinalIgnoreCase)
                    && File.Exists(target))
                {
                    continue;
                }

                using Stream? source = assembly.GetManifestResourceStream(resource);
                if (source == null)
                {
                    continue;
                }

                // Byte-for-byte, so the UTF-8 BOM every script carries survives.
                // Without the BOM, Windows PowerShell 5.1 reads the file as ANSI
                // and the box-drawing banners arrive mangled.
                using FileStream dest = File.Create(target);
                source.CopyTo(dest);
                count++;
            }

            if (count == 0)
            {
                throw new InvalidOperationException(
                    "No embedded scripts found in this build. Check the EmbeddedResource globs in TechnicianToolkit.Engine.csproj.");
            }

            return count;
        }

        /// <summary>
        /// Set LogDirectory in the extracted config.json, which is what
        /// Resolve-LogDirectory reads. Without it every tool falls back to its own
        /// directory and writes its reports in among the extracted scripts.
        ///
        /// An existing non-empty LogDirectory is left alone: a technician who set
        /// one through HEARTH means it.
        /// </summary>
        private static void PointConfigAtReports(string suiteDir, string reportDir)
        {
            string configPath = Path.Combine(suiteDir, ConfigFileName);

            try
            {
                JsonNode? root = File.Exists(configPath)
                    ? JsonNode.Parse(File.ReadAllText(configPath))
                    : new JsonObject();

                if (root is not JsonObject config)
                {
                    return;
                }

                string? current = config["LogDirectory"]?.GetValue<string>();
                if (!string.IsNullOrWhiteSpace(current))
                {
                    return;
                }

                config["LogDirectory"] = reportDir;

                File.WriteAllText(
                    configPath,
                    config.ToJsonString(new JsonSerializerOptions { WriteIndented = true }),
                    new UTF8Encoding(false));
            }
            catch (Exception)
            {
                // A malformed config.json is the technician's to fix, and
                // Get-TKConfig already falls back to defaults. Reports landing in
                // the suite directory is untidy, not fatal, so this must not stop
                // the application from starting.
            }
        }

        /// <summary>
        /// Prefer a writable folder beside the executable, which keeps everything
        /// on the USB stick together.
        ///
        /// Resolved from <see cref="Environment.ProcessPath"/>, NOT from
        /// AppContext.BaseDirectory. Those differ in exactly the configuration
        /// that ships: IncludeAllContentForSelfExtract is mandatory for the
        /// PowerShell SDK under single-file, and it makes BaseDirectory point at
        /// the bundle's extraction folder under TEMP -- a hashed path that changes
        /// with every build and gets cleaned up. Reports written there would be
        /// unfindable, which defeats the point of watching for them.
        ///
        /// The fallback is LocalApplicationData rather than TEMP for the same
        /// reason: when the medium is read-only, a technician's reports still have
        /// to survive being generated.
        /// </summary>
        private static string ResolveRoot()
        {
            string? exeDir = Path.GetDirectoryName(Environment.ProcessPath ?? string.Empty);

            if (!string.IsNullOrEmpty(exeDir))
            {
                string beside = Path.Combine(exeDir, RootFolderName);
                if (TryEnsureWritable(beside))
                {
                    return beside;
                }
            }

            string fallback = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                RootFolderName);

            Directory.CreateDirectory(fallback);
            return fallback;
        }

        private static bool TryEnsureWritable(string dir)
        {
            try
            {
                Directory.CreateDirectory(dir);
                string probe = Path.Combine(dir, ".write-test");
                File.WriteAllText(probe, "ok");
                File.Delete(probe);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
