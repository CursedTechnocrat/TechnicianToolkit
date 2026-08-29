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

namespace TechnicianToolkit.Engine
{
    /// <summary>
    /// Every toolkit script, the shared module, config.json and the licence are
    /// embedded in this assembly. They are written to a working directory before
    /// anything runs.
    ///
    /// Lifted from the launcher prototype (launcher/Program.cs, added in #22),
    /// which is superseded by this application.
    ///
    /// Two properties of this are load-bearing and must not be optimised away:
    ///
    /// 1. Scripts run BY PATH, never from a string. Every tool bootstrap block
    ///    depends on $PSScriptRoot and $PSCommandPath resolving to a real
    ///    location, and on the shared module sitting beside the script.
    /// 2. Because the whole suite is present locally, neither the GRIMOIRE hub
    ///    nor any tool ever reaches out to GitHub for a missing file. That is
    ///    the offline guarantee, and it holds exactly as it does today.
    /// </summary>
    public static class ScriptExtractor
    {
        /// <summary>Prefix applied to every embedded resource by the .csproj.</summary>
        private const string ResourcePrefix = "TKScripts.";

        private const string WorkFolderName = "TechnicianToolkit";

        /// <summary>
        /// Extract the suite and return the directory holding it.
        /// </summary>
        public static string Extract() => Extract(out _);

        /// <summary>
        /// Extract the suite, reporting how many files were written. Callers must
        /// use this count rather than counting the directory afterwards: tools
        /// write their HTML reports into the same working directory, so a
        /// directory listing grows with every run.
        /// </summary>
        public static string Extract(out int fileCount)
        {
            string workDir = ResolveWorkDir();
            fileCount = Extract(workDir);
            return workDir;
        }

        /// <summary>
        /// Extract the suite into a specific directory. Files are overwritten on
        /// every run so what executes is always the known-good embedded version --
        /// that is what makes the build reproducible in the field.
        /// </summary>
        public static int Extract(string workDir)
        {
            Directory.CreateDirectory(workDir);

            // The resources live beside this type, so this works identically in
            // the console harness and in the phase 02 window.
            Assembly assembly = typeof(ScriptExtractor).Assembly;

            List<string> resources = assembly
                .GetManifestResourceNames()
                .Where(n => n.StartsWith(ResourcePrefix, StringComparison.Ordinal))
                .ToList();

            int count = 0;
            foreach (string resource in resources)
            {
                string target = Path.Combine(workDir, resource.Substring(ResourcePrefix.Length));

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
        /// Prefer a writable folder next to the executable, which keeps
        /// everything on the USB stick together. Fall back to TEMP when the
        /// medium is read-only.
        /// </summary>
        private static string ResolveWorkDir()
        {
            string preferred = Path.Combine(AppContext.BaseDirectory, WorkFolderName);
            if (TryEnsureWritable(preferred))
            {
                return preferred;
            }

            string fallback = Path.Combine(Path.GetTempPath(), WorkFolderName);
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
