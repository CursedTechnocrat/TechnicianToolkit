// ReportIndex.cs - The artifacts a run leaves behind.
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
using System.Globalization;
using System.IO;
using System.Linq;

namespace TechnicianToolkit.Engine
{
    /// <summary>One file a tool produced: an HTML report, a CSV, a transcript.</summary>
    public sealed class ReportArtifact
    {
        public string FullPath { get; init; } = string.Empty;
        public string FileName => Path.GetFileName(FullPath);
        public long Bytes { get; init; }
        public DateTime WrittenUtc { get; init; }

        /// <summary>
        /// The tool that produced it. Every tool names its output
        /// TOOLNAME_yyyyMMdd_HHmmss.ext, so the prefix identifies the source
        /// without having to have watched the run that made it.
        /// </summary>
        public string ToolName
        {
            get
            {
                string name = Path.GetFileNameWithoutExtension(FullPath);
                int underscore = name.IndexOf('_');
                return underscore > 0 ? name.Substring(0, underscore) : name;
            }
        }

        public string WrittenLocalDisplay =>
            WrittenUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

        public string SizeDisplay =>
            Bytes >= 1024 * 1024 ? (Bytes / 1024d / 1024d).ToString("0.0", CultureInfo.InvariantCulture) + " MB"
            : Bytes >= 1024 ? (Bytes / 1024d).ToString("0.0", CultureInfo.InvariantCulture) + " KB"
            : Bytes + " B";
    }

    /// <summary>
    /// Lists what is in the report directory.
    ///
    /// Attribution works by diffing a snapshot taken before a run against one
    /// taken after, rather than by watching the filesystem during it. A tool can
    /// write several artifacts, or none, or rewrite one it wrote earlier, and a
    /// diff answers all three honestly. It also does not depend on catching an
    /// event while the UI thread is busy rendering output.
    /// </summary>
    public static class ReportIndex
    {
        public static IReadOnlyList<ReportArtifact> List(string reportDirectory)
        {
            if (string.IsNullOrEmpty(reportDirectory) || !Directory.Exists(reportDirectory))
            {
                return Array.Empty<ReportArtifact>();
            }

            var artifacts = new List<ReportArtifact>();

            foreach (string path in Directory.EnumerateFiles(reportDirectory))
            {
                FileInfo info;
                try
                {
                    info = new FileInfo(path);
                }
                catch (IOException)
                {
                    // A tool may still be writing it; it will appear next time.
                    continue;
                }

                artifacts.Add(new ReportArtifact
                {
                    FullPath = path,
                    Bytes = info.Length,
                    WrittenUtc = info.LastWriteTimeUtc,
                });
            }

            return artifacts.OrderByDescending(a => a.WrittenUtc).ToList();
        }

        /// <summary>
        /// Paths present now, for diffing against a later listing.
        /// </summary>
        public static HashSet<string> Snapshot(string reportDirectory) =>
            new HashSet<string>(
                List(reportDirectory).Select(a => a.FullPath),
                StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// What appeared -- or was rewritten -- since the snapshot was taken.
        /// </summary>
        public static IReadOnlyList<ReportArtifact> Since(
            string reportDirectory, HashSet<string> before, DateTime startedUtc) =>
            List(reportDirectory)
                .Where(a => !before.Contains(a.FullPath) || a.WrittenUtc >= startedUtc)
                .ToList();
    }
}
