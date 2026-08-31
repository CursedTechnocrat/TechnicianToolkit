// Probe.cs - Headless verification mode for the Phase 00 spike.
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
using System.Text;

namespace TechnicianToolkit.Spike
{
    /// <summary>
    /// A WinExe has no console to write to, so the probe collects everything
    /// into a log file instead. This is what makes the spike verifiable from a
    /// script — and, later, from CI on a clean runner.
    /// </summary>
    internal sealed class ProbeSink : IHostSink
    {
        private readonly StringBuilder _log = new StringBuilder();

        private readonly HashSet<ConsoleColor> _colors = new HashSet<ConsoleColor>();

        internal int ExitCallCount { get; private set; }
        internal int PromptCallCount { get; private set; }
        internal int ClearCallCount { get; private set; }
        internal int ProgressCallCount { get; private set; }
        internal IReadOnlyCollection<ConsoleColor> ColorsSeen => _colors;

        public void Write(string text, ConsoleColor? foreground, ConsoleColor? background)
        {
            if (foreground.HasValue)
            {
                _colors.Add(foreground.Value);
            }
            _log.Append(text);
        }

        public void Clear()
        {
            ClearCallCount++;
            _log.AppendLine("[host] Clear-Host reached SetBufferContents");
        }

        public void Progress(string activity, string status, int percentComplete, bool completed)
        {
            ProgressCallCount++;
            _log.AppendLine($"[host] progress: {activity} — {status} ({percentComplete}%)");
        }

        public string? ReadLine(string? prompt)
        {
            PromptCallCount++;
            _log.AppendLine($"[host] Read-Host reached the host (prompt: {prompt ?? "none"})");
            return string.Empty;
        }

        public int PromptForChoice(
            string caption, string message, IReadOnlyList<string> choices, int defaultChoice)
        {
            PromptCallCount++;
            _log.AppendLine($"[host] PromptForChoice reached the host: {caption}");
            return defaultChoice < 0 ? 0 : defaultChoice;
        }

        public void ScriptRequestedExit(int exitCode)
        {
            ExitCallCount++;
            _log.AppendLine($"[host] script called exit {exitCode} — absorbed by SetShouldExit");
        }

        public override string ToString() => _log.ToString();
    }

    internal static class Probe
    {
        /// <summary>
        /// Runs both spike checks and writes a report. Returns a process exit
        /// code: 0 if the PowerShell engine loaded and ran, 1 if it did not.
        /// </summary>
        internal static int Run(string outputPath)
        {
            var report = new StringBuilder();
            int exitCode = 0;

            report.AppendLine("TechnicianToolkit Phase 00 spike — probe report");
            report.AppendLine("================================================");
            report.AppendLine($"Executable : {Environment.ProcessPath}");
            report.AppendLine($"BaseDir    : {AppContext.BaseDirectory}");
            report.AppendLine($"Elevated   : {IsElevated()}");
            report.AppendLine();

            // ── Check 1: does the engine load and work at all? ───────────────
            report.AppendLine("--- Check 1: runtime ---");
            var runtimeSink = new ProbeSink();
            try
            {
                string runtime = new ToolRunner(runtimeSink).RuntimeReport();
                report.AppendLine(runtime);
                if (runtime.StartsWith("FAILED", StringComparison.Ordinal) ||
                    runtime.Contains("CIM FAILED", StringComparison.Ordinal))
                {
                    exitCode = 1;
                }
            }
            catch (Exception ex)
            {
                report.AppendLine($"THREW: {ex.GetType().FullName}: {ex.Message}");
                exitCode = 1;
            }
            report.AppendLine();

            // ── Check 2: does a real toolkit script run? ─────────────────────
            report.AppendLine("--- Check 2: WARD ---");
            var wardSink = new ProbeSink();
            try
            {
                var runner = new ToolRunner(wardSink);
                string workDir = ToolRunner.ExtractScripts();
                report.AppendLine($"Extracted to: {workDir}");

                string ward = Path.Combine(workDir, "ward.ps1");
                IReadOnlyList<string> errors = runner.RunScript(
                    ward, new Dictionary<string, object> { ["Unattended"] = true });

                report.AppendLine($"Error records : {errors.Count}");
                report.AppendLine($"exit calls    : {wardSink.ExitCallCount}");
                report.AppendLine($"prompt calls  : {wardSink.PromptCallCount}");
                report.AppendLine($"Clear-Host    : {wardSink.ClearCallCount}");
                report.AppendLine();
                report.AppendLine("---- WARD output ----");
                report.AppendLine(wardSink.ToString());
                foreach (string error in errors)
                {
                    report.AppendLine("ERR: " + error);
                }
            }
            catch (Exception ex)
            {
                report.AppendLine($"THREW: {ex.GetType().FullName}: {ex.Message}");
                report.AppendLine(ex.StackTrace);
                exitCode = 1;
            }
            report.AppendLine();

            // ── Check 3: does the host actually carry everything? ────────────
            // This is the load-bearing claim: that Write-Host, prompts, streams
            // and Clear-Host all route through the host, so tool scripts need
            // no rewriting. Unlike Check 2 it needs no Administrator rights.
            report.AppendLine("--- Check 3: host surface ---");
            var surfaceSink = new ProbeSink();
            try
            {
                string workDir = ToolRunner.ExtractScripts();
                new ToolRunner(surfaceSink).ExerciseHostSurface(workDir);

                var colors = new List<string>();
                foreach (ConsoleColor color in surfaceSink.ColorsSeen)
                {
                    colors.Add(color.ToString());
                }
                colors.Sort();

                report.AppendLine($"colors seen   : {colors.Count} ({string.Join(", ", colors)})");
                report.AppendLine($"Clear-Host    : {surfaceSink.ClearCallCount}");
                report.AppendLine($"prompt calls  : {surfaceSink.PromptCallCount}");
                report.AppendLine($"progress calls: {surfaceSink.ProgressCallCount}");

                bool ok = colors.Count >= 4
                       && surfaceSink.ClearCallCount >= 1
                       && surfaceSink.PromptCallCount >= 1
                       && surfaceSink.ProgressCallCount >= 1;
                report.AppendLine($"VERDICT       : {(ok ? "PASS" : "FAIL")}");
                if (!ok)
                {
                    exitCode = 1;
                }

                report.AppendLine();
                report.AppendLine("---- host surface output ----");
                report.AppendLine(surfaceSink.ToString());
            }
            catch (Exception ex)
            {
                report.AppendLine($"THREW: {ex.GetType().FullName}: {ex.Message}");
                exitCode = 1;
            }

            File.WriteAllText(outputPath, report.ToString());
            return exitCode;
        }

        private static bool IsElevated()
        {
            try
            {
                using var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
                var principal = new System.Security.Principal.WindowsPrincipal(identity);
                return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }
    }
}
