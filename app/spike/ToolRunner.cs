// ToolRunner.cs - Extracts embedded scripts and runs one in a hosted runspace.
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
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;

namespace TechnicianToolkit.Spike
{
    /// <summary>
    /// Writes the embedded scripts to disk and runs one of them in a runspace
    /// bound to <see cref="SpikeHost"/>.
    ///
    /// Scripts are run BY PATH, never from a string: every toolkit script's
    /// bootstrap block depends on $PSScriptRoot and $PSCommandPath resolving to
    /// a real location, and on the shared module sitting beside it.
    /// </summary>
    internal sealed class ToolRunner
    {
        private const string ResourcePrefix = "TKScripts.";

        private readonly IHostSink _sink;

        internal ToolRunner(IHostSink sink) => _sink = sink;

        /// <summary>
        /// Extract every embedded script into a working folder and return it.
        /// Lifted from the launcher prototype's Program.cs, unchanged in spirit.
        /// </summary>
        internal static string ExtractScripts()
        {
            string workDir = Path.Combine(Path.GetTempPath(), "TechnicianToolkit.Spike");
            Directory.CreateDirectory(workDir);

            Assembly assembly = Assembly.GetExecutingAssembly();
            IEnumerable<string> resources = assembly
                .GetManifestResourceNames()
                .Where(n => n.StartsWith(ResourcePrefix, StringComparison.Ordinal));

            int count = 0;
            foreach (string resource in resources)
            {
                string target = Path.Combine(workDir, resource.Substring(ResourcePrefix.Length));
                using Stream? source = assembly.GetManifestResourceStream(resource);
                if (source == null)
                {
                    continue;
                }

                using FileStream dest = File.Create(target);
                source.CopyTo(dest);
                count++;
            }

            if (count == 0)
            {
                throw new InvalidOperationException("No embedded scripts found in this build.");
            }

            return workDir;
        }

        /// <summary>
        /// Run a script file with the given parameters, streaming everything
        /// through the host. Returns the collected error records, if any.
        /// </summary>
        internal IReadOnlyList<string> RunScript(
            string scriptPath, IDictionary<string, object> parameters)
        {
            var errors = new List<string>();

            var host = new SpikeHost(_sink);
            var iss = InitialSessionState.CreateDefault2();

            // Toolkit scripts are unsigned files on disk; without this the
            // machine's execution policy decides whether the spike runs at all.
            iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;

            using Runspace runspace = RunspaceFactory.CreateRunspace(host, iss);
            runspace.Open();

            using var ps = System.Management.Automation.PowerShell.Create();
            ps.Runspace = runspace;

            ps.AddCommand(scriptPath, useLocalScope: false);
            foreach (KeyValuePair<string, object> parameter in parameters)
            {
                ps.AddParameter(parameter.Key, parameter.Value);
            }

            // Anything a script emits to the success stream without Write-Host
            // still has to be visible, so format it the way a console would.
            ps.AddCommand("Out-String").AddParameter("Stream", true);

            var output = new PSDataCollection<PSObject>();
            output.DataAdded += (_, e) =>
            {
                foreach (PSObject item in output.ReadAll())
                {
                    _sink.Write(item?.ToString() + Environment.NewLine, null, null);
                }
            };

            ps.Streams.Error.DataAdded += (_, e) =>
            {
                foreach (ErrorRecord record in ps.Streams.Error.ReadAll())
                {
                    errors.Add(record.ToString());
                    _sink.Write("ERROR: " + record + Environment.NewLine, ConsoleColor.Red, null);
                }
            };

            try
            {
                ps.Invoke(null, output);
            }
            catch (Exception ex)
            {
                errors.Add(ex.Message);
                _sink.Write("EXCEPTION: " + ex.Message + Environment.NewLine, ConsoleColor.Red, null);
            }

            return errors;
        }

        /// <summary>
        /// Exercises every path the host actually has to carry, using the real
        /// module's console helpers rather than a synthetic stand-in.
        ///
        /// This is the check that matters most: the plan's claim is that all
        /// 3,512 Write-Host calls and 220 prompts route through the host, which
        /// is why no tool script needs rewriting. That claim is what this
        /// verifies — and it needs no Administrator rights, unlike running a
        /// whole tool.
        /// </summary>
        internal void ExerciseHostSurface(string workDir)
        {
            var host = new SpikeHost(_sink);
            using Runspace runspace = RunspaceFactory.CreateRunspace(host);
            runspace.Open();

            using var ps = System.Management.Automation.PowerShell.Create();
            ps.Runspace = runspace;
            ps.AddScript(@"
                param($ModulePath)
                Import-Module $ModulePath -Force -ErrorAction Stop

                # Every colored console helper the suite writes through.
                Write-Section 'Host surface check'
                Write-Ok      'green path'
                Write-Warn    'yellow path'
                Write-Fail    'red path'
                Write-Info    'gray path'
                Write-Step    'step path'

                # The 37 Clear-Host sites.
                Clear-Host

                # Progress, then a prompt.
                Write-Progress -Activity 'Spike' -Status 'halfway' -PercentComplete 50
                Write-Progress -Activity 'Spike' -Status 'done' -Completed
                $null = Read-Host 'answer me'

                # Non-Write-Host streams.
                Write-Warning 'warning stream'
                Write-Verbose 'verbose stream' -Verbose
            ").AddParameter("ModulePath", Path.Combine(workDir, "TechnicianToolkit.psm1"));

            ps.Invoke();

            foreach (ErrorRecord record in ps.Streams.Error)
            {
                _sink.Write("ERROR: " + record + Environment.NewLine, ConsoleColor.Red, null);
            }
        }

        /// <summary>
        /// A cheap sanity check that does not touch the toolkit at all: prove
        /// the engine loaded, report its version, and confirm CIM works, since
        /// the suite leans on Get-CimInstance in 35 places.
        /// </summary>
        internal string RuntimeReport()
        {
            var host = new SpikeHost(_sink);
            using Runspace runspace = RunspaceFactory.CreateRunspace(host);
            runspace.Open();

            using var ps = System.Management.Automation.PowerShell.Create();
            ps.Runspace = runspace;
            ps.AddScript(@"
                $v   = $PSVersionTable.PSVersion.ToString()
                $ed  = $PSVersionTable.PSEdition
                $ps  = try { [string]$PSHOME } catch { '(unavailable)' }
                $os  = try { (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).Caption }
                       catch { 'CIM FAILED: ' + $_.Exception.Message }
                $mod = (Get-Module -ListAvailable Microsoft.PowerShell.Management |
                        Select-Object -First 1).Version
                ""PowerShell : $v ($ed)`nPSHOME     : $ps`nCIM        : $os`nMgmt module: $mod""
            ");

            System.Collections.ObjectModel.Collection<PSObject> result = ps.Invoke();
            if (ps.Streams.Error.Count > 0)
            {
                return "FAILED: " + string.Join("; ", ps.Streams.Error.Select(e => e.ToString()));
            }

            return result.Count > 0 ? result[0]?.ToString() ?? "(no output)" : "(no output)";
        }
    }
}
