// Program.cs - Console harness: list, describe and run any toolkit tool headlessly.
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
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TechnicianToolkit.Engine;

namespace TechnicianToolkit.Harness
{
    /// <summary>
    /// The phase 01 exit criterion, made runnable: any tool by name with
    /// parameters, output streamed live, cancellable mid-run with Ctrl+C.
    /// </summary>
    internal static class Program
    {
        private static async Task<int> Main(string[] args)
        {
            try
            {
                if (args.Length == 0 || IsHelp(args[0]))
                {
                    PrintUsage();
                    return 0;
                }

                string command = args[0].ToLowerInvariant();
                string[] rest = args.Skip(1).ToArray();

                return command switch
                {
                    "extract" => Extract(),
                    "probe" => await ProbeAsync().ConfigureAwait(false),
                    "config" => Config(rest),
                    "list" => List(),
                    "show" => Show(rest),
                    "run" => await RunAsync(rest).ConfigureAwait(false),
                    _ => Unknown(command),
                };
            }
            catch (Exception ex)
            {
                Error(ex.Message);
                return 1;
            }
        }

        private static bool IsHelp(string arg) =>
            arg is "-h" or "--help" or "/?" or "help";

        private static void PrintUsage()
        {
            Console.WriteLine("Technician Toolkit - console harness (phase 01)");
            Console.WriteLine();
            Console.WriteLine("  extract               Write the embedded suite to disk and report where.");
            Console.WriteLine("  probe                 Open a runspace and prove the hosted engine works (CI gate).");
            Console.WriteLine("  config [set K V [S]]  Show or write toolkit configuration.");
            Console.WriteLine("  list                  List every tool in the GRIMOIRE registry.");
            Console.WriteLine("  show <tool>           Show a tool declared parameters.");
            Console.WriteLine("  run <tool> [args]     Run a tool. Ctrl+C cancels it mid-run.");
            Console.WriteLine("                        --cancel-after <n>  cancel after n seconds (for CI).");
            Console.WriteLine();
            Console.WriteLine("A tool is named by registry key, short name or file name:");
            Console.WriteLine("  run 8            run WARD            run ward.ps1");
            Console.WriteLine();
            Console.WriteLine("Parameters are passed through as they are declared:");
            Console.WriteLine("  run WARD -Unattended");
            Console.WriteLine("  run CIPHER -Action Status -Drive C");
        }

        private static int Unknown(string command)
        {
            Error("Unknown command: " + command);
            PrintUsage();
            return 2;
        }

        /// <summary>
        /// Every command needs the suite on disk first, since the tools run by
        /// path and the AST readers read the real files.
        /// </summary>
        private static string EnsureExtracted() => ScriptExtractor.Prepare().SuiteDirectory;

        private static int Extract()
        {
            ToolkitLayout layout = ScriptExtractor.Prepare();
            Console.WriteLine("Extracted " + layout.FilesWritten + " file(s).");
            Console.WriteLine("  suite   : " + layout.SuiteDirectory);
            Console.WriteLine("  reports : " + layout.ReportDirectory);
            return 0;
        }

        /// <summary>
        /// The CI gate the plan asks for: open a real runspace and prove the
        /// hosted engine works, so that a dependency bump cannot silently
        /// reintroduce either of the two single-file failures phase 00 found.
        ///
        /// It has to construct the host and run something, because that is where
        /// both failures live. Constructing PSHostUserInterface is what throws
        /// without IncludeAllContentForSelfExtract, and resolving Get-CimInstance
        /// is what fails when PowerShell's built-in modules land outside $PSHOME.
        /// Extracting and reading the catalog would exercise neither.
        /// </summary>
        private static async Task<int> ProbeAsync()
        {
            ToolkitLayout layout = ScriptExtractor.Prepare();
            Console.WriteLine("suite   : " + layout.SuiteDirectory);
            Console.WriteLine("reports : " + layout.ReportDirectory);
            Console.WriteLine("files   : " + layout.FilesWritten);
            Console.WriteLine("elevated: " + Elevation.IsElevated());
            Console.WriteLine();

            // A throwaway script beside the suite, so it runs by path with the
            // module next to it exactly as a real tool does.
            string probeScript = Path.Combine(layout.SuiteDirectory, "__probe.ps1");
            File.WriteAllText(probeScript, ProbeScript, new UTF8Encoding(true));

            try
            {
                var sink = new ConsoleSink();
                var runner = new ToolRunner(sink, layout.SuiteDirectory);

                ToolRunResult result = await runner
                    .RunAsync(probeScript, new Dictionary<string, object?>())
                    .ConfigureAwait(false);

                Console.WriteLine();
                if (result.Errors.Count > 0)
                {
                    Error("probe produced " + result.Errors.Count + " error record(s).");
                    return 1;
                }

                Console.WriteLine("[probe] engine OK.");
                return 0;
            }
            finally
            {
                try { File.Delete(probeScript); } catch { /* best effort */ }
            }
        }

        /// <summary>
        /// Deliberately touches the two things phase 00 found broken, plus the
        /// module's own console helpers so the host surface is exercised too.
        /// </summary>
        private const string ProbeScript = @"
Import-Module (Join-Path $PSScriptRoot 'TechnicianToolkit.psm1') -Force -ErrorAction Stop

Write-Section 'Engine probe'
Write-Info ('PowerShell : ' + $PSVersionTable.PSVersion + ' (' + $PSVersionTable.PSEdition + ')')
Write-Info ('PSHOME     : ' + $PSHOME)

# 35 sites in the suite depend on this resolving.
$os = (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).Caption
Write-Ok ('CIM        : ' + $os)

$mgmt = (Get-Module -ListAvailable Microsoft.PowerShell.Management | Select-Object -First 1).Version
Write-Ok ('Management : ' + $mgmt)

Write-Warn 'warning stream reaches the host'
Write-Fail 'error-coloured line reaches the host'
Write-Progress -Activity 'Probe' -Status 'halfway' -PercentComplete 50
Write-Progress -Activity 'Probe' -Status 'done' -Completed
Clear-Host
Write-Ok 'host surface OK'
";

        /// <summary>
        /// Read or write toolkit configuration, through the same Get-TKConfig and
        /// Set-TKConfig the window and hearth.ps1 use. Headless configuration is
        /// useful on its own for imaging a technician's stick, and it is what
        /// makes the settings write path testable without clicking Save.
        /// </summary>
        private static int Config(string[] args)
        {
            string suite = EnsureExtracted();

            if (args.Length == 0)
            {
                foreach (ConfigField current in ToolkitConfig.Read(suite))
                {
                    Console.WriteLine(string.Format("  {0,-32}{1}",
                        current.DisplayName,
                        string.IsNullOrEmpty(current.Value) ? "(unset)" : current.Value));
                }
                return 0;
            }

            if (!string.Equals(args[0], "set", StringComparison.OrdinalIgnoreCase) || args.Length < 3)
            {
                Error("Usage: config          (list values)");
                Error("       config set <Key> <Value> [Section]");
                return 2;
            }

            var field = new ConfigField
            {
                Key = args[1],
                Value = args[2],
                Section = args.Length > 3 ? args[3] : string.Empty,
            };

            ToolkitConfig.Write(suite, new[] { field });
            Console.WriteLine("set " + field.DisplayName + " = " + field.Value);

            // Read it back rather than trusting the write: Set-TKConfig rewrites
            // the whole file, so a bad section or a serialisation quirk would
            // otherwise only surface later, in a tool.
            ConfigField? readBack = ToolkitConfig.Read(suite)
                .FirstOrDefault(f => f.Key == field.Key && f.Section == field.Section);

            if (readBack == null || readBack.Value != field.Value)
            {
                Error("read-back mismatch: got " + (readBack?.Value ?? "(missing)"));
                return 1;
            }

            Console.WriteLine("verified.");
            return 0;
        }

        private static int List()
        {
            string workDir = EnsureExtracted();
            string grimoire = Path.Combine(workDir, "grimoire.ps1");

            IReadOnlyList<ToolEntry> catalog = ToolCatalog.Load(grimoire);
            IReadOnlyList<string> categories = ToolCatalog.LoadCategoryOrder(grimoire);

            // Anything in a category the hub does not order still has to appear,
            // or a tool could silently vanish from the listing.
            var ordered = categories.ToList();
            foreach (ToolEntry tool in catalog)
            {
                if (!string.IsNullOrEmpty(tool.Category) && !ordered.Contains(tool.Category))
                {
                    ordered.Add(tool.Category);
                }
            }

            Console.WriteLine(catalog.Count + " tools in the registry.");

            foreach (string category in ordered)
            {
                List<ToolEntry> inCategory = catalog.Where(t => t.Category == category).ToList();
                if (inCategory.Count == 0)
                {
                    continue;
                }

                Console.WriteLine();
                Console.WriteLine("  " + category);
                Console.WriteLine("  " + new string('-', Math.Max(category.Length, 40)));

                foreach (ToolEntry tool in inCategory)
                {
                    Console.WriteLine(string.Format(
                        "  {0,-4}{1,-22}{2,-8}{3}",
                        tool.Key, tool.ShortName, tool.Version, Truncate(tool.Description, 60)));
                }
            }

            return 0;
        }

        private static int Show(string[] args)
        {
            if (args.Length == 0)
            {
                Error("show needs a tool name.");
                return 2;
            }

            string workDir = EnsureExtracted();
            ToolEntry? tool = ToolCatalog.Resolve(ToolCatalog.Load(Path.Combine(workDir, "grimoire.ps1")), args[0]);
            if (tool == null)
            {
                Error("No tool matched: " + args[0]);
                return 2;
            }

            string scriptPath = Path.Combine(workDir, tool.File);
            IReadOnlyList<ToolParameter> parameters = ToolParameters.Load(scriptPath);

            Console.WriteLine(tool.Name + "  (" + tool.File + ", version " + tool.Version + ")");
            Console.WriteLine(tool.Description);
            Console.WriteLine();

            if (parameters.Count == 0)
            {
                Console.WriteLine("  Declares no parameters.");
                return 0;
            }

            Console.WriteLine("  " + parameters.Count + " parameter(s):");
            foreach (ToolParameter parameter in parameters)
            {
                string flags = parameter.IsMandatory ? " [required]" : string.Empty;
                Console.WriteLine("    -" + parameter.Name + "  <" + parameter.TypeName + ">"
                                  + "  renders as " + parameter.Control + flags);

                if (parameter.ValidValues.Count > 0)
                {
                    Console.WriteLine("        one of: " + string.Join(", ", parameter.ValidValues));
                }
                if (parameter.ValidationPattern != null)
                {
                    Console.WriteLine("        matching: " + parameter.ValidationPattern);
                }
                if (parameter.DefaultValue != null)
                {
                    Console.WriteLine("        default: " + parameter.DefaultValue);
                }
            }

            return 0;
        }

        private static async Task<int> RunAsync(string[] args)
        {
            if (args.Length == 0)
            {
                Error("run needs a tool name.");
                return 2;
            }

            string workDir = EnsureExtracted();
            ToolEntry? tool = ToolCatalog.Resolve(ToolCatalog.Load(Path.Combine(workDir, "grimoire.ps1")), args[0]);
            if (tool == null)
            {
                Error("No tool matched: " + args[0]);
                return 2;
            }

            string scriptPath = Path.Combine(workDir, tool.File);
            IReadOnlyList<ToolParameter> declared = ToolParameters.Load(scriptPath);

            // --cancel-after is the harness own flag, not the tool: it fires the
            // same cancellation Ctrl+C does, so CI can exercise the stop path
            // without a keyboard.
            string[] toolArgs = TakeCancelAfter(args.Skip(1).ToArray(), out int? cancelAfterSeconds);
            Dictionary<string, object?> parameters = ParseParameters(toolArgs, declared);

            var sink = new ConsoleSink();
            var runner = new ToolRunner(sink, workDir);

            using var cancellation = new CancellationTokenSource();
            if (cancelAfterSeconds.HasValue)
            {
                Console.WriteLine("[harness] will cancel after " + cancelAfterSeconds.Value + "s.");
                cancellation.CancelAfter(TimeSpan.FromSeconds(cancelAfterSeconds.Value));
            }

            ConsoleCancelEventHandler onCancel = (_, e) =>
            {
                // Cancel the run, not the harness: the whole point of hosting the
                // engine is that a tool can be stopped without killing the process.
                e.Cancel = true;
                Console.WriteLine();
                Console.WriteLine("[harness] cancelling...");
                cancellation.Cancel();
            };
            Console.CancelKeyPress += onCancel;

            try
            {
                WarnIfToolNeedsElevation(scriptPath);

                Console.WriteLine("[harness] running " + tool.Name + " from " + scriptPath);
                Console.WriteLine();

                ToolRunResult result = await runner
                    .RunAsync(scriptPath, parameters, cancellation.Token)
                    .ConfigureAwait(false);

                Console.WriteLine();
                if (result.Cancelled)
                {
                    Console.WriteLine("[harness] cancelled.");
                    return 130;
                }

                if (result.RefusedNeedsAdmin)
                {
                    Console.WriteLine("[harness] REFUSED: this tool asserts Administrator and did no work.");
                    return 1;
                }

                Console.WriteLine("[harness] finished. errors: " + result.Errors.Count
                                  + ", exit code: "
                                  + (result.ExitCode.HasValue ? result.ExitCode.Value.ToString() : "(none)"));

                return result.Succeeded ? 0 : 1;
            }
            finally
            {
                Console.CancelKeyPress -= onCancel;
            }
        }

        /// <summary>
        /// Pull the harness own --cancel-after flag out of the argument list so
        /// it never reaches the tool.
        /// </summary>
        private static string[] TakeCancelAfter(string[] args, out int? seconds)
        {
            seconds = null;
            var remaining = new List<string>(args.Length);

            for (int i = 0; i < args.Length; i++)
            {
                if (!string.Equals(args[i], "--cancel-after", StringComparison.OrdinalIgnoreCase))
                {
                    remaining.Add(args[i]);
                    continue;
                }

                if (i + 1 >= args.Length || !int.TryParse(args[i + 1], out int parsed))
                {
                    throw new ArgumentException("--cancel-after needs a whole number of seconds.");
                }

                seconds = parsed;
                i++;
            }

            return remaining.ToArray();
        }

        /// <summary>
        /// Turn -Name value / -Switch into a parameter dictionary, using the
        /// declared param() block to know which names take a value.
        /// </summary>
        private static Dictionary<string, object?> ParseParameters(
            string[] args, IReadOnlyList<ToolParameter> declared)
        {
            var parameters = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];
                if (!arg.StartsWith("-", StringComparison.Ordinal))
                {
                    throw new ArgumentException("Expected a -Parameter name but found: " + arg);
                }

                string name = arg.TrimStart('-');
                ToolParameter? match = declared.FirstOrDefault(p =>
                    string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));

                if (match == null)
                {
                    throw new ArgumentException(
                        "This tool declares no parameter named " + name + ". Run: show <tool>");
                }

                if (match.IsSwitch)
                {
                    parameters[match.Name] = true;
                    continue;
                }

                if (i + 1 >= args.Length)
                {
                    throw new ArgumentException("-" + match.Name + " needs a value.");
                }

                parameters[match.Name] = args[++i];
            }

            return parameters;
        }

        /// <summary>
        /// Assert-AdminPrivilege calls exit, and a script exit inside a hosted
        /// runspace never reaches PSHost.SetShouldExit -- phase 00 established
        /// that. So a tool that refuses for want of Administrator ends its
        /// pipeline silently, with no error record and no exit code to read.
        ///
        /// The shipped application sidesteps this by requesting Administrator in
        /// its manifest. The harness has no manifest, so it says so up front
        /// rather than letting a run look like it simply produced nothing.
        /// </summary>
        private static void WarnIfToolNeedsElevation(string scriptPath)
        {
            if (Elevation.IsElevated())
            {
                return;
            }

            if (!ToolTraits.Inspect(scriptPath).RequiresAdmin)
            {
                return;
            }

            ConsoleColor previous = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("[harness] This session is NOT elevated and this tool asserts Administrator.");
            Console.WriteLine("[harness] It will print its refusal and stop. Re-run the harness elevated.");
            Console.ForegroundColor = previous;
            Console.WriteLine();
        }

        private static string Truncate(string value, int max) =>
            value.Length <= max ? value : value.Substring(0, max - 3) + "...";

        private static void Error(string message)
        {
            ConsoleColor previous = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Error.WriteLine("[!!] " + message);
            Console.ForegroundColor = previous;
        }
    }
}
