using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

namespace TechnicianToolkit.Launcher
{
    /// <summary>
    /// Portable bootstrap for the TechnicianToolkit suite.
    ///
    /// All toolkit scripts are embedded in this executable. On launch they are
    /// written to a working directory and grimoire.ps1 (the hub menu) is started
    /// against them, with every command-line argument passed straight through.
    /// Because all scripts are present locally, nothing is ever downloaded — the
    /// program runs fully offline with no update checks, which is exactly what a
    /// USB-carried field tool needs.
    /// </summary>
    internal static class Program
    {
        // Prefix applied to every embedded script's LogicalName in the .csproj.
        private const string ResourcePrefix = "TKScripts.";

        // Subfolder created under the working root that holds the extracted suite.
        private const string WorkFolderName = "TechnicianToolkit";

        private static int Main(string[] args)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();

                string workDir = ResolveWorkDir();
                ExtractScripts(assembly, workDir);

                string grimoire = Path.Combine(workDir, "grimoire.ps1");
                if (!File.Exists(grimoire))
                {
                    Console.Error.WriteLine("[!!] grimoire.ps1 was not embedded in this build. Cannot continue.");
                    return 2;
                }

                string powershell = FindPowerShell();
                return RunGrimoire(powershell, grimoire, workDir, args);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[!!] TechnicianToolkit launcher failed:");
                Console.Error.WriteLine("     " + ex.Message);
                return 1;
            }
        }

        /// <summary>
        /// Prefer a writable folder next to the .exe (keeps everything on the USB
        /// stick and self-cleaning), but fall back to %TEMP% if the medium is
        /// read-only.
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

        /// <summary>
        /// Write every embedded script to the working directory verbatim. Bytes
        /// (including any UTF-8 BOM) are preserved exactly as they sit in the repo.
        /// Files are overwritten on every run so the suite is always the known-good
        /// embedded version — that is the "no update checks" guarantee.
        /// </summary>
        private static void ExtractScripts(Assembly assembly, string workDir)
        {
            IEnumerable<string> resources = assembly
                .GetManifestResourceNames()
                .Where(n => n.StartsWith(ResourcePrefix, StringComparison.Ordinal));

            int count = 0;
            foreach (string resource in resources)
            {
                string fileName = resource.Substring(ResourcePrefix.Length);
                string target = Path.Combine(workDir, fileName);

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
        }

        /// <summary>
        /// Prefer Windows PowerShell 5.1 (the version every toolkit script targets
        /// and which ships with every Windows install). Fall back to PowerShell 7+
        /// on PATH if the inbox copy is somehow absent.
        /// </summary>
        private static string FindPowerShell()
        {
            string systemRoot = Environment.GetEnvironmentVariable("SystemRoot") ?? @"C:\Windows";
            string winPs = Path.Combine(systemRoot, "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
            if (File.Exists(winPs))
            {
                return winPs;
            }

            // Let the OS resolve pwsh from PATH as a last resort.
            return "pwsh";
        }

        private static int RunGrimoire(string powershell, string grimoirePath, string workDir, string[] passThroughArgs)
        {
            var psi = new ProcessStartInfo
            {
                FileName = powershell,
                UseShellExecute = false,
                WorkingDirectory = workDir,
            };

            psi.ArgumentList.Add("-NoProfile");
            psi.ArgumentList.Add("-ExecutionPolicy");
            psi.ArgumentList.Add("Bypass");
            psi.ArgumentList.Add("-File");
            psi.ArgumentList.Add(grimoirePath);

            // Forward anything the user passed (e.g. -WhatIf) to the hub unchanged.
            foreach (string arg in passThroughArgs)
            {
                psi.ArgumentList.Add(arg);
            }

            using Process? process = Process.Start(psi);
            if (process == null)
            {
                Console.Error.WriteLine("[!!] Failed to start PowerShell (" + powershell + ").");
                return 3;
            }

            process.WaitForExit();
            return process.ExitCode;
        }
    }
}
