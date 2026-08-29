// ConsoleSink.cs - Renders the hosted engine output to a real console.
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
using System.Management.Automation;
using System.Security;
using TechnicianToolkit.Engine;

namespace TechnicianToolkit.Harness
{
    /// <summary>
    /// The console implementation of the host sink.
    ///
    /// This is deliberately the dullest possible front end: it proves the engine
    /// carries everything a tool emits, without a GUI in the way. The phase 02
    /// window implements the same interface and marshals onto the dispatcher.
    /// </summary>
    internal sealed class ConsoleSink : IHostSink
    {
        private readonly object _gate = new object();
        private string _lastProgressActivity = string.Empty;

        /// <summary>Set when a script calls exit, so the harness can report it.</summary>
        internal int? LastExitCode { get; private set; }

        public void Write(string text, ConsoleColor? foreground, ConsoleColor? background)
        {
            lock (_gate)
            {
                ConsoleColor previousFore = Console.ForegroundColor;
                ConsoleColor previousBack = Console.BackgroundColor;

                if (foreground.HasValue)
                {
                    Console.ForegroundColor = foreground.Value;
                }
                if (background.HasValue)
                {
                    Console.BackgroundColor = background.Value;
                }

                Console.Write(text);

                Console.ForegroundColor = previousFore;
                Console.BackgroundColor = previousBack;
            }
        }

        public void Clear()
        {
            lock (_gate)
            {
                try
                {
                    Console.Clear();
                }
                catch (System.IO.IOException)
                {
                    // Output is redirected to a file or a pipe; there is nothing
                    // to clear and that must not take the run down.
                }
            }
        }

        public void Progress(string activity, string status, int percentComplete, bool completed)
        {
            lock (_gate)
            {
                if (completed)
                {
                    Console.WriteLine("[progress] " + activity + " - done");
                    _lastProgressActivity = string.Empty;
                    return;
                }

                // Only re-announce the activity when it changes, or a tool with a
                // busy progress loop drowns its own output.
                string prefix = activity == _lastProgressActivity ? "  " : "[progress] " + activity + ": ";
                _lastProgressActivity = activity;

                string percent = percentComplete >= 0 ? " (" + percentComplete + "%)" : string.Empty;
                Console.WriteLine(prefix + status + percent);
            }
        }

        public string? ReadLine(string? prompt)
        {
            lock (_gate)
            {
                if (!string.IsNullOrEmpty(prompt))
                {
                    Console.Write(prompt + ": ");
                }
                return Console.ReadLine();
            }
        }

        public SecureString ReadLineAsSecureString(string? prompt)
        {
            lock (_gate)
            {
                Console.Write((string.IsNullOrEmpty(prompt) ? "Password" : prompt) + ": ");

                var secure = new SecureString();
                while (true)
                {
                    ConsoleKeyInfo key = Console.ReadKey(intercept: true);

                    if (key.Key == ConsoleKey.Enter)
                    {
                        Console.WriteLine();
                        break;
                    }

                    if (key.Key == ConsoleKey.Backspace)
                    {
                        if (secure.Length > 0)
                        {
                            secure.RemoveAt(secure.Length - 1);
                            Console.Write("\b \b");
                        }
                        continue;
                    }

                    if (!char.IsControl(key.KeyChar))
                    {
                        secure.AppendChar(key.KeyChar);
                        Console.Write("*");
                    }
                }

                secure.MakeReadOnly();
                return secure;
            }
        }

        public PSCredential? PromptForCredential(string caption, string message, string userName, string targetName)
        {
            lock (_gate)
            {
                if (!string.IsNullOrEmpty(caption))
                {
                    Console.WriteLine(caption);
                }
                if (!string.IsNullOrEmpty(message))
                {
                    Console.WriteLine(message);
                }
            }

            string user = string.IsNullOrEmpty(userName)
                ? ReadLine("Username") ?? string.Empty
                : userName;

            SecureString password = ReadLineAsSecureString("Password for " + user);

            if (string.IsNullOrEmpty(user))
            {
                return null;
            }

            return new PSCredential(user, password);
        }

        public int PromptForChoice(string caption, string message, IReadOnlyList<string> choices, int defaultChoice)
        {
            lock (_gate)
            {
                if (!string.IsNullOrEmpty(caption))
                {
                    Console.WriteLine();
                    Console.WriteLine(caption);
                }
                if (!string.IsNullOrEmpty(message))
                {
                    Console.WriteLine(message);
                }

                for (int i = 0; i < choices.Count; i++)
                {
                    string marker = i == defaultChoice ? " (default)" : string.Empty;
                    Console.WriteLine("  [" + i + "] " + choices[i].Replace("&", string.Empty) + marker);
                }

                while (true)
                {
                    Console.Write("Choice: ");
                    string? answer = Console.ReadLine();

                    if (string.IsNullOrWhiteSpace(answer))
                    {
                        return defaultChoice;
                    }

                    if (int.TryParse(answer, out int index) && index >= 0 && index < choices.Count)
                    {
                        return index;
                    }

                    Console.WriteLine("Enter a number between 0 and " + (choices.Count - 1) + ".");
                }
            }
        }

        public void ScriptRequestedExit(int exitCode)
        {
            // Absorbed, never acted on: this is a script ending, not the harness.
            LastExitCode = exitCode;
        }
    }
}
