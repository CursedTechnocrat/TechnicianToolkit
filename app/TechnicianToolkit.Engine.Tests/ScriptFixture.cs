// ScriptFixture.cs - Throwaway .ps1 files for the AST reader tests to parse.
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
using System.IO;

namespace TechnicianToolkit.Engine.Tests
{
    /// <summary>
    /// A temporary directory that deletes itself, plus a helper for writing
    /// script files into it.
    ///
    /// The AST readers are tested against purpose-written fixtures rather than
    /// against the repository's own scripts. A test that asserts on grimoire.ps1
    /// fails the day someone adds a tool, which teaches the team to edit the
    /// test rather than to read it -- and it cannot express the malformed cases
    /// that matter most here.
    /// </summary>
    internal sealed class ScriptFixture : IDisposable
    {
        public string Directory { get; }

        public ScriptFixture()
        {
            Directory = Path.Combine(Path.GetTempPath(), "tk-tests-" + Guid.NewGuid().ToString("N"));
            System.IO.Directory.CreateDirectory(Directory);
        }

        /// <summary>Write a script and return its full path.</summary>
        public string Write(string fileName, string content)
        {
            string path = Path.Combine(Directory, fileName);
            File.WriteAllText(path, content);
            return path;
        }

        public void Dispose()
        {
            try
            {
                if (System.IO.Directory.Exists(Directory))
                {
                    System.IO.Directory.Delete(Directory, recursive: true);
                }
            }
            catch (IOException)
            {
                // A leaked temp directory must never fail a test run.
            }
        }
    }
}
