// ScriptExtractorTests.cs - Writing the embedded suite out to disk.
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

using System.IO;
using System.Linq;
using Xunit;

namespace TechnicianToolkit.Engine.Tests
{
    /// <summary>
    /// These run against the resources genuinely embedded in the engine
    /// assembly, so they also assert that the .csproj globs still catch the
    /// suite. A dropped glob makes the shipped executable useless and would
    /// otherwise only show up at runtime.
    /// </summary>
    public sealed class ScriptExtractorTests
    {
        [Fact]
        public void Extract_writes_the_whole_suite()
        {
            using var fixture = new ScriptFixture();
            string suite = Path.Combine(fixture.Directory, "suite");

            int written = ScriptExtractor.Extract(suite);

            // 42 tools plus the module, config.json and the licence. Asserted as a
            // floor so adding a tool does not fail the test, while a broken glob
            // that embeds only a handful still does.
            Assert.True(written > 40, "expected the full suite, got " + written + " files");
            Assert.True(File.Exists(Path.Combine(suite, "grimoire.ps1")));
            Assert.True(File.Exists(Path.Combine(suite, "TechnicianToolkit.psm1")));
        }

        // The GPL requires every recipient of the program to receive the licence.
        // The shipped .exe is the whole suite in one file, so the licence has to
        // travel inside it and land beside the scripts.
        [Fact]
        public void Extract_writes_the_licence_alongside_the_scripts()
        {
            using var fixture = new ScriptFixture();
            string suite = Path.Combine(fixture.Directory, "suite");

            ScriptExtractor.Extract(suite);

            string licence = Path.Combine(suite, "LICENSE");
            Assert.True(File.Exists(licence));
            Assert.Contains("GNU GENERAL PUBLIC LICENSE", File.ReadAllText(licence));
        }

        // Windows PowerShell 5.1 reads a BOM-less file as ANSI, which mangles the
        // box-drawing banners. The extractor copies bytes rather than text for
        // exactly this reason, and the Pester suite gates the BOM at the source.
        [Fact]
        public void Extract_preserves_the_UTF8_BOM_on_every_script()
        {
            using var fixture = new ScriptFixture();
            string suite = Path.Combine(fixture.Directory, "suite");

            ScriptExtractor.Extract(suite);

            foreach (string script in Directory.GetFiles(suite, "*.ps1"))
            {
                var head = new byte[3];
                using (FileStream stream = File.OpenRead(script))
                {
                    Assert.Equal(3, stream.Read(head, 0, 3));
                }

                Assert.True(
                    head[0] == 0xEF && head[1] == 0xBB && head[2] == 0xBF,
                    Path.GetFileName(script) + " lost its UTF-8 BOM in extraction");
            }
        }

        // What runs must always be the known-good embedded version. A tool edited
        // or corrupted on the stick is replaced on the next launch.
        [Fact]
        public void Extract_overwrites_a_modified_script()
        {
            using var fixture = new ScriptFixture();
            string suite = Path.Combine(fixture.Directory, "suite");
            ScriptExtractor.Extract(suite);

            string grimoire = Path.Combine(suite, "grimoire.ps1");
            File.WriteAllText(grimoire, "# tampered with");

            ScriptExtractor.Extract(suite);

            Assert.DoesNotContain("# tampered with", File.ReadAllText(grimoire));
        }

        // The one file that must survive: HEARTH and the settings screen write to
        // it, so overwriting would silently wipe a technician's org name and log
        // path on every launch.
        [Fact]
        public void Extract_seeds_config_json_but_never_overwrites_it()
        {
            using var fixture = new ScriptFixture();
            string suite = Path.Combine(fixture.Directory, "suite");

            ScriptExtractor.Extract(suite);
            string config = Path.Combine(suite, "config.json");
            Assert.True(File.Exists(config));

            File.WriteAllText(config, "{ \"OrgName\": \"Contoso\" }");
            ScriptExtractor.Extract(suite);

            Assert.Contains("Contoso", File.ReadAllText(config));
        }

        [Fact]
        public void Extract_creates_the_target_directory()
        {
            using var fixture = new ScriptFixture();
            string suite = Path.Combine(fixture.Directory, "does", "not", "exist");

            ScriptExtractor.Extract(suite);

            Assert.True(Directory.Exists(suite));
            Assert.NotEmpty(Directory.GetFiles(suite, "*.ps1"));
        }

        [Fact]
        public void Extract_reports_how_many_files_it_wrote()
        {
            using var fixture = new ScriptFixture();
            string suite = Path.Combine(fixture.Directory, "suite");

            int written = ScriptExtractor.Extract(suite);

            // config.json is skipped on the second pass, so the count drops by
            // exactly one -- which is what proves the count reflects work done
            // rather than the number of resources embedded.
            int second = ScriptExtractor.Extract(suite);

            Assert.Equal(written - 1, second);
            Assert.Equal(written, Directory.GetFiles(suite).Length);
        }
    }
}
