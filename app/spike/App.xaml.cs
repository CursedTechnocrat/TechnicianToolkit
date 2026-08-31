// App.xaml.cs - WPF entry point for the Phase 00 spike.
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
using System.Windows;

namespace TechnicianToolkit.Spike
{
    public partial class App : Application
    {
        /// <summary>
        /// "--probe [path]" runs the spike's checks headlessly and writes a
        /// report, so a build can be verified from a script instead of by
        /// clicking buttons. Anything else opens the window.
        /// </summary>
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            if (e.Args.Length > 0 &&
                string.Equals(e.Args[0], "--probe", StringComparison.OrdinalIgnoreCase))
            {
                string outputPath = e.Args.Length > 1
                    ? e.Args[1]
                    : Path.Combine(AppContext.BaseDirectory, "probe-report.txt");

                int code = Probe.Run(outputPath);
                Shutdown(code);
                return;
            }

            new MainWindow().Show();
        }
    }
}
