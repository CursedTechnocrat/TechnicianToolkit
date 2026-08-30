// RunHistory.cs - What ran, when, how it ended, and what it produced.
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
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace TechnicianToolkit.Engine
{
    /// <summary>How a run ended.</summary>
    public enum RunOutcome
    {
        Succeeded,
        CompletedWithErrors,
        Cancelled,

        /// <summary>
        /// The tool asserted Administrator and did no work. Kept distinct from a
        /// failure because nothing was attempted, and distinct from success
        /// because nothing was achieved.
        /// </summary>
        RefusedNeedsAdmin,
    }

    /// <summary>One entry in the history.</summary>
    public sealed class RunRecord
    {
        public string Tool { get; init; } = string.Empty;
        public DateTime StartedUtc { get; init; }
        public double DurationSeconds { get; init; }
        public RunOutcome Outcome { get; init; }
        public int ErrorCount { get; init; }
        public int? ExitCode { get; init; }

        /// <summary>Parameters as the form supplied them, secrets excluded.</summary>
        public string Parameters { get; init; } = string.Empty;

        /// <summary>Full paths of the artifacts this run produced.</summary>
        public List<string> Artifacts { get; init; } = new List<string>();

        [JsonIgnore]
        public DateTime StartedLocal => StartedUtc.ToLocalTime();

        [JsonIgnore]
        public string OutcomeDisplay => Outcome switch
        {
            RunOutcome.Succeeded => "OK",
            RunOutcome.CompletedWithErrors => ErrorCount + " ERROR" + (ErrorCount == 1 ? string.Empty : "S"),
            RunOutcome.Cancelled => "CANCELLED",
            RunOutcome.RefusedNeedsAdmin => "REFUSED",
            _ => "?",
        };
    }

    /// <summary>
    /// The run history, persisted beside the suite and the reports.
    ///
    /// It lives in the root folder rather than in the report directory on
    /// purpose: the report directory is what the window watches for a
    /// technician's artifacts, and the application's own bookkeeping does not
    /// belong in the list they are reading.
    /// </summary>
    public sealed class RunHistory
    {
        private const string FileName = "history.json";

        /// <summary>
        /// Entries beyond this are dropped oldest-first. A history is for
        /// answering "what did I already run on this machine", which is a
        /// question about today, not about last quarter.
        /// </summary>
        private const int MaxEntries = 500;

        private static readonly JsonSerializerOptions Json = new JsonSerializerOptions
        {
            WriteIndented = true,
            Converters = { new JsonStringEnumConverter() },
        };

        private readonly string _path;
        private readonly List<RunRecord> _records = new List<RunRecord>();

        public RunHistory(string rootDirectory)
        {
            _path = Path.Combine(rootDirectory, FileName);
            Load();
        }

        /// <summary>Newest first.</summary>
        public IReadOnlyList<RunRecord> Records => _records;

        public void Add(RunRecord record)
        {
            _records.Insert(0, record);

            while (_records.Count > MaxEntries)
            {
                _records.RemoveAt(_records.Count - 1);
            }

            Save();
        }

        public void Clear()
        {
            _records.Clear();
            Save();
        }

        private void Load()
        {
            if (!File.Exists(_path))
            {
                return;
            }

            try
            {
                List<RunRecord>? loaded = JsonSerializer.Deserialize<List<RunRecord>>(
                    File.ReadAllText(_path), Json);

                if (loaded != null)
                {
                    _records.AddRange(loaded);
                }
            }
            catch (Exception)
            {
                // A corrupt or hand-edited history is not worth refusing to start
                // over. It is a convenience record, not the source of truth for
                // anything -- the reports themselves are.
            }
        }

        private void Save()
        {
            try
            {
                string? directory = Path.GetDirectoryName(_path);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.WriteAllText(_path, JsonSerializer.Serialize(_records, Json), new UTF8Encoding(false));
            }
            catch (Exception)
            {
                // Read-only medium, or the folder went away. Losing the history is
                // survivable; failing a completed run because of it is not.
            }
        }

        /// <summary>
        /// Render the parameters a run was given, with anything secret left out.
        /// A SecureString never reaches here as text, but the key name would still
        /// disclose that a credential was supplied, which is fine, while its value
        /// must never be written to disk.
        /// </summary>
        public static string DescribeParameters(IDictionary<string, object?> parameters)
        {
            if (parameters.Count == 0)
            {
                return string.Empty;
            }

            return string.Join("  ", parameters.Select(p =>
            {
                if (p.Value is System.Security.SecureString)
                {
                    return "-" + p.Key + " ********";
                }

                if (p.Value is bool flag)
                {
                    return flag ? "-" + p.Key : string.Empty;
                }

                return "-" + p.Key + " " + p.Value;
            }).Where(text => text.Length > 0));
        }
    }
}
