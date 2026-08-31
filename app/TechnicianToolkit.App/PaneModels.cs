// PaneModels.cs - View models for the history and queue panes.
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
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using TechnicianToolkit.Engine;

namespace TechnicianToolkit.App
{
    /// <summary>One row in the history pane.</summary>
    public sealed class HistoryItem
    {
        public RunRecord Record { get; init; } = new RunRecord();

        public string WhenDisplay =>
            Record.StartedLocal.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture)
            + Environment.NewLine
            + Record.DurationSeconds.ToString("0.0", CultureInfo.InvariantCulture) + "s";

        public string ParametersDisplay => Record.Parameters;

        public string ArtifactsDisplay
        {
            get
            {
                if (Record.Artifacts.Count == 0)
                {
                    return Record.ExitCode.HasValue && Record.ExitCode.Value != 0
                        ? "no artifacts  ·  exit " + Record.ExitCode.Value
                        : "no artifacts";
                }

                return string.Join("  ", Record.Artifacts.Select(Path.GetFileName));
            }
        }

        public Brush OutcomeForeground => Brushes(Record.Outcome).Foreground;
        public Brush OutcomeBackground => Brushes(Record.Outcome).Background;

        private static (Brush Foreground, Brush Background) Brushes(RunOutcome outcome)
        {
            string key = outcome switch
            {
                RunOutcome.Succeeded => "Green",
                RunOutcome.CompletedWithErrors => "Red",
                RunOutcome.Cancelled => "TextDim",
                RunOutcome.RefusedNeedsAdmin => "Yellow",
                _ => "TextDim",
            };

            string dim = outcome switch
            {
                RunOutcome.Succeeded => "GreenDim",
                RunOutcome.CompletedWithErrors => "RedDim",
                RunOutcome.RefusedNeedsAdmin => "YellowDim",
                _ => "Surface2",
            };

            return ((Brush)Application.Current.FindResource(key),
                    (Brush)Application.Current.FindResource(dim));
        }
    }

    /// <summary>Where a queued item has got to.</summary>
    public enum QueueState
    {
        Waiting,
        Running,
        Done,
        Failed,
        Skipped,
    }

    /// <summary>
    /// One tool waiting to run, with the parameters captured at the moment it was
    /// queued. Capturing them then rather than reading the form later is what
    /// lets a technician queue the same tool twice with different arguments,
    /// which is the whole point of mirroring what RITUAL does for recipes.
    /// </summary>
    public sealed class QueueItem : INotifyPropertyChanged
    {
        private QueueState _state = QueueState.Waiting;
        private int _position;

        public ToolItem Tool { get; init; } = new ToolItem();

        public Dictionary<string, object?> Parameters { get; init; } = new Dictionary<string, object?>();

        public int Position
        {
            get => _position;
            set { _position = value; Changed(nameof(Position)); }
        }

        public QueueState State
        {
            get => _state;
            set { _state = value; Changed(nameof(State)); Changed(nameof(StateDisplay)); }
        }

        public string StateDisplay => State switch
        {
            QueueState.Waiting => "waiting",
            QueueState.Running => "running...",
            QueueState.Done => "done",
            QueueState.Failed => "failed",
            QueueState.Skipped => "skipped",
            _ => string.Empty,
        };

        public string ParametersDisplay
        {
            get
            {
                string described = RunHistory.DescribeParameters(Parameters);
                return described.Length == 0 ? "(no parameters)" : described;
            }
        }

        private void Changed(string property) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));

        public event PropertyChangedEventHandler? PropertyChanged;
    }
}
