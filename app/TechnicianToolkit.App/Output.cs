// Output.cs - The console output model and its mapping onto the theme palette.
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
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace TechnicianToolkit.App
{
    /// <summary>
    /// A run of text sharing one colour. A single console line can hold several,
    /// because Write-Host -NoNewline lets a script build a line in pieces.
    /// </summary>
    public sealed class OutputSegment
    {
        public string Text { get; init; } = string.Empty;
        public Brush Brush { get; init; } = Brushes.Gray;
    }

    /// <summary>One line in the output pane.</summary>
    public sealed class OutputLine : INotifyPropertyChanged
    {
        // Replaced wholesale on every change rather than mutated in place. The
        // segments are surfaced through a dependency property, and a DP compares
        // references to decide whether anything happened -- mutating one list and
        // raising PropertyChanged looks like no change at all, and the line
        // silently renders empty.
        private OutputSegment[] _segments = Array.Empty<OutputSegment>();

        public IReadOnlyList<OutputSegment> Segments => _segments;

        public string Text => string.Concat(_segments.Select(s => s.Text));

        public void Append(string text, Brush brush)
        {
            var grown = new OutputSegment[_segments.Length + 1];
            Array.Copy(_segments, grown, _segments.Length);
            grown[^1] = new OutputSegment { Text = text, Brush = brush };

            _segments = grown;
            Changed();
        }

        /// <summary>
        /// A carriage return with no line feed means the script is redrawing the
        /// line in place -- the reboot countdowns in covenant and restoration do
        /// this every second. In a console the cursor returns to column zero; the
        /// equivalent here is to throw the line away and start it again.
        /// </summary>
        public void Reset()
        {
            _segments = Array.Empty<OutputSegment>();
            Changed();
        }

        private void Changed()
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Segments)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Text)));
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }

    /// <summary>
    /// TextBlock.Inlines is not a dependency property, so it cannot be bound to
    /// directly. This attached property rebuilds the inlines whenever the bound
    /// segments change, which keeps the output pane virtualised: the collection
    /// holds data, not controls.
    /// </summary>
    public static class InlineBinder
    {
        public static readonly DependencyProperty SegmentsProperty =
            DependencyProperty.RegisterAttached(
                "Segments",
                typeof(IReadOnlyList<OutputSegment>),
                typeof(InlineBinder),
                new PropertyMetadata(null, OnSegmentsChanged));

        public static void SetSegments(DependencyObject element, IReadOnlyList<OutputSegment> value) =>
            element.SetValue(SegmentsProperty, value);

        public static IReadOnlyList<OutputSegment> GetSegments(DependencyObject element) =>
            (IReadOnlyList<OutputSegment>)element.GetValue(SegmentsProperty);

        private static void OnSegmentsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not TextBlock block)
            {
                return;
            }

            block.Inlines.Clear();

            if (e.NewValue is not IReadOnlyList<OutputSegment> segments)
            {
                return;
            }

            foreach (OutputSegment segment in segments)
            {
                block.Inlines.Add(new Run(segment.Text) { Foreground = segment.Brush });
            }
        }
    }

    /// <summary>
    /// Maps the console colours the suite actually writes with onto the report
    /// palette, so a green success line in the window is the same green as in
    /// the HTML report the same run produces.
    /// </summary>
    public static class ConsolePalette
    {
        private static readonly Dictionary<ConsoleColor, Brush> Map = Build();

        public static Brush Default { get; } = Freeze("#FFC8D4E0");

        public static Brush For(ConsoleColor? color) =>
            color.HasValue && Map.TryGetValue(color.Value, out Brush? brush) ? brush : Default;

        private static Dictionary<ConsoleColor, Brush> Build()
        {
            Brush text = Freeze("#FFC8D4E0");
            Brush dim = Freeze("#FF637587");

            return new Dictionary<ConsoleColor, Brush>
            {
                // The colour schema every tool declares at the top of its file.
                [ConsoleColor.Cyan] = Freeze("#FF00E5CC"),
                [ConsoleColor.DarkCyan] = Freeze("#FF00B3A0"),
                [ConsoleColor.Green] = Freeze("#FF3FB950"),
                [ConsoleColor.DarkGreen] = Freeze("#FF2E8B3C"),
                [ConsoleColor.Yellow] = Freeze("#FFE3B341"),
                [ConsoleColor.DarkYellow] = Freeze("#FFB98F33"),
                [ConsoleColor.Red] = Freeze("#FFF85149"),
                [ConsoleColor.DarkRed] = Freeze("#FFC23B35"),
                [ConsoleColor.Magenta] = Freeze("#FFC792EA"),
                [ConsoleColor.DarkMagenta] = Freeze("#FF9A6FC0"),
                [ConsoleColor.Blue] = Freeze("#FF58A6FF"),
                [ConsoleColor.DarkBlue] = Freeze("#FF3E7BC0"),
                [ConsoleColor.White] = text,
                [ConsoleColor.Gray] = dim,
                [ConsoleColor.DarkGray] = Freeze("#FF3D4A5C"),
                [ConsoleColor.Black] = Freeze("#FF0A0E14"),
            };
        }

        private static Brush Freeze(string hex)
        {
            var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
            brush.Freeze();
            return brush;
        }
    }
}
