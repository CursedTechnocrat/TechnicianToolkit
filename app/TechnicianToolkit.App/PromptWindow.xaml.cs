// PromptWindow.xaml.cs - Turns the host's prompt callbacks into dialogs.
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
using System.Management.Automation;
using System.Security;
using System.Windows;
using System.Windows.Controls;

namespace TechnicianToolkit.App
{
    /// <summary>
    /// One dialog covering every prompt shape the host can raise: a line of
    /// text, a masked secret, a credential pair, and a choice between labelled
    /// options.
    ///
    /// The plan is explicit that these are the fallback path, not the primary
    /// one -- covenant.ps1 alone has 26 Read-Host calls, and answering each in a
    /// separate modal would be miserable. The generated form driving -Unattended
    /// is what avoids reaching this in the normal case.
    /// </summary>
    public partial class PromptWindow : Window
    {
        private string? _line;
        private SecureString? _secure;
        private PSCredential? _credential;
        private int _choice = -1;

        private PromptWindow()
        {
            InitializeComponent();
        }

        private static PromptWindow Create(Window? owner, string caption, string message)
        {
            var window = new PromptWindow();

            // A prompt raised before the main window is up (or during a headless
            // render) has no owner; centring on the screen is the fallback.
            if (owner != null && owner.IsLoaded)
            {
                window.Owner = owner;
            }
            else
            {
                window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }

            window.CaptionText.Text = string.IsNullOrWhiteSpace(caption) ? "INPUT REQUESTED" : caption.ToUpperInvariant();
            window.MessageText.Text = message;
            return window;
        }

        /// <summary>Read-Host.</summary>
        public static string? AskLine(Window? owner, string? prompt)
        {
            PromptWindow window = Create(owner, "INPUT REQUESTED",
                string.IsNullOrWhiteSpace(prompt) ? "The running tool is waiting for input." : prompt!);

            var box = new TextBox { Margin = new Thickness(0) };
            window.InputHost.Children.Add(box);
            window.AddConfirmCancel(() => window._line = box.Text);

            box.Loaded += (_, _) => box.Focus();
            window.ShowDialog();
            return window._line;
        }

        /// <summary>Read-Host -AsSecureString.</summary>
        public static SecureString AskSecure(Window? owner, string? prompt)
        {
            PromptWindow window = Create(owner, "SECRET REQUESTED",
                string.IsNullOrWhiteSpace(prompt) ? "The running tool is waiting for a secret." : prompt!);

            var box = new PasswordBox();
            window.InputHost.Children.Add(box);

            // SecurePassword hands back a copy, so the secret never becomes a
            // managed string on the way through.
            window.AddConfirmCancel(() => window._secure = box.SecurePassword);

            box.Loaded += (_, _) => box.Focus();
            window.ShowDialog();

            SecureString result = window._secure ?? new SecureString();
            if (!result.IsReadOnly())
            {
                result.MakeReadOnly();
            }
            return result;
        }

        /// <summary>PromptForCredential.</summary>
        public static PSCredential? AskCredential(
            Window? owner, string caption, string message, string userName, string targetName)
        {
            PromptWindow window = Create(owner,
                string.IsNullOrWhiteSpace(caption) ? "CREDENTIAL REQUESTED" : caption,
                string.IsNullOrWhiteSpace(message)
                    ? "The running tool needs credentials" + (string.IsNullOrEmpty(targetName) ? "." : " for " + targetName + ".")
                    : message);

            var user = new TextBox { Text = userName ?? string.Empty, Margin = new Thickness(0, 0, 0, 10) };
            var password = new PasswordBox();

            window.InputHost.Children.Add(Label("USERNAME"));
            window.InputHost.Children.Add(user);
            window.InputHost.Children.Add(Label("PASSWORD"));
            window.InputHost.Children.Add(password);

            window.AddConfirmCancel(() =>
            {
                if (!string.IsNullOrWhiteSpace(user.Text))
                {
                    window._credential = new PSCredential(user.Text, password.SecurePassword);
                }
            });

            user.Loaded += (_, _) => user.Focus();
            window.ShowDialog();
            return window._credential;
        }

        /// <summary>
        /// PromptForChoice. The choices become a button row, which is what the
        /// plan asks for: the answer sits directly beneath the question.
        /// </summary>
        public static int AskChoice(
            Window? owner, string caption, string message, IReadOnlyList<string> choices, int defaultChoice)
        {
            PromptWindow window = Create(owner,
                string.IsNullOrWhiteSpace(caption) ? "CONFIRM" : caption,
                message);

            for (int i = 0; i < choices.Count; i++)
            {
                int index = i;
                var button = new Button
                {
                    // PowerShell marks the accelerator with &; it means nothing here.
                    Content = choices[i].Replace("&", string.Empty),
                    Margin = new Thickness(8, 0, 0, 0),
                    Style = (Style)window.FindResource(i == defaultChoice ? "PrimaryButton" : "GhostButton"),
                };
                button.Click += (_, _) =>
                {
                    window._choice = index;
                    window.Close();
                };
                window.ButtonHost.Children.Add(button);
            }

            window._choice = defaultChoice;
            window.ShowDialog();
            return window._choice;
        }

        private static TextBlock Label(string text) => new TextBlock
        {
            Text = text,
            Style = (Style)Application.Current.FindResource("DimLabel"),
            Margin = new Thickness(0, 0, 0, 5),
        };

        private void AddConfirmCancel(Action onConfirm)
        {
            var cancel = new Button
            {
                Content = "Cancel",
                Style = (Style)FindResource("GhostButton"),
                Margin = new Thickness(8, 0, 0, 0),
            };
            cancel.Click += (_, _) => Close();

            var ok = new Button
            {
                Content = "OK",
                Style = (Style)FindResource("PrimaryButton"),
                Margin = new Thickness(8, 0, 0, 0),
                IsDefault = true,
            };
            ok.Click += (_, _) =>
            {
                onConfirm();
                Close();
            };

            ButtonHost.Children.Add(cancel);
            ButtonHost.Children.Add(ok);
        }
    }
}
