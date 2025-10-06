using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;

namespace AIHotKey
{
    public class SettingsForm : Form
    {
        private TextBox txtKey;
        private Button btnSave;
        private CheckBox chkShow;
        private CheckBox chkStartup;

        private const string EnvVarName = "OPENAI_API_KEY";
        private const string AppName = "AIHotKey";
        private const string RunRegPath = @"Software\Microsoft\Windows\CurrentVersion\Run";

        public SettingsForm()
        {
            Text = "Settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Size = new Size(550, 220);
            StartPosition = FormStartPosition.CenterParent;

            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 3,
                Padding = new Padding(20),
                AutoSize = true
            };
            
            // Column styles
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
            
            // Row styles
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));

            // Row 0: API Key
            var lblKey = new Label 
            { 
                Text = "OpenAI API Key:", 
                Dock = DockStyle.Fill, 
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = false
            };

            txtKey = new TextBox 
            { 
                Dock = DockStyle.Fill,
                UseSystemPasswordChar = true, 
                PlaceholderText = "sk-...",
                Anchor = AnchorStyles.Left | AnchorStyles.Right
            };
            
            btnSave = new Button 
            { 
                Text = "Save", 
                Dock = DockStyle.Fill,
                UseVisualStyleBackColor = true
            };
            btnSave.Click += (s, e) => SaveKey();

            panel.Controls.Add(lblKey, 0, 0);
            panel.Controls.Add(txtKey, 1, 0);
            panel.Controls.Add(btnSave, 2, 0);

            // Row 1: Show Password
            var lblShow = new Label 
            { 
                Text = "Show password:", 
                Dock = DockStyle.Fill, 
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = false
            };
            
            chkShow = new CheckBox 
            { 
                Text = "Show", 
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = false
            };
            chkShow.CheckedChanged += (s, e) => txtKey.UseSystemPasswordChar = !chkShow.Checked;

            panel.Controls.Add(lblShow, 0, 1);
            panel.Controls.Add(chkShow, 1, 1);

            // Row 2: Run at Startup
            chkStartup = new CheckBox 
            { 
                Text = "Run at startup", 
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = false
            };
            chkStartup.CheckedChanged += (s, e) =>
            {
                try
                {
                    SetRunAtStartup(chkStartup.Checked);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to update startup setting:\n" + ex.Message, "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    chkStartup.CheckedChanged -= (s2, e2) => { };
                    chkStartup.Checked = !chkStartup.Checked;
                    chkStartup.CheckedChanged += (s2, e2) => { };
                }
            };

            panel.Controls.Add(chkStartup, 0, 2);
            panel.SetColumnSpan(chkStartup, 3);

            Controls.Add(panel);

            // Load current values
            LoadCurrentSettings();
        }

        private void LoadCurrentSettings()
        {
            try
            {
                var existing = Environment.GetEnvironmentVariable(EnvVarName, EnvironmentVariableTarget.User)
                               ?? Environment.GetEnvironmentVariable(EnvVarName, EnvironmentVariableTarget.Process);
                if (!string.IsNullOrWhiteSpace(existing))
                    txtKey.Text = existing;

                chkStartup.Checked = IsRunAtStartupEnabled();
            }
            catch { }
        }

        private void SaveKey()
        {
            var key = txtKey.Text?.Trim();
            if (string.IsNullOrEmpty(key))
            {
                MessageBox.Show("Please enter a valid OpenAI API Key.", "Info",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Immediate availability
                Environment.SetEnvironmentVariable(EnvVarName, key, EnvironmentVariableTarget.Process);
                Environment.SetEnvironmentVariable(EnvVarName, key, EnvironmentVariableTarget.User);

                // Persist for future processes
                var psi = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c setx {EnvVarName} \"{key.Replace("\"", "\\\"")}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                };
                using var p = Process.Start(psi);
                p?.WaitForExit();

                MessageBox.Show("API Key saved.\nNewly started apps will see it; already running apps may need a restart.",
                    "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save API Key:\n" + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static bool IsRunAtStartupEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunRegPath, writable: false);
                if (key == null) return false;
                var val = key.GetValue(AppName) as string;
                if (string.IsNullOrEmpty(val)) return false;

                var exe = $"\"{Application.ExecutablePath}\"";
                return string.Equals(val, exe, StringComparison.OrdinalIgnoreCase)
                       || string.Equals(val, Application.ExecutablePath, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static void SetRunAtStartup(bool enable)
        {
            using var key = Registry.CurrentUser.OpenSubKey(RunRegPath, writable: true)
                           ?? Registry.CurrentUser.CreateSubKey(RunRegPath, true);

            if (enable)
            {
                var exe = $"\"{Application.ExecutablePath}\"";
                key.SetValue(AppName, exe);
            }
            else
            {
                if (key.GetValue(AppName) != null)
                    key.DeleteValue(AppName, false);
            }
        }
    }
}
