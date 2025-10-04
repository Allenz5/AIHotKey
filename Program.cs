using System;
using System.Diagnostics;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using OpenAI.Chat;

namespace AIHotKey
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        const int WM_HOTKEY = 0x0312;
        const int HOTKEY_ID = 1;
        const uint MOD_CONTROL = 0x0002;
        const uint VK_OEM_PLUS = 0xBB;

        private TextBox txtKey;
        private Button btnSave;
        private CheckBox chkShow;
        private CheckBox chkStartup;
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;

        private const string AppName = "AIHotKey";
        private const string RunRegPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string EnvVarName = "OPENAI_API_KEY";

        public MainForm()
        {
            Text = AppName;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = false;
            Size = new Size(540, 180);
            StartPosition = FormStartPosition.Manual;
            TopMost = true;

            var area = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(area.Right - Width - 10, area.Bottom - Height - 10);

            // UI
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 3,
                Padding = new Padding(10)
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));

            var lbl = new Label { Text = "OpenAI API Key:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };

            txtKey = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true, PlaceholderText = "sk-..." };
            btnSave = new Button { Text = "Save", Dock = DockStyle.Fill };
            btnSave.Click += (s, e) => SaveKey();

            chkShow = new CheckBox { Text = "Show", Dock = DockStyle.Left };
            chkShow.CheckedChanged += (s, e) => txtKey.UseSystemPasswordChar = !chkShow.Checked;

            chkStartup = new CheckBox { Text = "Run at startup", Dock = DockStyle.Left };
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
                    chkStartup.Checked = !chkStartup.Checked; // revert
                    chkStartup.CheckedChanged += (s2, e2) => { };
                }
            };

            panel.Controls.Add(lbl, 0, 0);
            panel.Controls.Add(txtKey, 1, 0);
            panel.Controls.Add(btnSave, 2, 0);

            panel.Controls.Add(new Label
            {
                Text = "Hotkey: Ctrl + '+' to rewrite selected text and paste",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            }, 0, 1);
            panel.SetColumnSpan(panel.GetControlFromPosition(0, 1), 2);
            panel.Controls.Add(chkShow, 2, 1);

            panel.Controls.Add(chkStartup, 0, 2);
            panel.SetColumnSpan(chkStartup, 3);

            Controls.Add(panel);

            // Hotkey
            if (!RegisterHotKey(Handle, HOTKEY_ID, MOD_CONTROL, VK_OEM_PLUS))
                MessageBox.Show("Failed to register hotkey.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            FormClosed += (s, e) => UnregisterHotKey(Handle, HOTKEY_ID);

            // Tray
            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("Show Window", null, (s, e) => ShowFromTray());
            trayMenu.Items.Add("Exit", null, (s, e) => ExitApp());

            trayIcon = new NotifyIcon
            {
                Text = AppName,
                Icon = SystemIcons.Information,
                Visible = true,
                ContextMenuStrip = trayMenu
            };
            trayIcon.DoubleClick += (s, e) => ShowFromTray();

            // Prefill env var + startup state
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

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                _ = OnHotkeyAsync();
                return;
            }
            base.WndProc(ref m);
        }

        // Close button -> minimize to tray
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
                trayIcon.ShowBalloonTip(1500, AppName, "Still running in the tray. Right-click to Exit.", ToolTipIcon.Info);
            }
            else
            {
                base.OnFormClosing(e);
            }
        }

        private void ShowFromTray()
        {
            Show();
            WindowState = FormWindowState.Normal;
            Activate();
        }

        private void ExitApp()
        {
            trayIcon.Visible = false;
            Application.Exit();
        }

        private async Task OnHotkeyAsync()
        {
            // Copy selection
            SendKeys.SendWait("^c");
            await Task.Delay(120);

            string orig = "";
            try { if (Clipboard.ContainsText()) orig = Clipboard.GetText(); } catch { }

            if (string.IsNullOrWhiteSpace(orig))
            {
                System.Media.SystemSounds.Beep.Play();
                return;
            }

            var key = GetCurrentKey();
            if (string.IsNullOrWhiteSpace(key))
            {
                MessageBox.Show("Please enter and save your OpenAI API Key first.", "Missing API Key",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                ChatClient client = new(model: "gpt-4o", apiKey: key);
                ChatCompletion completion = client.CompleteChat(
                    "Please make the following paragraph smoother and grammatically correct; " +
                    "return only the plain revised text without quotes and nothing else:\n\n\"" + orig + "\"");
                string revised = completion.Content[0].Text;
                await TryPasteTextAsync(revised, orig);
            }
            catch (Exception ex)
            {
                MessageBox.Show("[OpenAI error] " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task TryPasteTextAsync(string textToPaste, string previousClipboard)
        {
            try
            {
                Clipboard.SetText(textToPaste);
                await Task.Delay(80);
                SendKeys.SendWait("^v");
            }
            catch { }
            finally
            {
                try
                {
                    if (!string.IsNullOrEmpty(previousClipboard))
                        Clipboard.SetText(previousClipboard);
                    else
                        Clipboard.Clear();
                }
                catch { }
            }
        }

        private string GetCurrentKey()
        {
            var key = txtKey?.Text?.Trim();
            if (!string.IsNullOrEmpty(key))
            {
                Environment.SetEnvironmentVariable(EnvVarName, key, EnvironmentVariableTarget.Process);
                return key;
            }

            key = Environment.GetEnvironmentVariable(EnvVarName, EnvironmentVariableTarget.Process)
                  ?? Environment.GetEnvironmentVariable(EnvVarName, EnvironmentVariableTarget.User);
            return key?.Trim();
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

        // ===== Startup helpers =====
        private bool IsRunAtStartupEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunRegPath, writable: false);
                if (key == null) return false;
                var val = key.GetValue(AppName) as string;
                if (string.IsNullOrEmpty(val)) return false;

                // Some installers wrap with quotes; normalize both forms.
                var exe = $"\"{Application.ExecutablePath}\"";
                return string.Equals(val, exe, StringComparison.OrdinalIgnoreCase)
                       || string.Equals(val, Application.ExecutablePath, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private void SetRunAtStartup(bool enable)
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
