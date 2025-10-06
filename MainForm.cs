using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenAI.Chat;

namespace AIHotKey
{
    public class MainForm : Form
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        const int WM_HOTKEY = 0x0312;

        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private Panel contentPanel;
        private MenuStrip menuStrip;
        private List<Profile> profiles;
        private Dictionary<int, Profile> hotkeyProfileMap;
        private int nextHotkeyId = 1;
        private Profile? currentProfile;

        private const string AppName = "AIHotKey";
        private const string EnvVarName = "OPENAI_API_KEY";

        public MainForm()
        {
            Text = AppName;
            FormBorderStyle = FormBorderStyle.Sizable;
            Size = new Size(600, 500);
            StartPosition = FormStartPosition.CenterScreen;

            profiles = ProfileManager.LoadProfiles();
            hotkeyProfileMap = new Dictionary<int, Profile>();

            // Menu Bar
            menuStrip = new MenuStrip();
            
            // Settings button (always first)
            var settingsMenu = new ToolStripMenuItem("Settings");
            settingsMenu.Click += (s, e) => OpenSettings();
            menuStrip.Items.Add(settingsMenu);
            
            // Add Profile button (after Settings)
            var addProfileMenu = new ToolStripMenuItem("+ Add Profile");
            addProfileMenu.Click += (s, e) => AddNewProfile();
            menuStrip.Items.Add(addProfileMenu);
            
            // Profile buttons will be added after these
            
            MainMenuStrip = menuStrip;
            Controls.Add(menuStrip);

            // Content panel (will display selected profile)
            contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15, 35, 15, 15) // Add top padding to avoid overlap with menu
            };
            Controls.Add(contentPanel);

            // Add buttons for existing profiles
            RebuildProfileButtons();

            // Show first profile by default
            if (profiles.Count > 0)
            {
                ShowProfile(profiles[0]);
            }

            // Register all hotkeys
            Load += (s, e) => RegisterAllHotkeys();
            FormClosed += (s, e) =>
            {
                UnregisterAllHotkeys();
                ProfileManager.SaveProfiles(profiles);
            };

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
        }

        private void RebuildProfileButtons()
        {
            // Remove all profile buttons (but keep Settings and Add Profile)
            var itemsToRemove = menuStrip.Items.Cast<ToolStripItem>()
                .Where(item => item.Tag is Profile)
                .ToList();
            
            foreach (var item in itemsToRemove)
            {
                menuStrip.Items.Remove(item);
            }

            // Add profile buttons after "Add Profile" button (index 2, after Settings and Add Profile)
            for (int i = 0; i < profiles.Count; i++)
            {
                var profile = profiles[i];
                var profileButton = new ToolStripMenuItem(profile.Name);
                profileButton.Tag = profile;
                profileButton.Click += (s, e) => ShowProfile(profile);
                menuStrip.Items.Add(profileButton);
            }
        }

        private void AddNewProfile()
        {
            int profileNumber = profiles.Count + 1;
            var newProfile = new Profile
            {
                Name = $"Hotkey {profileNumber}",
                Modifiers = 0,
                VirtualKey = 0,
                Prompt = ""
            };

            profiles.Add(newProfile);
            RebuildProfileButtons();
            ShowProfile(newProfile);
            SaveProfiles();
        }

        private void ShowProfile(Profile profile)
        {
            currentProfile = profile;
            contentPanel.Controls.Clear();

            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4
            };
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 80)); // Hotkey
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // Prompt
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30)); // Info
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // Delete button

            // Hotkey Configuration
            var hotkeyGroup = new GroupBox
            {
                Text = $"Hotkey Configuration - {profile.Name}",
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };
            var hotkeyPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1
            };
            hotkeyPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            hotkeyPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            var lblHotkey = new Label
            {
                Text = "Press keys:",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            var txtHotkey = new TextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Text = profile.GetHotkeyString(),
                PlaceholderText = "Click here and press your desired hotkey combination"
            };
            txtHotkey.KeyDown += (s, e) => TxtHotkey_KeyDown(profile, txtHotkey, e);

            hotkeyPanel.Controls.Add(lblHotkey, 0, 0);
            hotkeyPanel.Controls.Add(txtHotkey, 1, 0);
            hotkeyGroup.Controls.Add(hotkeyPanel);

            // Prompt Configuration
            var promptGroup = new GroupBox
            {
                Text = "AI Prompt Configuration",
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };
            var promptPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2
            };
            promptPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            promptPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var lblPrompt = new Label
            {
                Text = "Custom prompt (the selected text will be appended):",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            var txtPrompt = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Text = profile.Prompt,
                PlaceholderText = "Enter your custom prompt here..."
            };
            txtPrompt.TextChanged += (s, e) =>
            {
                profile.Prompt = txtPrompt.Text;
                SaveProfiles();
            };

            promptPanel.Controls.Add(lblPrompt, 0, 0);
            promptPanel.Controls.Add(txtPrompt, 0, 1);
            promptGroup.Controls.Add(promptPanel);

            // Info label
            var lblInfo = new Label
            {
                Text = "Note: Changes are saved automatically.",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.Gray,
                Font = new Font(Font.FontFamily, 8)
            };

            // Delete button
            var btnDelete = new Button
            {
                Text = "Delete Profile",
                Dock = DockStyle.Right,
                Width = 120,
                BackColor = Color.IndianRed,
                ForeColor = Color.White
            };
            btnDelete.Click += (s, e) => DeleteProfile(profile);

            mainPanel.Controls.Add(hotkeyGroup, 0, 0);
            mainPanel.Controls.Add(promptGroup, 0, 1);
            mainPanel.Controls.Add(lblInfo, 0, 2);
            mainPanel.Controls.Add(btnDelete, 0, 3);

            contentPanel.Controls.Add(mainPanel);
        }

        private void DeleteProfile(Profile profile)
        {
            if (profiles.Count <= 1)
            {
                MessageBox.Show("Cannot delete the last profile. At least one profile is required.", 
                    "Cannot Delete", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var result = MessageBox.Show($"Are you sure you want to delete '{profile.Name}'?", 
                "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                // Unregister hotkey if registered
                if (profile.HotkeyId > 0)
                {
                    UnregisterHotKey(Handle, profile.HotkeyId);
                    hotkeyProfileMap.Remove(profile.HotkeyId);
                }

                profiles.Remove(profile);
                
                // Renumber remaining profiles
                for (int i = 0; i < profiles.Count; i++)
                {
                    profiles[i].Name = $"Hotkey {i + 1}";
                }
                
                RebuildProfileButtons();
                SaveProfiles();

                // Show first profile
                if (profiles.Count > 0)
                {
                    ShowProfile(profiles[0]);
                }
            }
        }

        private void TxtHotkey_KeyDown(Profile profile, TextBox txtHotkey, KeyEventArgs e)
        {
            e.SuppressKeyPress = true;
            e.Handled = true;

            // Build modifiers
            uint modifiers = 0;
            if (e.Control) modifiers |= 0x0002; // MOD_CONTROL
            if (e.Alt) modifiers |= 0x0001; // MOD_ALT
            if (e.Shift) modifiers |= 0x0004; // MOD_SHIFT

            // Get the actual key (not modifier keys)
            var key = e.KeyCode;
            if (key == Keys.ControlKey || key == Keys.ShiftKey || key == Keys.Menu) // Menu is Alt
                return;

            if (key == Keys.None)
                return;

            // Check if this hotkey is already used by another profile
            foreach (var p in profiles)
            {
                if (p != profile && p.Modifiers == modifiers && p.VirtualKey == (uint)key)
                {
                    MessageBox.Show("This hotkey is already used by another profile. Please choose a different one.", 
                        "Hotkey Conflict", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // Unregister old hotkey if it exists
            if (profile.HotkeyId > 0)
            {
                UnregisterHotKey(Handle, profile.HotkeyId);
                hotkeyProfileMap.Remove(profile.HotkeyId);
            }

            // Try to register new hotkey
            uint vk = (uint)key;
            int newHotkeyId = nextHotkeyId++;
            
            if (RegisterHotKey(Handle, newHotkeyId, modifiers, vk))
            {
                // Store old values for potential revert
                uint oldModifiers = profile.Modifiers;
                uint oldVK = profile.VirtualKey;
                int oldHotkeyId = profile.HotkeyId;
                
                profile.Modifiers = modifiers;
                profile.VirtualKey = vk;
                profile.HotkeyId = newHotkeyId;
                hotkeyProfileMap[newHotkeyId] = profile;
                
                // Display the hotkey
                var keyString = "";
                if (e.Control) keyString += "Ctrl + ";
                if (e.Alt) keyString += "Alt + ";
                if (e.Shift) keyString += "Shift + ";
                keyString += key.ToString();
                
                txtHotkey.Text = keyString;
                SaveProfiles();
                
                MessageBox.Show($"Hotkey successfully set to: {keyString}", "Success", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Failed to register hotkey. It may be in use by another application.", 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
                // Reregister old hotkey if there was one
                if (profile.HotkeyId > 0 && profile.Modifiers != 0 && profile.VirtualKey != 0)
                {
                    RegisterHotKey(Handle, profile.HotkeyId, profile.Modifiers, profile.VirtualKey);
                    hotkeyProfileMap[profile.HotkeyId] = profile;
                }
            }
        }

        private void RegisterAllHotkeys()
        {
            hotkeyProfileMap.Clear();
            
            foreach (var profile in profiles)
            {
                // Skip profiles with no hotkey set
                if (profile.Modifiers == 0 && profile.VirtualKey == 0)
                    continue;
                    
                int hotkeyId = nextHotkeyId++;
                if (RegisterHotKey(Handle, hotkeyId, profile.Modifiers, profile.VirtualKey))
                {
                    profile.HotkeyId = hotkeyId;
                    hotkeyProfileMap[hotkeyId] = profile;
                }
                else
                {
                    MessageBox.Show($"Failed to register hotkey for '{profile.Name}'. It may be in use.", 
                        "Hotkey Registration Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void UnregisterAllHotkeys()
        {
            foreach (var profile in profiles)
            {
                if (profile.HotkeyId > 0)
                {
                    UnregisterHotKey(Handle, profile.HotkeyId);
                }
            }
            hotkeyProfileMap.Clear();
        }

        private void SaveProfiles()
        {
            ProfileManager.SaveProfiles(profiles);
        }

        private void OpenSettings()
        {
            using var settingsForm = new SettingsForm();
            settingsForm.ShowDialog(this);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
            {
                int hotkeyId = m.WParam.ToInt32();
                if (hotkeyProfileMap.TryGetValue(hotkeyId, out Profile? profile))
                {
                    _ = OnHotkeyAsync(profile);
                    return;
                }
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

        private async Task OnHotkeyAsync(Profile profile)
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
                ChatClient client = new(model: "gpt-4o-mini", apiKey: key);
                
                // Use custom prompt with the selected text
                var fullPrompt = profile.Prompt + "\n\n\"" + orig + "\"";
                
                ChatCompletion completion = client.CompleteChat(fullPrompt);
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

        private string? GetCurrentKey()
        {
            var key = Environment.GetEnvironmentVariable(EnvVarName, EnvironmentVariableTarget.Process)
                  ?? Environment.GetEnvironmentVariable(EnvVarName, EnvironmentVariableTarget.User);
            return key?.Trim();
        }
    }
}
