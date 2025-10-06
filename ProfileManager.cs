using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace AIHotKey
{
    public static class ProfileManager
    {
        private static readonly string ProfilesFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AIHotKey",
            "profiles.json"
        );

        public static List<Profile> LoadProfiles()
        {
            try
            {
                if (File.Exists(ProfilesFilePath))
                {
                    string json = File.ReadAllText(ProfilesFilePath);
                    var profiles = JsonSerializer.Deserialize<List<Profile>>(json);
                    if (profiles != null && profiles.Count > 0)
                        return profiles;
                }
            }
            catch { }

            return CreateDefaultProfiles();
        }

        public static void SaveProfiles(List<Profile> profiles)
        {
            try
            {
                string? directory = Path.GetDirectoryName(ProfilesFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                string json = JsonSerializer.Serialize(profiles, new JsonSerializerOptions
                {
                    WriteIndented = true
                });
                File.WriteAllText(ProfilesFilePath, json);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Failed to save profiles:\n{ex.Message}", 
                    "Error", System.Windows.Forms.MessageBoxButtons.OK, 
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private static List<Profile> CreateDefaultProfiles()
        {
            return new List<Profile>
            {
                new Profile
                {
                    Name = "Hotkey 1",
                    Modifiers = 0x0002, // MOD_CONTROL
                    VirtualKey = 0xBB, // VK_OEM_PLUS (Ctrl + +)
                    Prompt = "Please make the following paragraph smoother and grammatically correct; return only the plain revised text without quotes:\n\n"
                }
            };
        }
    }
}

