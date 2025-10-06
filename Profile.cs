using System;

namespace AIHotKey
{
    public class Profile
    {
        public string Name { get; set; }
        public uint Modifiers { get; set; }
        public uint VirtualKey { get; set; }
        public string Prompt { get; set; }
        public int HotkeyId { get; set; }

        public Profile()
        {
            Name = "New Profile";
            Modifiers = 0;
            VirtualKey = 0;
            Prompt = "";
            HotkeyId = -1;
        }

        public Profile(string name, uint modifiers, uint virtualKey, string prompt)
        {
            Name = name;
            Modifiers = modifiers;
            VirtualKey = virtualKey;
            Prompt = prompt;
            HotkeyId = -1;
        }

        public string GetHotkeyString()
        {
            if (Modifiers == 0 && VirtualKey == 0)
                return "Not set";
                
            var keyString = "";
            if ((Modifiers & 0x0002) != 0) keyString += "Ctrl + ";
            if ((Modifiers & 0x0001) != 0) keyString += "Alt + ";
            if ((Modifiers & 0x0004) != 0) keyString += "Shift + ";
            keyString += ((System.Windows.Forms.Keys)VirtualKey).ToString();
            return keyString;
        }
    }
}

