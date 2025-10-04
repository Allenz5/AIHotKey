using System.Configuration;

namespace CopyToastSimple.Properties
{
    internal sealed partial class Settings : ApplicationSettingsBase
    {
        private static Settings defaultInstance = ((Settings)(ApplicationSettingsBase.Synchronized(new Settings())));

        public static Settings Default
        {
            get
            {
                return defaultInstance;
            }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("")]
        public string OpenAI_ApiKey
        {
            get
            {
                return ((string)(this["OpenAI_ApiKey"]));
            }
            set
            {
                this["OpenAI_ApiKey"] = value;
            }
        }
    }
}
