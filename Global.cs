using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using TypicalReply.Config;

namespace TypicalReply
{
    internal class Global
    {
        #region params
        private static string ConfigFileName { get; } = "TypicalReplyConfig.json";
        internal Config.Config Config { get; }
        #endregion

        #region singletonization
        private static object _lockObject = new object();
        private static Global _instance = null;         

        internal Global()
        {
            string userConfigPath = Path.Combine(StandardPath.GetUserDir(), ConfigFileName);
            try
            {
                TypicalReplyConfig typicalReplyConfig = ConfigSerializer<TypicalReplyConfig>.LoadFromFile(userConfigPath);
                if (typicalReplyConfig?.ConfigList is null || !typicalReplyConfig.ConfigList.Any())
                {
                    throw new Exception("Invalid config file.");
                }
                string culture = CultureInfo.CurrentUICulture.Name;
                Config = typicalReplyConfig.ConfigList.FirstOrDefault(_ => _.Culture?.Equals(culture) ?? false);
                if (Config is null)
                {
                    string lang = culture.Split('-')[0];
                    Config = typicalReplyConfig.ConfigList.FirstOrDefault(_ => _.Culture?.Equals(lang) ?? false);
                }
                if (Config is null)
                {
                    Config = typicalReplyConfig.ConfigList.FirstOrDefault();
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
                throw;
            }

        }

        internal static Global GetInstance()
        {
            lock (_lockObject)
            {
                _instance = _instance ?? new Global();
            }
            return _instance;
        }
        #endregion
    }
}
