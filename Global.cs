using System;
using System.Collections.Generic;
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
        internal static string ConfigFileName { get; } = "TypicalReplyConfig.json";
        internal string UserConfigPath { get; }
        internal TypicalReplyConfig Config { get; }
        #endregion

        #region singletonization
        private static object _lockObject = new object();
        private static Global _instance = null;         

        internal Global()
        {
            UserConfigPath = Path.Combine(StandardPath.GetUserDir(), ConfigFileName);
            Config = ConfigSerializer<TypicalReplyConfig>.LoadFromFile(UserConfigPath);
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
