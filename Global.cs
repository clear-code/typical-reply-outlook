using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TypicalReply.Config;

namespace TypicalReply
{
    internal class Global
    {
        #region params
        internal static string ConfigFileName { get; } = "TypicalReplyConfig.json";
        internal string UserConfigPath { get; }
        #endregion

        #region singletonization
        private static object _lockObject = new object();
        private static Global _instance = null;         

        internal Global()
        {
            UserConfigPath = Path.Combine(StandardPath.GetUserDir(), ConfigFileName);
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
