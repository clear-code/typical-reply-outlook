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
        internal Config.Config Config { get; }
        #endregion

        #region singletonization
        private static object _lockObject = new object();
        private static Global _instance = null;         

        internal Global()
        {
            try
            {
                string configFilePath = Path.Combine(StandardPath.GetUserDir(), Const.Config.FileName);
                var configLoader = new ConfigLoader(configFilePath, Const.RegistryPath.DefaultPolicy);
                this.Config = configLoader.Load();
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
