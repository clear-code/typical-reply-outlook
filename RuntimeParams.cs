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
    internal class RuntimeParams
    {
        #region params
        internal Config.Config Config { get; }
        #endregion

        #region singletonization
        private static object _lockObject = new object();
        private static RuntimeParams _instance = null;         

        internal RuntimeParams()
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

        internal static RuntimeParams GetInstance()
        {
            lock (_lockObject)
            {
                _instance = _instance ?? new RuntimeParams();
            }
            return _instance;
        }
        #endregion
    }
}
