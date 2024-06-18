using System;
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
                var configLoader = new ConfigLoader();
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
