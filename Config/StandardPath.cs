using System;
using System.IO;

namespace TypicalReply.Config
{
    public class StandardPath
    {
        public static string GetUserDir()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, "TypicalReply");
        }

        public static string GetDefaultConfigDir()
        {
            return Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "DefaultConfig");
        }
    }
}
