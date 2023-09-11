using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TypicalReply.Config;

namespace TypicalReply
{
    internal static class Logger
    {
        private static StreamWriter LogStream { get; } = new StreamWriter(new FileStream(Path.Combine(StandardPath.GetUserDir(), "TypicalReply.log"), FileMode.Create));

        internal static void Log(string message) => NoException(() => LogImpl(message));
        internal static void Log(Exception e) => NoException(() => LogImpl(e));

        private static void NoException(Action func)
        {
            try { func(); } catch { }
        }

        private static void LogImpl(string message)
        {
            lock (LogStream)
            {
                LogStream.WriteLine($"{GetTimestamp()} : {message}");
                LogStream.Flush();
            }
        }

        private static void LogImpl(Exception e)
        {
            LogImpl(e.ToString());
        }

        private static string GetTimestamp()
        {
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }
    }
}
