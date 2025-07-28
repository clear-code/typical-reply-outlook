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
        private static int MaxGeneration = 10;

        private static long MaxLogSize = 10 * 1024 * 1024;

        private static string LogFileNameBase = "TypicalReply";

        private static string FilePath = Path.Combine(StandardPath.GetUserDir(), "TypicalReply.log");
        private static StreamWriter LogStream { get; set; }

        private static object LockObject = new object();

        internal static void Log(string message) => NoException(() => LogImpl(message));
        internal static void Log(Exception e) => NoException(() => LogImpl(e));

        private static void NoException(Action func)
        {
            try { func(); } catch { }
        }

        private static void LogImpl(string message)
        {
            lock (LockObject)
            {
                RotateIfNeed();
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

        private static void RotateIfNeed()
        {
            if (!File.Exists(FilePath))
            {
                LogStream?.Close();
                LogStream = null;
            }

            if (LogStream is null)
            {
                var fileStream = new FileStream(FilePath, FileMode.OpenOrCreate);
                fileStream.Seek(0, SeekOrigin.End);
                LogStream = new StreamWriter(fileStream);
            }

            var fi = new FileInfo(FilePath);
            if(fi.Length > MaxLogSize)
            {
                Rotate();
            }
        }

        private static void Rotate()
        {
            lock (LockObject)
            {
                LogStream?.Close();

                string previousFileName;
                string previousFilePath;
                string rotatedFileName;
                string rotatedFilePath;
                string userDir = StandardPath.GetUserDir();
                for (int i = MaxGeneration - 1; i >= 0; i--)
                {
                    if (i > 0)
                    {
                        previousFileName = $"{LogFileNameBase}_{i}.log";
                    }
                    else
                    {
                        previousFileName = $"{LogFileNameBase}.log";
                    }

                    previousFilePath = Path.Combine(userDir, previousFileName);

                    if (!File.Exists(previousFilePath))
                    {
                        continue;
                    }
                    rotatedFileName = $"{LogFileNameBase}_{ i + 1 }.log";
                    rotatedFilePath = Path.Combine(userDir, rotatedFileName);

                    File.Copy(previousFilePath, rotatedFilePath, true);
                    File.Delete(previousFilePath);
                }

                LogStream = new StreamWriter(new FileStream(FilePath, FileMode.Create));
            }
        }
    }
}
