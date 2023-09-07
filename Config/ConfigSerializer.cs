using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.IO;

namespace TypicalReply.Config
{
    internal static class ConfigSerializer<T>
    {
        internal static T LoadFromFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return default;
            }
            var jsonString = File.ReadAllText(filePath, Encoding.UTF8);
            return JsonSerializer.Deserialize<T>(jsonString);
        }

        internal static void SaveAsFile(T obj, string filePath)
        {
            if (!File.Exists(filePath))
            {
                string directory = Path.GetDirectoryName(filePath);
                Directory.CreateDirectory(directory);
                File.Create(filePath);
            }
            var jsonString = JsonSerializer.Serialize(obj);
            File.WriteAllText(filePath, jsonString);
        }
    }
}
