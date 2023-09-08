using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.IO;
using System.Text.Json.Serialization;

namespace TypicalReply.Config
{
    internal static class ConfigSerializer<T>
    {
        private static readonly JsonSerializerOptions _options = new JsonSerializerOptions
        {
            Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },
            PropertyNameCaseInsensitive = true,
        };

        internal static T LoadFromFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return default;
            }
            var jsonString = File.ReadAllText(filePath, Encoding.UTF8);
            return JsonSerializer.Deserialize<T>(jsonString, _options);
        }

        internal static void SaveAsFile(T obj, string filePath)
        {
            string jsonString = JsonSerializer.Serialize(obj, _options);
            string tmpfile = Path.GetTempFileName();
            File.WriteAllText(tmpfile, jsonString);
            string directory = Path.GetDirectoryName(filePath);
            Directory.CreateDirectory(directory);
            File.Move(tmpfile, filePath);
        }
    }
}
