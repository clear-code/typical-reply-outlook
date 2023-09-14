using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace TypicalReply.Config
{
    internal class ConfigLoader
    {
        private string FilePath { get; }
        private string RegistryPath { get; }

        internal ConfigLoader(string filePath, string registryPath)
        {
            FilePath = filePath;
            RegistryPath = registryPath;
        }

        private TypicalReplyConfig LoadFromFile()
        {
            try
            {
                return ConfigSerializer<TypicalReplyConfig>.LoadFromFile(FilePath);
            }
            catch(Exception ex)
            {
                Logger.Log(ex);
                return null;
            }
        }

        private TypicalReplyConfig LoadFromRegistry()
        {
            Dictionary<string, string> registryDict = RegistryConfig.ReadDict(RegistryPath);
            if (registryDict is null)
            {
                return null;
            }
            if (registryDict.Count == 0)
            {
                return null;
            }
            string configString = registryDict[nameof(TypicalReplyConfig)];
            if (string.IsNullOrEmpty(configString))
            {
                return null;
            }
            try
            {
                TypicalReplyConfig conf = ConfigSerializer<TypicalReplyConfig>.Deserialize(configString);
                return conf;
            }
            catch(Exception ex)
            {
                Logger.Log(ex);
                return null;
            }
        }

        private Config GetConfigForCurrentUICulture(TypicalReplyConfig typicalReplyConfig)
        {
            string culture = CultureInfo.CurrentUICulture.Name;
            Config config = typicalReplyConfig?.ConfigList?.FirstOrDefault(_ => _.Culture?.Equals(culture) ?? false);
            if (config is null)
            {
                string lang = culture.Split('-')[0];
                config = typicalReplyConfig?.ConfigList?.FirstOrDefault(_ => _.Culture?.Equals(lang) ?? false);
            }
            if (config is null)
            {
                config = typicalReplyConfig?.ConfigList?.FirstOrDefault();
            }
            return config;
        }

        internal Config Load()
        {
            TypicalReplyConfig registryTypicalReplyConfig = LoadFromRegistry();
            TypicalReplyConfig fileTypicalReplyConfig = LoadFromFile();

            if(registryTypicalReplyConfig is null && fileTypicalReplyConfig is null)
            {
                throw new Exception("No TypicalReplyConfig found");
            }
            var registryConfig = GetConfigForCurrentUICulture(registryTypicalReplyConfig);
            var fileConfig = GetConfigForCurrentUICulture(fileTypicalReplyConfig);
            if (registryConfig is null && fileConfig is null)
            {
                throw new Exception("No Config found");
            }
            if (registryConfig is null)
            {
                Logger.Log("Registy config not found");
                Logger.Log("Config file found");
                return fileConfig;
            }
            if (fileConfig is null)
            {
                Logger.Log("Config file not found");
                Logger.Log("Registy config found");
                return registryConfig;
            }
            return registryConfig.Merge(fileConfig);
        }
    }
}
