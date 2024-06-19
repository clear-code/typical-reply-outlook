using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.IO;

namespace TypicalReply.Config
{
    internal class ConfigLoader
    {
        private string UserFilePath { get; }
        private string DefaultFilePath { get; }
        private string RegistryPath { get; }

        internal ConfigLoader()
        {
            UserFilePath = Path.Combine(StandardPath.GetUserDir(), Const.Config.FileName);
            DefaultFilePath = Path.Combine(StandardPath.GetDefaultConfigDir(), Const.Config.FileName);
            Logger.Log($"Default config file path: {DefaultFilePath}");
            RegistryPath = Const.RegistryPath.DefaultPolicy;
        }

        private TypicalReplyConfig LoadFromFile(string filePath)
        {
            try
            {
                return ConfigSerializer<TypicalReplyConfig>.LoadFromFile(filePath);
            }
            catch (Exception ex)
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
            catch (Exception ex)
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

        private (Config config, int priority) DetemineConfigWithPriority((Config config, int priority) left, (Config config, int priority) right)
        {
            if (left.config == null)
            {
                return right;
            }
            else if (right.config == null)
            {
                return left;
            }
            else if (left.priority == right.priority)
            {
                return (left.config.Merge(right.config), left.priority);
            }
            else if (left.priority > right.priority)
            {
                return left;
            }
            else
            {
                return right;
            }
        }

        internal Config Load()
        {
            TypicalReplyConfig registryTypicalReplyConfig = LoadFromRegistry();
            TypicalReplyConfig defaultFileTypicalReplyConfig = LoadFromFile(DefaultFilePath);
            TypicalReplyConfig userFileTypicalReplyConfig = LoadFromFile(UserFilePath);

            int registryPriority = registryTypicalReplyConfig?.Priority ?? -1;
            int defaultFilePriority = defaultFileTypicalReplyConfig?.Priority ?? -1;
            int userFilePriority = userFileTypicalReplyConfig?.Priority ?? -1;

            if (registryTypicalReplyConfig is null && userFileTypicalReplyConfig is null && defaultFileTypicalReplyConfig is null)
            {
                throw new Exception("No TypicalReplyConfig found");
            }

            var registryConfig = GetConfigForCurrentUICulture(registryTypicalReplyConfig);
            var defaultFileConfig = GetConfigForCurrentUICulture(defaultFileTypicalReplyConfig);
            var userFileConfig = GetConfigForCurrentUICulture(userFileTypicalReplyConfig);

            if (registryConfig is null && userFileConfig is null && defaultFileConfig is null)
            {
                throw new Exception("No Config found");
            }
            if (registryConfig is null)
            {
                Logger.Log("Registy config not found");
            }
            if (defaultFileConfig is null)
            {
                Logger.Log("Default config file not found");
            }
            if (userFileConfig is null)
            {
                Logger.Log("User Config file not found");
            }

            (Config config, int priority) configAndPriority = DetemineConfigWithPriority((registryConfig, registryPriority), (defaultFileConfig, defaultFilePriority));
            configAndPriority = DetemineConfigWithPriority(configAndPriority, (userFileConfig, userFilePriority));
            return configAndPriority.config;
        }
    }
}
