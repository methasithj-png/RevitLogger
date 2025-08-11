using System;
using System.IO;
using System.Text;
using System.Text.Json;

namespace RevitProjectCloseLogger
{
    internal static class SettingsManager
    {
        private const string SettingsFileName = "settings.json";
        private const string AppFolderName = "RevitProjectCloseLogger";

        private class Settings
        {
            public bool ExportEnabled { get; set; } = true;
        }

        private static string GetAppFolder()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var path = Path.Combine(appData, AppFolderName);
            Directory.CreateDirectory(path);
            return path;
        }

        private static string GetSettingsPath()
        {
            return Path.Combine(GetAppFolder(), SettingsFileName);
        }

        public static bool IsExportEnabled()
        {
            try
            {
                var path = GetSettingsPath();
                if (!File.Exists(path)) return true; // default enabled
                var json = File.ReadAllText(path, Encoding.UTF8);
                var settings = JsonSerializer.Deserialize<Settings>(json);
                return settings?.ExportEnabled ?? true;
            }
            catch
            {
                return true;
            }
        }

        public static void SetExportEnabled(bool enabled)
        {
            try
            {
                var settings = new Settings { ExportEnabled = enabled };
                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(GetSettingsPath(), json, Encoding.UTF8);
            }
            catch
            {
                // ignore
            }
        }

        public static string GetLogsFolder()
        {
            var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var folder = Path.Combine(docs, AppFolderName, "Logs");
            Directory.CreateDirectory(folder);
            return folder;
        }
    }
}