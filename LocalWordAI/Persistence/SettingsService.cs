using Newtonsoft.Json;
using System;
using System.IO;

namespace LocalWordAI.Persistence
{
    public class AppSettings
    {
        public string LmStudioBaseUrl { get; set; } = "http://localhost:1234";
        public string ModelName { get; set; } = "";
        public int TimeoutSeconds { get; set; } = 120;
        public int MaxTokens { get; set; } = 4096;
        public double Temperature { get; set; } = 0.2;
        public bool UseTrackChangesDefault { get; set; } = true;
        public bool ShowConfirmBeforeApply { get; set; } = true;
        public bool LogTechnicalOnly { get; set; } = true;
        public bool AutoDetectSelection { get; set; } = true;
        public string SkillsDirectory { get; set; } = "";
        public string AuthorName { get; set; } = "Local AI";
        public int ChunkSizeWords { get; set; } = 2000;
    }

    public class SettingsService
    {
        private static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "LocalWordAI", "settings.json");

        public AppSettings Current { get; private set; } = new AppSettings();

        public void Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var json = File.ReadAllText(SettingsPath);
                    Current = JsonConvert.DeserializeObject<AppSettings>(json) ?? new AppSettings();
                }
                else
                {
                    Current = new AppSettings();
                    Save();
                }

                if (string.IsNullOrEmpty(Current.SkillsDirectory))
                    Current.SkillsDirectory = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "LocalWordAI", "Skills");
            }
            catch
            {
                Current = new AppSettings();
            }
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath));
                File.WriteAllText(SettingsPath, JsonConvert.SerializeObject(Current, Formatting.Indented));
            }
            catch { }
        }

        public bool IsValidLocalUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return false;
            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri)) return false;
            var host = uri.Host.ToLower();
            return host == "localhost" || host == "127.0.0.1" || host == "::1";
        }
    }
}
