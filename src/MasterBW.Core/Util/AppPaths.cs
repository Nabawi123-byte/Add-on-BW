using System;
using System.IO;
using System.Text;
using Newtonsoft.Json;

namespace MasterBW.Core.Util
{
    /// <summary>Lokasi folder data MASTER BW (per pengguna, tidak butuh hak admin).</summary>
    public static class AppPaths
    {
        public static string DataRoot
        {
            get
            {
                var overrideDir = Environment.GetEnvironmentVariable("MASTERBW_HOME");
                var root = !string.IsNullOrWhiteSpace(overrideDir)
                    ? overrideDir
                    : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MasterBW");
                Directory.CreateDirectory(root);
                return root;
            }
        }

        public static string SettingsFile => Path.Combine(DataRoot, "settings.json");
        public static string EngineEndpointFile => Path.Combine(DataRoot, "engine.json");
        public static string LogDir => Ensure(Path.Combine(DataRoot, "logs"));
        public static string CacheDir => Ensure(Path.Combine(DataRoot, "cache"));
        public static string HistoryDir => Ensure(Path.Combine(DataRoot, "history"));
        public static string UserTemplatesDir => Ensure(Path.Combine(DataRoot, "templates"));
        public static string ReferencesFile => Path.Combine(DataRoot, "referensi.json");

        /// <summary>Folder instalasi (tempat MasterBW.Engine.exe dan folder models).</summary>
        public static string InstallDir
        {
            get
            {
                var env = Environment.GetEnvironmentVariable("MASTERBW_INSTALL");
                if (!string.IsNullOrWhiteSpace(env)) return env;
                var local = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "MasterBW");
                return local;
            }
        }

        public static string ModelsDir => Path.Combine(InstallDir, "models");

        private static string Ensure(string dir)
        {
            Directory.CreateDirectory(dir);
            return dir;
        }

        public static T ReadJson<T>(string path) where T : class
        {
            try
            {
                if (!File.Exists(path)) return null;
                return JsonConvert.DeserializeObject<T>(File.ReadAllText(path, Encoding.UTF8));
            }
            catch
            {
                return null;
            }
        }

        public static void WriteJson(string path, object value)
        {
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            var tmp = path + ".tmp";
            File.WriteAllText(tmp, JsonConvert.SerializeObject(value, Formatting.Indented), Encoding.UTF8);
            if (File.Exists(path)) File.Delete(path);
            File.Move(tmp, path);
        }
    }

    /// <summary>Log sederhana ke file harian. Tidak pernah melempar exception.</summary>
    public static class Log
    {
        private static readonly object Gate = new object();

        public static void Info(string message) => Write("INFO", message);
        public static void Warn(string message) => Write("WARN", message);
        public static void Error(string message, Exception ex = null) => Write("ERROR", ex == null ? message : message + " :: " + ex);

        private static void Write(string level, string message)
        {
            try
            {
                lock (Gate)
                {
                    var file = Path.Combine(AppPaths.LogDir, "masterbw-" + DateTime.Now.ToString("yyyyMMdd") + ".log");
                    File.AppendAllText(file, DateTime.Now.ToString("HH:mm:ss.fff") + " [" + level + "] " + message + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch
            {
                // log tidak boleh mengganggu Word/Excel
            }
        }
    }
}
