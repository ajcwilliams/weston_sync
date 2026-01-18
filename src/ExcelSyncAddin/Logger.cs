using System;
using System.IO;

namespace ExcelSyncAddin
{
    /// <summary>
    /// Simple file logger for debugging the add-in.
    /// </summary>
    public static class Logger
    {
        private static readonly string LogPath;
        private static readonly object LockObj = new object();
        private static bool _enabled = true;

        static Logger()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var logDir = Path.Combine(appData, "ExcelSyncAddin");
            Directory.CreateDirectory(logDir);
            LogPath = Path.Combine(logDir, "sync.log");
        }

        public static bool Enabled
        {
            get => _enabled;
            set => _enabled = value;
        }

        public static void Log(string message)
        {
            if (!_enabled) return;

            try
            {
                lock (LockObj)
                {
                    var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} | {message}";
                    File.AppendAllText(LogPath, line + Environment.NewLine);
                }
            }
            catch
            {
                // Ignore logging errors
            }
        }

        public static void Log(string format, params object[] args)
        {
            Log(string.Format(format, args));
        }

        public static void Error(string message, Exception ex)
        {
            Log($"ERROR: {message} - {ex.GetType().Name}: {ex.Message}");
        }
    }
}
