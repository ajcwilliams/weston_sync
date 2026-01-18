using System;
using System.Configuration;

namespace ExcelSyncAddin
{
    /// <summary>
    /// Configuration settings for the sync add-in.
    /// </summary>
    public static class Config
    {
        private static string _serverUrl;
        private static int _rtdRefreshMs = 100;
        private static bool _loggingEnabled = true;

        /// <summary>
        /// URL of the sync server (e.g., "http://localhost:5000/sync").
        /// </summary>
        public static string ServerUrl
        {
            get
            {
                if (_serverUrl == null)
                {
                    _serverUrl = GetSetting("ServerUrl", "http://localhost:5000/sync");
                }
                return _serverUrl;
            }
            set => _serverUrl = value;
        }

        /// <summary>
        /// RTD refresh interval in milliseconds.
        /// </summary>
        public static int RtdRefreshMs
        {
            get => _rtdRefreshMs;
            set => _rtdRefreshMs = Math.Max(50, value);
        }

        /// <summary>
        /// Whether logging is enabled.
        /// </summary>
        public static bool LoggingEnabled
        {
            get => _loggingEnabled;
            set
            {
                _loggingEnabled = value;
                Logger.Enabled = value;
            }
        }

        private static string GetSetting(string key, string defaultValue)
        {
            try
            {
                var value = ConfigurationManager.AppSettings[key];
                return string.IsNullOrEmpty(value) ? defaultValue : value;
            }
            catch
            {
                return defaultValue;
            }
        }

        private static int GetSetting(string key, int defaultValue)
        {
            try
            {
                var value = ConfigurationManager.AppSettings[key];
                return int.TryParse(value, out var result) ? result : defaultValue;
            }
            catch
            {
                return defaultValue;
            }
        }
    }
}
