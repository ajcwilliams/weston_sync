using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace ExcelSyncAddin
{
    /// <summary>
    /// COM-visible RTD Server for receiving real-time updates in Excel.
    /// Usage in Excel: =RTD("ExcelSync.RtdServer", "", "key")
    /// </summary>
    [ComVisible(true)]
    [Guid("E5F6A7B8-C9D0-1234-5678-9ABCDEF01234")]
    [ProgId("ExcelSync.RtdServer")]
    [ClassInterface(ClassInterfaceType.None)]
    public class RtdServer : IRtdServer
    {
        private IRTDUpdateEvent _callback;
        private readonly Dictionary<int, string> _topicIdToKey = new Dictionary<int, string>();
        private readonly Dictionary<string, int> _keyToTopicId = new Dictionary<string, int>();
        private readonly Dictionary<string, object> _values = new Dictionary<string, object>();
        private readonly object _lock = new object();
        private Timer _heartbeatTimer;
        private bool _isRunning;
        private static RtdServer _instance;

        /// <summary>
        /// Singleton instance for external access (from SyncClient).
        /// </summary>
        public static RtdServer Instance => _instance;

        /// <summary>
        /// Called by Excel when the RTD server starts.
        /// </summary>
        public int ServerStart(IRTDUpdateEvent callback)
        {
            _callback = callback;
            _isRunning = true;
            _instance = this;

            // Heartbeat to keep connection alive
            _heartbeatTimer = new Timer(_ =>
            {
                if (_isRunning && _callback != null)
                {
                    try
                    {
                        _callback.HeartbeatInterval = 5000;
                    }
                    catch
                    {
                        // Excel may have closed
                    }
                }
            }, null, 5000, 5000);

            return 1; // Success
        }

        /// <summary>
        /// Called by Excel when a cell subscribes to a topic.
        /// =RTD("ExcelSync.RtdServer", "", "myKey")
        /// </summary>
        public object ConnectData(int topicId, ref Array strings, ref bool getNewValues)
        {
            if (strings.Length < 1)
                return "#ERROR: Key required";

            string key = strings.GetValue(0)?.ToString();
            if (string.IsNullOrEmpty(key))
                return "#ERROR: Empty key";

            lock (_lock)
            {
                _topicIdToKey[topicId] = key;
                _keyToTopicId[key] = topicId;

                // Return current value if we have one
                if (_values.TryGetValue(key, out var currentValue))
                {
                    getNewValues = true;
                    return currentValue;
                }
            }

            getNewValues = false;
            return "#WAITING";
        }

        /// <summary>
        /// Called by Excel when a cell unsubscribes from a topic.
        /// </summary>
        public void DisconnectData(int topicId)
        {
            lock (_lock)
            {
                if (_topicIdToKey.TryGetValue(topicId, out var key))
                {
                    _topicIdToKey.Remove(topicId);
                    _keyToTopicId.Remove(key);
                }
            }
        }

        /// <summary>
        /// Called by Excel to get updated values.
        /// </summary>
        public Array RefreshData(ref int topicCount)
        {
            lock (_lock)
            {
                var updatedTopics = new List<int>();
                var updatedValues = new List<object>();

                foreach (var kvp in _topicIdToKey)
                {
                    int topicId = kvp.Key;
                    string key = kvp.Value;

                    if (_values.TryGetValue(key, out var value))
                    {
                        updatedTopics.Add(topicId);
                        updatedValues.Add(value);
                    }
                }

                topicCount = updatedTopics.Count;

                if (topicCount == 0)
                {
                    return new object[0, 0];
                }

                // Create 2D array: [0,i] = topicId, [1,i] = value
                var result = new object[2, topicCount];
                for (int i = 0; i < topicCount; i++)
                {
                    result[0, i] = updatedTopics[i];
                    result[1, i] = updatedValues[i];
                }

                return result;
            }
        }

        /// <summary>
        /// Called by Excel to check if server is still alive.
        /// </summary>
        public int Heartbeat()
        {
            return 1; // Alive
        }

        /// <summary>
        /// Called by Excel when the RTD server is terminated.
        /// </summary>
        public void ServerTerminate()
        {
            _isRunning = false;
            _heartbeatTimer?.Dispose();
            _instance = null;

            lock (_lock)
            {
                _topicIdToKey.Clear();
                _keyToTopicId.Clear();
                _values.Clear();
            }
        }

        /// <summary>
        /// Update a value from external source (called by SyncClient).
        /// This triggers Excel to refresh the cell.
        /// </summary>
        public void UpdateValue(string key, object value)
        {
            lock (_lock)
            {
                _values[key] = value;
            }

            // Notify Excel that data has changed
            if (_callback != null)
            {
                try
                {
                    _callback.UpdateNotify();
                }
                catch
                {
                    // Excel may have closed
                }
            }
        }

        /// <summary>
        /// Check if a key is being tracked by any Excel cell.
        /// </summary>
        public bool IsKeyTracked(string key)
        {
            lock (_lock)
            {
                return _keyToTopicId.ContainsKey(key);
            }
        }

        /// <summary>
        /// Get current value for a key.
        /// </summary>
        public object GetValue(string key)
        {
            lock (_lock)
            {
                return _values.TryGetValue(key, out var value) ? value : null;
            }
        }
    }
}
