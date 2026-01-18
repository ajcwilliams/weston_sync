using System;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace ExcelSyncAddin
{
    /// <summary>
    /// Tracks which cells are being synced and their associated keys.
    /// Thread-safe for concurrent access from Excel and SignalR threads.
    /// </summary>
    public class CellTracker
    {
        private readonly ConcurrentDictionary<string, TrackedCell> _trackedCells = new ConcurrentDictionary<string, TrackedCell>();
        private readonly ConcurrentDictionary<string, string> _addressToKey = new ConcurrentDictionary<string, string>();
        private static CellTracker _instance;

        /// <summary>
        /// Singleton instance.
        /// </summary>
        public static CellTracker Instance => _instance ?? (_instance = new CellTracker());

        /// <summary>
        /// Register a cell for syncing.
        /// </summary>
        /// <param name="key">Unique sync key (e.g., "revenue_q1")</param>
        /// <param name="workbookName">Name of the workbook</param>
        /// <param name="sheetName">Name of the worksheet</param>
        /// <param name="cellAddress">Cell address (e.g., "A1")</param>
        public void TrackCell(string key, string workbookName, string sheetName, string cellAddress)
        {
            var fullAddress = GetFullAddress(workbookName, sheetName, cellAddress);

            var trackedCell = new TrackedCell
            {
                Key = key,
                WorkbookName = workbookName,
                SheetName = sheetName,
                CellAddress = cellAddress,
                FullAddress = fullAddress
            };

            _trackedCells[key] = trackedCell;
            _addressToKey[fullAddress] = key;

            Logger.Log($"Tracking cell: {key} -> {fullAddress}");
        }

        /// <summary>
        /// Unregister a cell from syncing.
        /// </summary>
        public void UntrackCell(string key)
        {
            if (_trackedCells.TryRemove(key, out var cell))
            {
                _addressToKey.TryRemove(cell.FullAddress, out _);
                Logger.Log($"Untracked cell: {key}");
            }
        }

        /// <summary>
        /// Check if a cell address is being tracked.
        /// </summary>
        public bool IsTracked(string workbookName, string sheetName, string cellAddress)
        {
            var fullAddress = GetFullAddress(workbookName, sheetName, cellAddress);
            return _addressToKey.ContainsKey(fullAddress);
        }

        /// <summary>
        /// Get the sync key for a cell address.
        /// </summary>
        public string GetKeyForAddress(string workbookName, string sheetName, string cellAddress)
        {
            var fullAddress = GetFullAddress(workbookName, sheetName, cellAddress);
            return _addressToKey.TryGetValue(fullAddress, out var key) ? key : null;
        }

        /// <summary>
        /// Get tracked cell info by key.
        /// </summary>
        public TrackedCell GetTrackedCell(string key)
        {
            return _trackedCells.TryGetValue(key, out var cell) ? cell : null;
        }

        /// <summary>
        /// Get all tracked cells.
        /// </summary>
        public IEnumerable<TrackedCell> GetAllTrackedCells()
        {
            return _trackedCells.Values;
        }

        /// <summary>
        /// Check if a key is being tracked.
        /// </summary>
        public bool IsKeyTracked(string key)
        {
            return _trackedCells.ContainsKey(key);
        }

        /// <summary>
        /// Clear all tracked cells.
        /// </summary>
        public void Clear()
        {
            _trackedCells.Clear();
            _addressToKey.Clear();
            Logger.Log("Cleared all tracked cells");
        }

        /// <summary>
        /// Number of tracked cells.
        /// </summary>
        public int Count => _trackedCells.Count;

        private static string GetFullAddress(string workbookName, string sheetName, string cellAddress)
        {
            return $"[{workbookName}]{sheetName}!{cellAddress}";
        }
    }

    /// <summary>
    /// Information about a tracked cell.
    /// </summary>
    public class TrackedCell
    {
        public string Key { get; set; }
        public string WorkbookName { get; set; }
        public string SheetName { get; set; }
        public string CellAddress { get; set; }
        public string FullAddress { get; set; }
    }
}
