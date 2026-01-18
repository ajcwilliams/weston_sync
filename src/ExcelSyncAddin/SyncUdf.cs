using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelSyncAddin
{
    /// <summary>
    /// User-Defined Functions (UDFs) for Excel.
    /// Provides the =SYNC() formula function.
    /// </summary>
    [ComVisible(true)]
    [Guid("C3D4E5F6-A7B8-9012-3456-789ABCDEF012")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class SyncUdf
    {
        /// <summary>
        /// WB_SYNC formula: Bidirectional cell synchronization.
        /// Usage: =WB_SYNC("myKey", A1)
        ///
        /// When A1 changes, Excel recalculates this formula, which sends the new value.
        /// When another client sends an update, RTD triggers recalc and this returns the new value.
        /// </summary>
        /// <param name="key">Unique sync key</param>
        /// <param name="sourceValue">The value to sync (reference another cell)</param>
        /// <returns>The synced value (remote if available, else local)</returns>
        public object WB_SYNC(string key, object sourceValue)
        {
            try
            {
                var localValue = sourceValue?.ToString() ?? "";

                // Send local value to server (debounced to avoid spam during recalc)
                SyncClient.Instance?.SendUpdateDebounced(key, localValue);

                // Check if we have a remote value from another client
                var remoteValue = RtdServer.Instance?.GetValue(key);

                // Return remote value if it exists and differs, else local
                if (remoteValue != null && remoteValue.ToString() != localValue)
                {
                    return remoteValue;
                }

                return sourceValue ?? "";
            }
            catch (Exception ex)
            {
                Logger.Error("SYNC function error", ex);
                return "#SYNC_ERROR";
            }
        }

        /// <summary>
        /// SYNCSTATUS formula: Returns the current connection status.
        /// Usage: =SYNCSTATUS()
        /// </summary>
        public string SyncStatus()
        {
            return SyncClient.Instance?.IsConnected == true ? "Connected" : "Disconnected";
        }

        /// <summary>
        /// SYNCCOUNT formula: Returns the number of tracked cells.
        /// Usage: =SYNCCOUNT()
        /// </summary>
        public int SyncCount()
        {
            return CellTracker.Instance.Count;
        }
    }

    /// <summary>
    /// Helper to get Excel Application instance.
    /// Note: In a real implementation, this would come from Excel-DNA or similar.
    /// For COM add-in, we access via the add-in's stored reference.
    /// </summary>
    internal static class ExcelDnaUtil
    {
        private static Application _application;

        public static Application Application
        {
            get => _application;
            set => _application = value;
        }
    }
}
