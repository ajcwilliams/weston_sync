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
        /// SYNC formula: Registers a cell for synchronization.
        /// Usage: =SYNC("myKey", A1)
        /// </summary>
        /// <param name="key">Unique sync key</param>
        /// <param name="sourceValue">The value to sync (reference another cell)</param>
        /// <returns>The synced value</returns>
        public object Sync(string key, object sourceValue)
        {
            try
            {
                // Get calling cell info
                var app = (Application)ExcelDnaUtil.Application;
                var caller = app.Caller as Range;

                if (caller != null)
                {
                    var workbook = (Workbook)caller.Worksheet.Parent;
                    var worksheet = caller.Worksheet;

                    // Register this cell for tracking
                    CellTracker.Instance.TrackCell(
                        key,
                        workbook.Name,
                        worksheet.Name,
                        caller.Address[false, false]
                    );
                }

                // Check if we have a remote value
                var remoteValue = RtdServer.Instance?.GetValue(key);
                if (remoteValue != null)
                {
                    return remoteValue;
                }

                // Return the source value
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
