using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Interop.Excel;

namespace ExcelSyncAddin
{
    /// <summary>
    /// Main COM Add-in entry point for Excel.
    /// Handles Excel lifecycle events and cell change detection.
    /// </summary>
    [ComVisible(true)]
    [Guid("D4E5F6A7-B8C9-0123-4567-89ABCDEF0123")]
    [ProgId("ExcelSync.Addin")]
    [ClassInterface(ClassInterfaceType.None)]
    public class SyncAddin : IDTExtensibility2
    {
        private Application _excelApp;
        private SyncClient _syncClient;
        private bool _isUpdatingFromServer;

        /// <summary>
        /// Called when the add-in is loaded.
        /// </summary>
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                _excelApp = (Application)application;
                Logger.Log("Add-in connecting...");

                // Initialize sync client
                _syncClient = new SyncClient(Config.ServerUrl);
                _syncClient.OnUpdateReceived += OnRemoteUpdateReceived;
                _syncClient.OnConnectionStateChanged += OnConnectionStateChanged;

                // Hook Excel events
                _excelApp.SheetChange += OnSheetChange;
                _excelApp.WorkbookOpen += OnWorkbookOpen;
                _excelApp.WorkbookBeforeClose += OnWorkbookBeforeClose;

                // Connect to server
                _ = _syncClient.ConnectAsync();

                Logger.Log("Add-in connected successfully");
            }
            catch (Exception ex)
            {
                Logger.Error("OnConnection failed", ex);
            }
        }

        /// <summary>
        /// Called when the add-in is unloaded.
        /// </summary>
        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                Logger.Log("Add-in disconnecting...");

                // Unhook events
                if (_excelApp != null)
                {
                    _excelApp.SheetChange -= OnSheetChange;
                    _excelApp.WorkbookOpen -= OnWorkbookOpen;
                    _excelApp.WorkbookBeforeClose -= OnWorkbookBeforeClose;
                }

                // Disconnect from server
                _syncClient?.Dispose();

                // Clear tracking
                CellTracker.Instance.Clear();

                _excelApp = null;
                Logger.Log("Add-in disconnected");
            }
            catch (Exception ex)
            {
                Logger.Error("OnDisconnection failed", ex);
            }
        }

        /// <summary>
        /// Called when Excel starts up with the add-in.
        /// </summary>
        public void OnStartupComplete(ref Array custom)
        {
            Logger.Log("Startup complete");
        }

        /// <summary>
        /// Called when Excel is about to shut down.
        /// </summary>
        public void OnBeginShutdown(ref Array custom)
        {
            Logger.Log("Beginning shutdown");
        }

        /// <summary>
        /// Not used.
        /// </summary>
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>
        /// Handle cell changes in Excel.
        /// </summary>
        private void OnSheetChange(object sheet, Range target)
        {
            // Skip if this change came from a server update (prevent echo)
            if (_isUpdatingFromServer) return;

            try
            {
                var worksheet = (Worksheet)sheet;
                var workbookName = worksheet.Parent.Name;
                var sheetName = worksheet.Name;

                // Check each cell in the changed range
                foreach (Range cell in target.Cells)
                {
                    var address = cell.Address[false, false]; // A1 format
                    var key = CellTracker.Instance.GetKeyForAddress(workbookName, sheetName, address);

                    if (key != null)
                    {
                        var value = cell.Value2?.ToString() ?? "";
                        Logger.Log($"Local change detected: {key} = {value}");

                        // Send to server
                        _syncClient?.SendUpdate(key, value);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("OnSheetChange failed", ex);
            }
        }

        /// <summary>
        /// Handle remote updates from the server.
        /// </summary>
        private void OnRemoteUpdateReceived(string key, string value)
        {
            try
            {
                var trackedCell = CellTracker.Instance.GetTrackedCell(key);
                if (trackedCell == null) return;

                Logger.Log($"Remote update received: {key} = {value}");

                // Find the cell and update it
                _isUpdatingFromServer = true;
                try
                {
                    var workbook = _excelApp.Workbooks[trackedCell.WorkbookName];
                    var worksheet = (Worksheet)workbook.Sheets[trackedCell.SheetName];
                    var cell = worksheet.Range[trackedCell.CellAddress];

                    // Update the cell value
                    cell.Value2 = value;
                }
                finally
                {
                    _isUpdatingFromServer = false;
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to apply remote update for {key}", ex);
                _isUpdatingFromServer = false;
            }
        }

        /// <summary>
        /// Handle connection state changes.
        /// </summary>
        private void OnConnectionStateChanged(bool connected)
        {
            Logger.Log($"Connection state: {(connected ? "Connected" : "Disconnected")}");

            // Could update status bar or UI here
            try
            {
                _excelApp.StatusBar = connected
                    ? "Excel Sync: Connected"
                    : "Excel Sync: Disconnected";
            }
            catch
            {
                // Status bar may not be available
            }
        }

        /// <summary>
        /// Handle workbook open.
        /// </summary>
        private void OnWorkbookOpen(Workbook workbook)
        {
            Logger.Log($"Workbook opened: {workbook.Name}");
            // Could scan for SYNC formulas here
        }

        /// <summary>
        /// Handle workbook close.
        /// </summary>
        private void OnWorkbookBeforeClose(Workbook workbook, ref bool cancel)
        {
            Logger.Log($"Workbook closing: {workbook.Name}");

            // Remove tracked cells for this workbook
            foreach (var cell in CellTracker.Instance.GetAllTrackedCells())
            {
                if (cell.WorkbookName == workbook.Name)
                {
                    CellTracker.Instance.UntrackCell(cell.Key);
                }
            }
        }
    }
}
