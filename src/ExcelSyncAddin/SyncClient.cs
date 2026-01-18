using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.SignalR.Client;

namespace ExcelSyncAddin
{
    /// <summary>
    /// SignalR client for bidirectional communication with the sync server.
    /// </summary>
    public class SyncClient : IDisposable
    {
        private HubConnection _connection;
        private readonly string _serverUrl;
        private bool _isConnected;
        private bool _isDisposed;
        private CancellationTokenSource _reconnectCts;
        private static SyncClient _instance;

        // Debounce: track last sent values to avoid duplicate sends
        private readonly ConcurrentDictionary<string, string> _lastSentValues = new ConcurrentDictionary<string, string>();
        private readonly ConcurrentDictionary<string, DateTime> _lastSentTimes = new ConcurrentDictionary<string, DateTime>();
        private const int DebounceMs = 100; // Minimum ms between sends for same key

        /// <summary>
        /// Singleton instance.
        /// </summary>
        public static SyncClient Instance => _instance;

        /// <summary>
        /// Event fired when a cell update is received from the server.
        /// </summary>
        public event Action<string, string> OnUpdateReceived;

        /// <summary>
        /// Event fired when connection state changes.
        /// </summary>
        public event Action<bool> OnConnectionStateChanged;

        /// <summary>
        /// Whether currently connected to the server.
        /// </summary>
        public bool IsConnected => _isConnected;

        public SyncClient(string serverUrl)
        {
            _serverUrl = serverUrl;
            _instance = this;
            InitializeConnection();
        }

        private void InitializeConnection()
        {
            _connection = new HubConnectionBuilder()
                .WithUrl(_serverUrl)
                .WithAutomaticReconnect(new[] {
                    TimeSpan.Zero,
                    TimeSpan.FromSeconds(2),
                    TimeSpan.FromSeconds(5),
                    TimeSpan.FromSeconds(10),
                    TimeSpan.FromSeconds(30)
                })
                .Build();

            // Handle incoming updates
            _connection.On<CellUpdateMessage>("ReceiveUpdate", message =>
            {
                OnUpdateReceived?.Invoke(message.Key, message.Value);

                // Also update the RTD server directly
                RtdServer.Instance?.UpdateValue(message.Key, message.Value);
            });

            _connection.Reconnecting += error =>
            {
                _isConnected = false;
                OnConnectionStateChanged?.Invoke(false);
                Logger.Log($"Reconnecting... Error: {error?.Message}");
                return Task.CompletedTask;
            };

            _connection.Reconnected += connectionId =>
            {
                _isConnected = true;
                OnConnectionStateChanged?.Invoke(true);
                Logger.Log($"Reconnected with ID: {connectionId}");

                // Request full state sync after reconnect
                _ = RequestFullStateSyncAsync();
                return Task.CompletedTask;
            };

            _connection.Closed += error =>
            {
                _isConnected = false;
                OnConnectionStateChanged?.Invoke(false);
                Logger.Log($"Connection closed. Error: {error?.Message}");
                return Task.CompletedTask;
            };
        }

        /// <summary>
        /// Connect to the sync server.
        /// </summary>
        public async Task ConnectAsync()
        {
            if (_isConnected) return;

            try
            {
                await _connection.StartAsync();
                _isConnected = true;
                OnConnectionStateChanged?.Invoke(true);
                Logger.Log($"Connected to {_serverUrl}");

                // Get initial state
                await RequestFullStateSyncAsync();
            }
            catch (Exception ex)
            {
                Logger.Log($"Connection failed: {ex.Message}");
                _isConnected = false;
                OnConnectionStateChanged?.Invoke(false);

                // Start background reconnection
                StartReconnectLoop();
            }
        }

        /// <summary>
        /// Request all current cell states from the server.
        /// </summary>
        private async Task RequestFullStateSyncAsync()
        {
            try
            {
                var states = await _connection.InvokeAsync<CellStateMessage[]>("GetAllState");

                foreach (var state in states)
                {
                    RtdServer.Instance?.UpdateValue(state.Key, state.Value);
                }

                Logger.Log($"Synced {states.Length} cell values from server");
            }
            catch (Exception ex)
            {
                Logger.Log($"Failed to get initial state: {ex.Message}");
            }
        }

        private void StartReconnectLoop()
        {
            _reconnectCts?.Cancel();
            _reconnectCts = new CancellationTokenSource();

            Task.Run(async () =>
            {
                var delays = new[] { 1000, 2000, 5000, 10000, 30000 };
                int attempt = 0;

                while (!_isConnected && !_reconnectCts.Token.IsCancellationRequested)
                {
                    var delay = delays[Math.Min(attempt, delays.Length - 1)];
                    await Task.Delay(delay, _reconnectCts.Token);

                    try
                    {
                        await _connection.StartAsync(_reconnectCts.Token);
                        _isConnected = true;
                        OnConnectionStateChanged?.Invoke(true);
                        Logger.Log("Reconnected successfully");
                        await RequestFullStateSyncAsync();
                        break;
                    }
                    catch
                    {
                        attempt++;
                        Logger.Log($"Reconnect attempt {attempt} failed");
                    }
                }
            }, _reconnectCts.Token);
        }

        /// <summary>
        /// Send a cell update to the server.
        /// </summary>
        public async Task SendUpdateAsync(string key, string value)
        {
            if (!_isConnected)
            {
                Logger.Log($"Cannot send update - not connected. Key: {key}");
                return;
            }

            try
            {
                await _connection.InvokeAsync("SendUpdate", key, value);
                Logger.Log($"Sent update: {key} = {value}");
            }
            catch (Exception ex)
            {
                Logger.Log($"Failed to send update: {ex.Message}");
            }
        }

        /// <summary>
        /// Send update synchronously (for use in Excel event handlers).
        /// </summary>
        public void SendUpdate(string key, string value)
        {
            Task.Run(() => SendUpdateAsync(key, value));
        }

        /// <summary>
        /// Send update with debouncing - skips if same value was just sent.
        /// Used by SYNC formula to avoid spamming during Excel recalculation.
        /// </summary>
        public void SendUpdateDebounced(string key, string value)
        {
            // Skip if same value was already sent
            if (_lastSentValues.TryGetValue(key, out var lastValue) && lastValue == value)
            {
                return;
            }

            // Skip if sent too recently (debounce)
            if (_lastSentTimes.TryGetValue(key, out var lastTime))
            {
                if ((DateTime.UtcNow - lastTime).TotalMilliseconds < DebounceMs)
                {
                    return;
                }
            }

            // Update tracking
            _lastSentValues[key] = value;
            _lastSentTimes[key] = DateTime.UtcNow;

            // Send async
            Task.Run(() => SendUpdateAsync(key, value));
        }

        /// <summary>
        /// Disconnect from the server.
        /// </summary>
        public async Task DisconnectAsync()
        {
            _reconnectCts?.Cancel();

            if (_connection != null)
            {
                try
                {
                    await _connection.StopAsync();
                }
                catch
                {
                    // Ignore errors during shutdown
                }
            }

            _isConnected = false;
            OnConnectionStateChanged?.Invoke(false);
        }

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            _reconnectCts?.Cancel();
            _reconnectCts?.Dispose();

            try
            {
                _connection?.DisposeAsync().AsTask().Wait(TimeSpan.FromSeconds(5));
            }
            catch
            {
                // Ignore errors during disposal
            }

            _instance = null;
        }
    }

    /// <summary>
    /// Message model for cell updates.
    /// </summary>
    public class CellUpdateMessage
    {
        public string Key { get; set; }
        public string Value { get; set; }
        public string SenderId { get; set; }
        public DateTime Timestamp { get; set; }
    }

    /// <summary>
    /// Message model for cell state.
    /// </summary>
    public class CellStateMessage
    {
        public string Key { get; set; }
        public string Value { get; set; }
        public string LastUpdatedBy { get; set; }
        public DateTime LastUpdated { get; set; }
    }
}
