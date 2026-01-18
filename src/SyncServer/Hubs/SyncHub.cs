using Microsoft.AspNetCore.SignalR;
using SyncServer.Models;
using SyncServer.Services;

namespace SyncServer.Hubs;

/// <summary>
/// SignalR hub for real-time cell synchronization between Excel clients.
/// </summary>
public class SyncHub : Hub
{
    private readonly StateManager _stateManager;
    private readonly ILogger<SyncHub> _logger;

    public SyncHub(StateManager stateManager, ILogger<SyncHub> logger)
    {
        _stateManager = stateManager;
        _logger = logger;
    }

    /// <summary>
    /// Called when a client sends a cell update.
    /// Broadcasts to all other clients and updates state.
    /// </summary>
    public async Task SendUpdate(string key, string value)
    {
        var senderId = Context.ConnectionId;

        _logger.LogInformation("Update received: {Key} = {Value} from {SenderId}",
            key, value, senderId);

        // Store the current state (persists to database)
        await _stateManager.SetValueAsync(key, value, senderId);

        // Broadcast to all OTHER clients (exclude sender)
        await Clients.Others.SendAsync("ReceiveUpdate", new CellUpdate
        {
            Key = key,
            Value = value,
            SenderId = senderId,
            Timestamp = DateTime.UtcNow
        });
    }

    /// <summary>
    /// Called when a client wants to get the current value of a cell.
    /// </summary>
    public CellState? GetValue(string key)
    {
        return _stateManager.GetValue(key);
    }

    /// <summary>
    /// Called when a client connects and wants all current state.
    /// </summary>
    public IEnumerable<CellState> GetAllState()
    {
        _logger.LogInformation("Client {ConnectionId} requesting full state sync",
            Context.ConnectionId);
        return _stateManager.GetAllStates();
    }

    /// <summary>
    /// Called when a new client connects.
    /// </summary>
    public override async Task OnConnectedAsync()
    {
        _logger.LogInformation("Client connected: {ConnectionId}", Context.ConnectionId);
        await base.OnConnectedAsync();
    }

    /// <summary>
    /// Called when a client disconnects.
    /// </summary>
    public override async Task OnDisconnectedAsync(Exception? exception)
    {
        _logger.LogInformation("Client disconnected: {ConnectionId}, Exception: {Exception}",
            Context.ConnectionId, exception?.Message ?? "None");
        await base.OnDisconnectedAsync(exception);
    }
}
