using System.Collections.Concurrent;
using SyncServer.Models;

namespace SyncServer.Services;

/// <summary>
/// Thread-safe in-memory storage for cell state.
/// Provides current values for late-joining clients.
/// </summary>
public class StateManager
{
    private readonly ConcurrentDictionary<string, CellState> _state = new();

    /// <summary>
    /// Update or add a cell's state.
    /// </summary>
    public void SetValue(string key, string value, string? updatedBy = null)
    {
        _state.AddOrUpdate(
            key,
            _ => new CellState
            {
                Key = key,
                Value = value,
                LastUpdatedBy = updatedBy,
                LastUpdated = DateTime.UtcNow
            },
            (_, existing) =>
            {
                existing.Value = value;
                existing.LastUpdatedBy = updatedBy;
                existing.LastUpdated = DateTime.UtcNow;
                return existing;
            });
    }

    /// <summary>
    /// Get a cell's current value.
    /// </summary>
    public CellState? GetValue(string key)
    {
        return _state.TryGetValue(key, out var state) ? state : null;
    }

    /// <summary>
    /// Get all current cell states.
    /// </summary>
    public IEnumerable<CellState> GetAllStates()
    {
        return _state.Values.ToList();
    }

    /// <summary>
    /// Remove a cell from tracking.
    /// </summary>
    public bool Remove(string key)
    {
        return _state.TryRemove(key, out _);
    }

    /// <summary>
    /// Get the number of tracked cells.
    /// </summary>
    public int Count => _state.Count;
}
