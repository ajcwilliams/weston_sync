using System.Collections.Concurrent;
using SyncServer.Models;

namespace SyncServer.Services;

/// <summary>
/// Thread-safe storage for cell state.
/// Uses in-memory cache for fast reads, PostgreSQL for persistence.
/// </summary>
public class StateManager
{
    private readonly ConcurrentDictionary<string, CellState> _cache = new();
    private readonly DatabaseService _db;
    private readonly ILogger<StateManager> _logger;

    public StateManager(DatabaseService db, ILogger<StateManager> logger)
    {
        _db = db;
        _logger = logger;
    }

    /// <summary>
    /// Load all state from database into cache.
    /// </summary>
    public async Task LoadFromDatabaseAsync()
    {
        var states = await _db.GetAllStatesAsync();
        foreach (var state in states)
        {
            _cache[state.Key] = state;
        }
        _logger.LogInformation("Loaded {Count} cell states from database", states.Count);
    }

    /// <summary>
    /// Update or add a cell's state.
    /// </summary>
    public async Task SetValueAsync(string key, string value, string? updatedBy = null)
    {
        var state = new CellState
        {
            Key = key,
            Value = value,
            LastUpdatedBy = updatedBy,
            LastUpdated = DateTime.UtcNow
        };

        // Update cache immediately
        _cache[key] = state;

        // Persist to database (fire and forget for speed, but log errors)
        _ = Task.Run(async () =>
        {
            try
            {
                await _db.SetValueAsync(key, value, updatedBy);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to persist value for key {Key}", key);
            }
        });
    }

    /// <summary>
    /// Get a cell's current value from cache.
    /// </summary>
    public CellState? GetValue(string key)
    {
        return _cache.TryGetValue(key, out var state) ? state : null;
    }

    /// <summary>
    /// Get all current cell states from cache.
    /// </summary>
    public IEnumerable<CellState> GetAllStates()
    {
        return _cache.Values.ToList();
    }

    /// <summary>
    /// Remove a cell from tracking.
    /// </summary>
    public async Task<bool> RemoveAsync(string key)
    {
        var removed = _cache.TryRemove(key, out _);
        if (removed)
        {
            await _db.DeleteAsync(key);
        }
        return removed;
    }

    /// <summary>
    /// Get the number of tracked cells.
    /// </summary>
    public int Count => _cache.Count;
}
