namespace SyncServer.Models;

/// <summary>
/// Represents a cell value update message between clients.
/// </summary>
public class CellUpdate
{
    /// <summary>
    /// Unique key identifying the synced cell (e.g., "revenue_q1").
    /// </summary>
    public required string Key { get; set; }

    /// <summary>
    /// The cell value as a string.
    /// </summary>
    public required string Value { get; set; }

    /// <summary>
    /// Connection ID of the sender (to prevent echo).
    /// </summary>
    public string? SenderId { get; set; }

    /// <summary>
    /// Timestamp when the update was created.
    /// </summary>
    public DateTime Timestamp { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// Represents the current state of a synced cell.
/// </summary>
public class CellState
{
    public required string Key { get; set; }
    public required string Value { get; set; }
    public string? LastUpdatedBy { get; set; }
    public DateTime LastUpdated { get; set; } = DateTime.UtcNow;
}
