using Npgsql;
using SyncServer.Models;

namespace SyncServer.Services;

/// <summary>
/// PostgreSQL persistence for cell state.
/// </summary>
public class DatabaseService : IAsyncDisposable
{
    private readonly string _connectionString;
    private readonly ILogger<DatabaseService> _logger;

    public DatabaseService(IConfiguration configuration, ILogger<DatabaseService> logger)
    {
        _connectionString = configuration.GetConnectionString("PostgreSQL")
            ?? throw new InvalidOperationException("PostgreSQL connection string not configured");
        _logger = logger;
    }

    /// <summary>
    /// Initialize database schema.
    /// </summary>
    public async Task InitializeAsync()
    {
        await using var conn = new NpgsqlConnection(_connectionString);
        await conn.OpenAsync();

        await using var cmd = new NpgsqlCommand(@"
            CREATE TABLE IF NOT EXISTS cell_state (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                last_updated_by TEXT,
                last_updated TIMESTAMPTZ DEFAULT NOW()
            );
            CREATE INDEX IF NOT EXISTS idx_cell_state_updated ON cell_state(last_updated);
        ", conn);

        await cmd.ExecuteNonQueryAsync();
        _logger.LogInformation("Database initialized");
    }

    /// <summary>
    /// Upsert a cell value.
    /// </summary>
    public async Task SetValueAsync(string key, string value, string? updatedBy = null)
    {
        await using var conn = new NpgsqlConnection(_connectionString);
        await conn.OpenAsync();

        await using var cmd = new NpgsqlCommand(@"
            INSERT INTO cell_state (key, value, last_updated_by, last_updated)
            VALUES (@key, @value, @updatedBy, NOW())
            ON CONFLICT (key) DO UPDATE SET
                value = @value,
                last_updated_by = @updatedBy,
                last_updated = NOW()
        ", conn);

        cmd.Parameters.AddWithValue("key", key);
        cmd.Parameters.AddWithValue("value", value);
        cmd.Parameters.AddWithValue("updatedBy", (object?)updatedBy ?? DBNull.Value);

        await cmd.ExecuteNonQueryAsync();
    }

    /// <summary>
    /// Get a single cell value.
    /// </summary>
    public async Task<CellState?> GetValueAsync(string key)
    {
        await using var conn = new NpgsqlConnection(_connectionString);
        await conn.OpenAsync();

        await using var cmd = new NpgsqlCommand(@"
            SELECT key, value, last_updated_by, last_updated
            FROM cell_state WHERE key = @key
        ", conn);

        cmd.Parameters.AddWithValue("key", key);

        await using var reader = await cmd.ExecuteReaderAsync();
        if (await reader.ReadAsync())
        {
            return new CellState
            {
                Key = reader.GetString(0),
                Value = reader.GetString(1),
                LastUpdatedBy = reader.IsDBNull(2) ? null : reader.GetString(2),
                LastUpdated = reader.GetDateTime(3)
            };
        }

        return null;
    }

    /// <summary>
    /// Get all cell states.
    /// </summary>
    public async Task<List<CellState>> GetAllStatesAsync()
    {
        var states = new List<CellState>();

        await using var conn = new NpgsqlConnection(_connectionString);
        await conn.OpenAsync();

        await using var cmd = new NpgsqlCommand(@"
            SELECT key, value, last_updated_by, last_updated
            FROM cell_state ORDER BY key
        ", conn);

        await using var reader = await cmd.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            states.Add(new CellState
            {
                Key = reader.GetString(0),
                Value = reader.GetString(1),
                LastUpdatedBy = reader.IsDBNull(2) ? null : reader.GetString(2),
                LastUpdated = reader.GetDateTime(3)
            });
        }

        return states;
    }

    /// <summary>
    /// Delete a cell.
    /// </summary>
    public async Task<bool> DeleteAsync(string key)
    {
        await using var conn = new NpgsqlConnection(_connectionString);
        await conn.OpenAsync();

        await using var cmd = new NpgsqlCommand("DELETE FROM cell_state WHERE key = @key", conn);
        cmd.Parameters.AddWithValue("key", key);

        return await cmd.ExecuteNonQueryAsync() > 0;
    }

    public ValueTask DisposeAsync()
    {
        return ValueTask.CompletedTask;
    }
}
