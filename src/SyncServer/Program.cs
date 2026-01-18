using SyncServer.Hubs;
using SyncServer.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services
builder.Services.AddSignalR();
builder.Services.AddSingleton<DatabaseService>();
builder.Services.AddSingleton<StateManager>();

// Configure CORS for Excel clients
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

var app = builder.Build();

// Initialize database and load state
var db = app.Services.GetRequiredService<DatabaseService>();
var stateManager = app.Services.GetRequiredService<StateManager>();

await db.InitializeAsync();
await stateManager.LoadFromDatabaseAsync();

app.UseCors();

// Map the SignalR hub
app.MapHub<SyncHub>("/sync");

// Health check endpoint
app.MapGet("/health", () => Results.Ok(new { status = "healthy", timestamp = DateTime.UtcNow }));

// Status endpoint showing connected state
app.MapGet("/status", (StateManager stateManager) => Results.Ok(new
{
    trackedCells = stateManager.Count,
    timestamp = DateTime.UtcNow
}));

Console.WriteLine("Excel Sync Server starting...");
Console.WriteLine("SignalR Hub: /sync");
Console.WriteLine("Health check: /health");

app.Run();
