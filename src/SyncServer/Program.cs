using SyncServer.Hubs;
using SyncServer.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services
builder.Services.AddSignalR();
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
