# Excel Sync Add-in

Real-time bidirectional cell synchronization across multiple Excel spreadsheets.

## Overview

This solution enables multiple Excel users to sync cell values in real-time. When one user changes a tracked cell, the change is instantly reflected in all other connected spreadsheets.

## Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                         SYNC SERVER                                  │
│              (ASP.NET Core + SignalR)                               │
│                    [Your VPS]                                        │
└─────────────────────────────────────────────────────────────────────┘
                              ↑↓ WebSocket
        ┌─────────────────────┼─────────────────────┐
        ↓                     ↓                     ↓
   Excel A               Excel B               Excel C
   (COM Add-in)          (COM Add-in)          (COM Add-in)
```

## Components

- **SyncServer**: ASP.NET Core server with SignalR hub for real-time messaging
- **ExcelSyncAddin**: COM add-in that handles both sending and receiving updates

## Requirements

### Server
- .NET 8.0 SDK
- Linux/Windows/macOS host

### Excel Client
- Windows with Excel 2016+
- .NET Framework 4.8
- Visual Studio 2022 (for building)

## Quick Start

### 1. Start the Server

```bash
cd src/SyncServer
dotnet run
```

Server will start at `http://localhost:5000`

### 2. Build the Excel Add-in

Open `ExcelSyncAddin.sln` in Visual Studio and build in Release mode.

### 3. Register the Add-in

Run as Administrator:
```batch
tools\Register.bat
```

### 4. Use in Excel

**Option A: RTD Formula (Recommended)**
```excel
=RTD("ExcelSync.RtdServer", "", "my_key")
```

This cell will receive updates whenever another client sends an update with the key "my_key".

**Option B: SYNC Formula**
```excel
=SYNC("revenue_q1", A1)
```

This both tracks cell A1 for outbound sync AND receives inbound updates.

## Configuration

Edit `src/ExcelSyncAddin/App.config`:

```xml
<appSettings>
  <add key="ServerUrl" value="http://your-server:5000/sync"/>
  <add key="RtdRefreshMs" value="100"/>
  <add key="LoggingEnabled" value="true"/>
</appSettings>
```

## Server Deployment

### Using Docker
```dockerfile
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app
COPY src/SyncServer/bin/Release/net8.0/publish .
EXPOSE 5000
ENTRYPOINT ["dotnet", "SyncServer.dll"]
```

### Manual Deployment
```bash
cd src/SyncServer
dotnet publish -c Release
# Copy publish folder to server
dotnet SyncServer.dll
```

### Reverse Proxy (nginx)
```nginx
location /sync {
    proxy_pass http://localhost:5000;
    proxy_http_version 1.1;
    proxy_set_header Upgrade $http_upgrade;
    proxy_set_header Connection "upgrade";
    proxy_set_header Host $host;
}
```

## Logging

Client logs are written to:
```
%LOCALAPPDATA%\ExcelSyncAddin\sync.log
```

## Troubleshooting

### RTD formula shows #N/A
1. Ensure the add-in is registered (run Register.bat as admin)
2. Restart Excel after registration
3. Check that the server is running

### Connection issues
1. Verify server URL in App.config
2. Check firewall rules for port 5000
3. Review logs at %LOCALAPPDATA%\ExcelSyncAddin\sync.log

### Formula not updating
1. Check Excel's RTD throttle: File → Options → Formulas → "Enable background refresh"
2. Press Ctrl+Alt+F9 to force recalculation

## API Reference

### SignalR Hub Methods

**Send Update**
```csharp
await connection.InvokeAsync("SendUpdate", key, value);
```

**Get Single Value**
```csharp
var state = await connection.InvokeAsync<CellState>("GetValue", key);
```

**Get All State**
```csharp
var states = await connection.InvokeAsync<CellState[]>("GetAllState");
```

### SignalR Hub Events

**Receive Update**
```csharp
connection.On<CellUpdate>("ReceiveUpdate", update => {
    // Handle incoming update
});
```

## License

MIT
