# Excel Bidirectional Sync Add-in Implementation Plan

## Overview
Build a unified C# COM Add-in for Excel that enables real-time bidirectional sync of user-defined cells across 5-20 distributed spreadsheets.

## Architecture

```
                         ┌─────────────────────────────┐
                         │       SYNC SERVER           │
                         │   (ASP.NET Core + SignalR)  │
                         │        [Your VPS]           │
                         └──────────────┬──────────────┘
                                        │ WebSocket
          ┌─────────────────────────────┼─────────────────────────────┐
          │                             │                             │
          ▼                             ▼                             ▼
┌──────────────────┐         ┌──────────────────┐         ┌──────────────────┐
│  Excel Client A  │         │  Excel Client B  │         │  Excel Client C  │
│ ┌──────────────┐ │         │ ┌──────────────┐ │         │ ┌──────────────┐ │
│ │ COM Add-in   │ │         │ │ COM Add-in   │ │         │ │ COM Add-in   │ │
│ │ ┌──────────┐ │ │         │ │ ┌──────────┐ │ │         │ │ ┌──────────┐ │ │
│ │ │RTD Server│◄┼─┼─ recv ──┼─┼─│RTD Server│◄┼─┼─ recv ──┼─┼─│RTD Server│ │ │
│ │ └──────────┘ │ │         │ │ └──────────┘ │ │         │ │ └──────────┘ │ │
│ │ ┌──────────┐ │ │         │ │ ┌──────────┐ │ │         │ │ ┌──────────┐ │ │
│ │ │Event Hook│─┼─┼─ send ──┼─┼─│Event Hook│─┼─┼─ send ──┼─┼─│Event Hook│ │ │
│ │ └──────────┘ │ │         │ │ └──────────┘ │ │         │ │ └──────────┘ │ │
│ │ ┌──────────┐ │ │         │ │ ┌──────────┐ │ │         │ │ ┌──────────┐ │ │
│ │ │WebSocket │ │ │         │ │ │WebSocket │ │ │         │ │ │WebSocket │ │ │
│ │ │  Client  │ │ │         │ │ │  Client  │ │ │         │ │ │  Client  │ │ │
│ │ └──────────┘ │ │         │ │ └──────────┘ │ │         │ │ └──────────┘ │ │
│ └──────────────┘ │         │ └──────────────┘ │         │ └──────────────┘ │
└──────────────────┘         └──────────────────┘         └──────────────────┘
```

## Project Structure

```
ExcelSyncAddin/
├── ExcelSyncAddin.sln
│
├── src/
│   ├── ExcelSyncAddin/                    # Main COM Add-in
│   │   ├── ExcelSyncAddin.csproj
│   │   ├── SyncAddin.cs                   # Add-in entry point, Excel event hooks
│   │   ├── RtdServer.cs                   # IRtdServer for inbound updates
│   │   ├── SyncClient.cs                  # WebSocket client (SignalR)
│   │   ├── TopicManager.cs                # Thread-safe topic management
│   │   ├── CellTracker.cs                 # Tracks which cells are synced
│   │   └── Config.cs                      # Server URL, sync settings
│   │
│   └── SyncServer/                        # Central sync server
│       ├── SyncServer.csproj
│       ├── Program.cs                     # ASP.NET Core entry
│       ├── Hubs/
│       │   └── SyncHub.cs                 # SignalR hub for real-time sync
│       ├── Services/
│       │   └── StateManager.cs            # In-memory state of all synced cells
│       └── Models/
│           └── CellUpdate.cs              # Cell update message model
│
├── tools/
│   ├── Register.bat                       # Register COM add-in
│   ├── Unregister.bat                     # Unregister
│   └── Deploy.ps1                         # Full deployment script
│
└── docs/
    └── README.md                          # Setup and usage instructions
```

## Key Implementation Details

### 1. SyncAddin.cs - COM Add-in Entry Point
- Implements `IDTExtensibility2` for Excel add-in lifecycle
- Hooks `Application.SheetChange` event to capture user edits
- Filters events to only send changes for tracked/synced cells
- Prevents echo loops (ignores changes from incoming sync)

### 2. RtdServer.cs - Inbound Updates via RTD
- Implements `IRtdServer` interface
- ProgId: `"ExcelSync.RtdServer"`
- Receives updates from SyncClient and pushes to Excel
- Topic format: `=RTD("ExcelSync.RtdServer", "", "channelId", "cellKey")`

### 3. SyncClient.cs - WebSocket Communication
- SignalR client connecting to SyncServer
- Single persistent connection per Excel instance
- Handles reconnection with exponential backoff
- Methods: `SendUpdate(cellKey, value)`, `OnReceiveUpdate` event

### 4. CellTracker.cs - User-Defined Sync Cells
- Users mark cells for sync using: `=SYNC("myKey", A1)` formula
- Tracks mapping: cellKey → Excel Range
- Persists sync configuration in workbook custom properties

### 5. SyncHub.cs (Server) - Real-time Message Routing
- SignalR Hub handling all connected Excel clients
- Broadcasts updates to all clients except sender
- Maintains in-memory state for late joiners

### 6. StateManager.cs (Server) - Cell State Storage
- `ConcurrentDictionary<string, CellState>` for current values
- Optional: Redis/DB persistence for durability
- Provides initial state sync when client connects

## How Users Mark Cells for Sync

**Option A: Formula-based (Recommended)**
```
Cell A1: =SYNC("price", B1)
```
- `SYNC` is a UDF that registers B1 for sync under key "price"
- Returns the synced value (either local B1 or remote update)

**Option B: Right-click context menu**
- User right-clicks cell → "Sync this cell" → enters key name
- Add-in adds cell to tracking list

## Data Flow Example

1. **User A** types "100" in cell A1 (synced as "price")
2. **SyncAddin** catches `SheetChange` event
3. **CellTracker** confirms A1 is tracked under key "price"
4. **SyncClient** sends `{key: "price", value: "100", sender: "A"}` to server
5. **SyncHub** broadcasts to all clients except A
6. **User B's SyncClient** receives update
7. **RtdServer** updates topic "price" → Excel refreshes cell

## Files to Create

### Excel Add-in (ExcelSyncAddin/)
| File | Purpose | ~Lines |
|------|---------|--------|
| SyncAddin.cs | Add-in lifecycle, Excel events | 150 |
| RtdServer.cs | IRtdServer implementation | 180 |
| SyncClient.cs | SignalR WebSocket client | 120 |
| TopicManager.cs | RTD topic management | 100 |
| CellTracker.cs | Track synced cells | 80 |
| SyncUdf.cs | SYNC() formula function | 60 |
| Config.cs | Configuration settings | 40 |

### Sync Server (SyncServer/)
| File | Purpose | ~Lines |
|------|---------|--------|
| Program.cs | ASP.NET Core startup | 30 |
| SyncHub.cs | SignalR hub | 80 |
| StateManager.cs | In-memory state | 60 |
| CellUpdate.cs | Message model | 20 |

### Tools & Config
| File | Purpose |
|------|---------|
| Register.bat | COM registration script |
| Unregister.bat | COM unregistration |
| appsettings.json | Server configuration |

## Excel Usage

### Setup (one-time)
1. Install add-in (run Register.bat as admin)
2. Configure server URL in Excel: `Sync Settings → Server: wss://yourserver.com/sync`

### Marking cells for sync
```excel
=SYNC("revenue_q1", D5)     ' Sync cell D5 under key "revenue_q1"
=SYNC("total", SUM(A1:A10)) ' Sync a formula result
```

### Direct RTD (advanced)
```excel
=RTD("ExcelSync.RtdServer", "", "channel1", "revenue_q1")
```

## Build & Deployment

### Client (Excel Add-in)
1. Open `ExcelSyncAddin.sln` in Visual Studio
2. Build → Release → Any CPU
3. Run `tools\Register.bat` as Administrator
4. Restart Excel

### Server
1. `cd src/SyncServer`
2. `dotnet publish -c Release`
3. Deploy to VPS, configure reverse proxy (nginx) for WSS
4. Run: `dotnet SyncServer.dll`

## Verification Plan

1. **Local Test**
   - Run SyncServer locally
   - Open two Excel instances
   - Mark same cell key in both
   - Edit in one → verify update in other

2. **Remote Test**
   - Deploy server to VPS
   - Connect from two different machines
   - Verify sync works over internet

3. **Stress Test**
   - 10 clients, rapid edits
   - Verify no message loss, acceptable latency

## Configuration

### Client (ExcelSyncAddin.config)
```xml
<appSettings>
  <add key="ServerUrl" value="wss://localhost:5000/sync"/>
  <add key="ReconnectDelayMs" value="1000"/>
  <add key="RtdRefreshMs" value="100"/>
</appSettings>
```

### Server (appsettings.json)
```json
{
  "Urls": "http://*:5000",
  "Sync": {
    "MaxClientsPerChannel": 50,
    "MessageRetentionMinutes": 60
  }
}
```

## Summary
- **Architecture**: Unified COM Add-in (C#) with SignalR server
- **Sync Method**: User-defined via `=SYNC("key", cell)` formula
- **Communication**: WebSocket (SignalR) for low-latency bidirectional sync
- **Scale**: 5-20 distributed spreadsheets
- **Deployment**: Add-in DLL + ASP.NET Core server on VPS

## Implementation Order
1. Create project structure and solution files
2. Implement SyncServer with SignalR hub
3. Implement RtdServer (inbound updates)
4. Implement SyncClient (WebSocket connection)
5. Implement SyncAddin (Excel event hooks)
6. Implement CellTracker and SYNC UDF
7. Create registration scripts
8. Test locally with two Excel instances
9. Create deployment documentation
