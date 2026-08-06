# Office LLM Harness — PowerPoint Add-in

## Local Development & Testing

### Prerequisites

- **PowerPoint** (Microsoft 365 or PowerPoint 2019+) on Windows or Mac
- **Node.js 18+** and `npm`
- **MCP Server** running on port 3000

### Quick Start

```bash
# 1. Start the MCP server (port 3000)
cd /home/volker/git/vsto/src/mcp-server
dotnet run

# 2. In a second terminal, start the webpack dev server (port 8080)
cd /home/volker/git/vsto/src/office-addin
npx webpack serve --mode development

# 3. Open PowerPoint and load the add-in
#    → File > Options > Trust Center > Trust Center Settings > Trusted Add-in Paths
#    → Add: /home/volker/git/vsto/src/office-addin/manifest.xml
#    OR
#    → File > Add-ins > My Add-ins > Manage Add-ins > Browse
#    → Select: /home/volker/git/vsto/src/office-addin/manifest.xml
```

### Architecture

```
┌─────────────────┐     HTTP      ┌──────────────────┐
│  PowerPoint      │◄────────────►│  MCP Server      │
│  Add-in (TS)    │  Port 8080   │  Port 3000       │
│                 │              │                  │
│  - Registers     │◄────────────►│  - Instance mgmt │
│  - Polls cmds    │  Port 8080   │  - Command queue │
│  - Reports res   │              │  - Tool routing  │
└─────────────────┘              └──────────────────┘
```

### Key URLs

| Service            | URL                                  | Purpose                          |
| ------------------ | ------------------------------------ | -------------------------------- |
| Webpack Dev Server | `http://localhost:8080`              | Task pane UI (index.html)        |
| MCP Server         | `http://127.0.0.1:3000`              | Command routing & tool execution |
| MCP Swagger UI     | `http://127.0.0.1:3000/docs`         | API documentation                |
| OpenAPI Spec       | `http://127.0.0.1:3000/openapi.json` | Machine-readable API spec        |
| Health Check       | `http://127.0.0.1:3000/health`       | Server status                    |

### Manifest Configuration

The manifest (`manifest.xml`) is configured for local development:

- **Task pane URL**: `http://localhost:8080/index.html`
- **Icons**: Served from `http://localhost:8080/assets/icon-32.png` and `icon-64.png`
- **MCP Server**: Connects to `http://127.0.0.1:3000`
- **Allowed domains**: `localhost`, `127.0.0.1`

### Testing the Add-in

1. **Registration**: Open PowerPoint → add-in should auto-register with MCP server
2. **Heartbeat**: Check MCP server logs for heartbeat messages every 10s
3. **Context Refresh**: Click "Refresh Context" button in task pane
4. **Command Polling**: Send a command via MCP server → add-in polls and executes it
5. **Diff Preview**: Trigger a mutation tool → diff preview appears in task pane
6. **Confirmation**: Click Approve/Reject → change is applied or rejected

### Troubleshooting

| Issue                    | Solution                                                                    |
| ------------------------ | --------------------------------------------------------------------------- |
| Port 3000 already in use | Kill existing process: `lsof -i :3000` then restart MCP server              |
| Add-in not loading       | Check manifest path is correct; verify `http://localhost:8080` is reachable |
| Commands not executing   | Verify MCP server is running; check browser console (F12) for errors        |
| CORS errors              | Manifest includes `localhost` and `127.0.0.1` in AppDomains                 |
| Icons not showing        | Run `npx webpack` to copy assets to dist/                                   |

### Building for Production

```bash
cd /home/volker/git/vsto/src/office-addin
npx webpack --mode production
```

This produces:

- `dist/bundle.js` — minified JavaScript
- `dist/index.html` — task pane HTML
- `dist/assets/icon-*.png` — icons
- `manifest.xml` — updated manifest (copied to parent dir)

For production deployment, you'll need:

1. HTTPS endpoint for the task pane URL
2. Valid SSL certificate (Office requires HTTPS in production)
3. Updated manifest with production URLs
4. Icon URLs pointing to HTTPS endpoints
