# Office LLM Harness (fork)

> **This is a fork of [volkermauel/officellm](https://github.com/volkermauel/officellm).**
> Changes introduced in this fork are listed in the [Fork Changes](#fork-changes) section below.

A cross-platform (macOS + Windows) Office add-in harness that exposes controlled document interaction tools to any MCP-compatible LLM client through a local MCP server, enabling LLM-assisted workflows in Word, Excel, PowerPoint and Outlook.

## Architecture

```
+--------------------+        +-------------------------+
| Word / Excel /     |        | Open WebUI              |
| PowerPoint /       |        | - Chat UI               |
| Outlook            |        | - Model routing         |
+---------+----------+        | - MCP external tools    |
          | Office JS API                  |
          v                               | MCP Streamable HTTP
+---------+----------+                    |
| Office JS Add-in   |                    |
| - Task pane UI     |                    |
| - Office API       |                    |
| - Confirmation UI  |                    |
+---------+----------+                    |
          | localhost HTTP                  |
          v                               |
+---------+-------------------------------+-------------+
| Local Office MCP Server (.NET 8)                      |
| - MCP protocol endpoint (port 3000)                   |
| - Tool registry & command dispatch                    |
| - Instance registry with heartbeat tracking           |
| - Bearer-token authentication (all endpoints)         |
| - Audit log (JSONL)                                   |
+------------------------------------------------------+
```

### Data Flow

1. **LLM calls tool** → `POST /mcp` with `tools/call` method
2. **MCP server queues command** → stored in `CommandStore`, add-in polls `GET /instances/{id}/commands`
3. **Add-in executes** → `PowerPoint.run()` calls Office JS API
4. **Add-in reports result** → `POST /instances/{id}/result`
5. **MCP server returns** → result passed back to LLM

For **mutation tools** (`update_shape_text`, `update_speaker_notes`), step 2–4 is replaced by a two-phase confirmation flow with diff preview.

## Components

| Component                 | Language        | Description                                                                                                                               |
| ------------------------- | --------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| **MCP Server**            | C# (.NET 8)     | Self-contained executable exposing MCP tools over Streamable HTTP. Command dispatch, instance registry, confirmation flow, audit logging. |
| **Unified Office Add-in** | TypeScript/HTML | Single Office JS Add-in that auto-detects host (Word/Excel/PowerPoint/Outlook) via `Office.onReady()`. One manifest for all hosts.        |
| **Express Server**        | Node.js         | Serves static add-in files + dynamic `manifest.xml` (URLs from Host header). Docker/K8s deployment.                                       |

## Project Structure

```
src/
├── mcp-server/           # .NET 8 MCP server
│   ├── OfficeMcpServer.csproj
│   ├── Program.cs        # Entry point, MCP endpoints, bridge server
│   ├── Models/
│   │   └── McpResponse.cs
│   └── Tools/
│       └── OfficeTools.cs
├── powerpoint-addin/     # Unified Office JS Add-in (all hosts)
│   ├── manifest.xml      # Unified manifest (Presentation + Document + Workbook + Mailbox)
│   ├── package.json
│   ├── webpack.config.js
│   ├── tsconfig.json
│   └── src/
│       ├── index.html    # Task pane UI
│       ├── app.ts        # Main entry point
│       └── communication.ts # Office API wrappers + HTTP client
scripts/
├── build.sh              # Build script (all/mcp/addin/dev)
└── dev.sh                # Development server
specs/                    # Speckit specifications
├── README.md
├── 001-spike/
├── 002-powerpoint-mvp/
├── 003-word-mvp/
├── 004-excel-mvp/
└── 005-outlook-mvp/
```

## Quick Start

### Prerequisites

- **macOS or Windows** with Office desktop (Microsoft 365 recommended; Office 2019+ minimum)
- **.NET 8 SDK** (for building the MCP server)
- **Node.js 18+** and **npm** (for building the add-in)

### Build

**First-time setup — generate the API token before building:**

```bash
node scripts/generate-token.mjs
```

This writes `.env` and `src/mcp-server/GeneratedToken.cs` (both gitignored). The same token is compiled into the MCP server binary and the add-in JS bundle. Re-run to rotate the token; then rebuild both components.

```bash
# Build everything (production)
./scripts/build.sh

# Build MCP server only
./scripts/build.sh mcp

# Build add-in only
./scripts/build.sh addin

# Development mode (add-in with hot reload)
./scripts/build.sh dev
```

### Run

```bash
# 1. Start the MCP server
dotnet run --project src/mcp-server/
# On first run the server prints the token file location, e.g.:
#   Token file: /Users/you/Library/Application Support/OfficeMcpServer/token  (macOS)
#   Token file: C:\Users\you\AppData\Roaming\OfficeMcpServer\token             (Windows)

# 2. Configure your MCP client (e.g. Claude Desktop) — see "MCP Client Setup" below

# 3. Sideload the Office add-in — see "Add-in Sideloading" below

# 4. Open Word / Excel / PowerPoint and activate the add-in from the task pane
```

### MCP Client Setup

The server requires a Bearer token on every request. After building, the token is written to:

| Platform | Path |
|---|---|
| macOS | `~/Library/Application Support/OfficeMcpServer/token` |
| Linux | `~/.config/OfficeMcpServer/token` |
| Windows | `%APPDATA%\OfficeMcpServer\token` |

**Claude Desktop** — add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "office": {
      "url": "http://127.0.0.1:3000/mcp",
      "headers": {
        "Authorization": "Bearer <paste token here>"
      }
    }
  }
}
```

Read the token with:
- macOS: `cat ~/Library/Application\ Support/OfficeMcpServer/token`
- Linux: `cat ~/.config/OfficeMcpServer/token`
- Windows: `type %APPDATA%\OfficeMcpServer\token`

### Add-in Sideloading

The single manifest (`src/powerpoint-addin/manifest.xml`) works for all Office hosts (Word, Excel, PowerPoint, Outlook).

**macOS** — `office-addin-debugging` does not support Word/Excel/PowerPoint on macOS.
Use the watched-folder method:

1. Start the dev server (generates `manifest.dev.xml`):
   ```bash
   cd src/powerpoint-addin && npm run dev
   ```

2. Copy the manifest into the Office app's watched folder:
   ```bash
   # Word
   mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
   cp src/powerpoint-addin/manifest.dev.xml \
      ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/

   # Excel
   mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
   cp src/powerpoint-addin/manifest.dev.xml \
      ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/

   # PowerPoint
   mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
   cp src/powerpoint-addin/manifest.dev.xml \
      ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
   ```

3. Quit the Office app completely (Cmd+Q) and reopen it.

4. The add-in appears as a button in the ribbon. In Word it shows up in the **Home** tab
   (not **Insert → Add-ins**; the "Add-ins" entry in the Developer menu opens the VBA
   dialog and will not work for Office JS add-ins).

> Note: `Insert → Add-ins → My Add-ins → Upload My Add-in` is present in some M365
> builds but hidden in others; the watched-folder method works universally.

**Windows** — shared folder sideloading:

1. Open any Office app → **File → Options → Trust Center → Trust Center Settings**
2. Select **Trusted Add-in Catalogs** → add the folder containing `manifest.xml`
3. Restart Office and insert the add-in from **Insert → My Add-ins**

## Development

### Office JS Add-in

```bash
cd src/powerpoint-addin
npm install
npm run dev    # Generates token if missing, then starts webpack dev server on port 3001
npm run build  # Generates token if missing, then produces production bundle
```

### MCP Server

```bash
cd src/mcp-server
dotnet restore
dotnet run     # Pre-build generates token if missing; server binds to 127.0.0.1:3000

# Production — build for your platform:
dotnet publish -c Release -r osx-arm64 --self-contained true  # macOS Apple Silicon
dotnet publish -c Release -r osx-x64  --self-contained true  # macOS Intel
dotnet publish -c Release -r win-x64  --self-contained true  # Windows
```

## Development Notes

These are hard-won lessons from building this project. Follow them to avoid known pitfalls.

### Office JS API

- **`PowerPoint.run()` is lowercase** — not `Run()`, `Excel.Run()`, etc. The Office JS API uses camelCase (`run`, not `Run`). A typo here silently fails with "PowerPoint.run() not available" because `PowerPoint.Run` is `undefined`.
- **Use slash-separated load paths** — `slide.load("shapes/items/id,name,textFrame/textRange/text")` loads nested collections and their properties in one call. This is the official pattern from Microsoft docs.
- **`presentation.load("slides")`** — loads the slides navigation property. After sync, `slides.items` is available.
- **`shape.textFrame` throws `InvalidArgument`** if the shape has no text frame. Always wrap in try/catch.
- **`context.load(collection, ["items"])` is INVALID** — use `presentation.load("slides")` or `slide.load("shapes/items/$none")` instead.
- **`@types/office-js` is incomplete** — many newer PowerPoint context types lack type definitions. Use `any` casts and `PowerPoint.run(async (context: any) => ...)`.
- **`Office.context.document.url`** — gives the document URL/path (if available). Use this for the real document name in `getOfficeState()`.
- **Official API reference**: https://learn.microsoft.com/en-us/javascript/api/powerpoint?view=powerpoint-js-preview

### MCP Protocol

- **`params.arguments`** — tool call arguments come under `params.arguments`, NOT `params.input`. The MCP spec uses `arguments`.
- **`inputSchema`** — tool definitions use `inputSchema` (not `parameters` or `schema`). The `properties` and `required` fields follow JSON Schema.
- **`@default` in C#** — JSON Schema uses `default` as the keyword. In C#, use `@default` (with `@` prefix since `default` is a reserved word).

### Add-in Architecture

- **Command polling must process results** — `startCommandPolling()` must call `processPendingCommands()`, not just `pollForCommands()`. The poll returns commands that must be dispatched to `processCommand()`.
- **`processCommand()` calls `reportResult()` internally** — don't double-report. The handler already POSTs the result back to the MCP server.
- **`Office.onReady()` can fire multiple times** — use an `isInitialized` guard to prevent double-registration.
- **No duplicate `<script>` tags** — `HtmlWebpackPlugin` already injects `bundle.js`. Adding a static `<script src="bundle.js">` causes double initialization.

### ASP.NET / MCP Server

- **Suppress noisy request logging** — polling endpoints fire every 2s. Set `Microsoft.AspNetCore` logging to `Warning` level:
  ```csharp
  builder.Logging.AddFilter("Microsoft.AspNetCore", LogLevel.Warning);
  ```
- **`CleanupTimedOut()` must remove instances** — don't just set `IsAlive = false` or the same dead instance logs "timed out" every 30 seconds. Actually remove it from the dictionary.
- **`Results.Json()` serializes anonymous types correctly** — `Dictionary<string, object>` with anonymous type values works fine with `System.Text.Json`.

### Deployment

- **Manifest uses `{{BASE_URL}}` placeholders** — replaced at request time by Express server. No hardcoded URLs.
- **`<AppDomain>` not `<Domain>`** — Office manifest XML uses `<AppDomain>`, not `<Domain>`.
- **`<Host Name="Presentation"/>`** — the `Name` attribute is capitalized.
- **No `<RequestedWidth>`** on TaskPaneApp — only valid for Content app types.
- **Traefik ingress** — use `traefik` ingress class, not `nginx`.
- **`{{ }}` in Helm** — no spaces inside braces. Some editors auto-format and insert `{ { } }` which breaks templates.

## Testing

### MCP Server (C#)

58 xUnit tests covering models, command routing, and HTTP endpoints:

```bash
dotnet test tests/mcp-server.Tests/
```

- Uses `WebApplicationFactory<Program>` for in-process HTTP integration tests.
- `McpToolEngine.ResetForTesting()` clears static state before each test class.
- Pre-commit hook (`.git/hooks/pre-commit`) enforces all tests pass.

### Express Server (Node.js)

6 tests for dynamic manifest generation:

```bash
cd server && npm test
```

### Add-in (TypeScript)

The add-in must be tested in a real Windows Office environment:

1. Build both components
2. Start the MCP server
3. Sideload the PowerPoint add-in
4. Verify task pane loads in PowerPoint
5. Test each tool via Open WebUI's MCP integration

## CI/CD

| Workflow      | Trigger             | Jobs                                                                                               |
| ------------- | ------------------- | -------------------------------------------------------------------------------------------------- |
| `ci.yml`      | Push to master/main | Build+Test → Docker push (`ghcr.io/volkermauel/officellm-static:latest`) → Windows `.exe` artifact |
| `release.yml` | Tag `v*`            | GitHub Release with `.exe` + Docker semver tag                                                     |

Download latest exe:

```bash
gh run download --name office-mcp-server-win-x64
```

## Specifications

See [`specs/`](specs/) for detailed feature specifications organized by implementation phase:

- [Phase 0: Spike](specs/001-spike/) - Minimal MCP server + PowerPoint add-in
- [Phase 1: PowerPoint MVP](specs/002-powerpoint-mvp/) - Full PowerPoint tool set
- [Phase 2: Word MVP](specs/003-word-mvp/) - Word tools + shared context
- [Phase 3: Excel MVP](specs/004-excel-mvp/) - Excel read/write tools
- [Phase 4: Outlook MVP](specs/005-outlook-mvp/) - Email tools + policy filter

## Fork Changes

This fork diverges from [volkermauel/officellm](https://github.com/volkermauel/officellm) in the following ways.

### macOS support

The upstream targets Windows exclusively (`RuntimeIdentifier: win-x64`, Windows-only sideloading instructions). This fork:

- Documents and tests `dotnet publish` targets for `osx-arm64` and `osx-x64`
- Uses `Environment.SpecialFolder.ApplicationData` for all file paths (resolves to `~/Library/Application Support` on macOS, `%APPDATA%` on Windows) — already correct in the upstream `AuditLog.cs`, now consistently used in the new token infrastructure
- Replaces the upstream `office-addin-debugging` sideloading instructions (CLI does not support macOS Office apps) with the correct watched-folder and Insert→Add-ins method

### Bearer-token authentication

The upstream server has **no authentication** — any local process can call any endpoint and read or modify any open document.

This fork adds compile-time token authentication across the entire server:

| File | Change |
|---|---|
| `scripts/generate-token.mjs` | **New.** Generates a 32-byte random token, writes `.env` (for webpack) and `src/mcp-server/GeneratedToken.cs` (for C#). Idempotent. |
| `.gitignore` | Added `.env` and `GeneratedToken.cs` — the token is never committed. |
| `src/mcp-server/OfficeMcpServer.csproj` | Added `<Compile Include="GeneratedToken.cs"/>` and a pre-build `<Target>` that auto-runs the generate script if the file is missing. |
| `src/mcp-server/AppBuilder.cs` | Added ASP.NET middleware that checks `Authorization: Bearer <token>` on every request. `?access_token=` is accepted as fallback for SignalR WebSocket connections (browsers cannot set WebSocket headers). `/manifest.xml` is the only exempt path (fetched by the Office host before the add-in has a token). On startup, `WriteTokenFile()` writes the compiled-in token to the platform config dir (`~/Library/Application Support/OfficeMcpServer/token` on macOS, `%APPDATA%\OfficeMcpServer\token` on Windows) with chmod 600 on Unix. |
| `src/powerpoint-addin/webpack.config.js` | Reads token from `.env`, injects it as `__MCP_TOKEN__` via `webpack.DefinePlugin`. Fails the build if `.env` is missing. |
| `src/powerpoint-addin/src/communication.ts` | Declares `__MCP_TOKEN__`, exposes `authHeaders()` helper, spreads auth header into every `fetch` call, appends `?access_token=` to the SignalR hub URL. |
| `src/powerpoint-addin/src/globals.d.ts` | **New.** TypeScript ambient declaration for `__MCP_TOKEN__`. |
| `src/powerpoint-addin/package.json` | Added `generate-token` script; `build` and `dev` scripts call it automatically. |

**Token rotation:** delete `.env` and `src/mcp-server/GeneratedToken.cs`, re-run `node scripts/generate-token.mjs`, then rebuild both components. The old token immediately stops working.

**External MCP clients** (Claude Desktop, etc.) read the token from the platform path above (written by the server on first startup) and include it as `Authorization: Bearer <token>` — see [MCP Client Setup](#mcp-client-setup).

### Pending security hardening (not yet implemented)

The following findings from an internal security review remain open:

- **Instance ID predictability** — IDs are sequential (`word_1`, `word_2`). Should use random UUIDs.
- **SignalR group join unauthenticated** — any connected WebSocket can join any instance group. Needs a per-instance join secret.
- **Result injection** — `POST /instances/{id}/result` does not verify the submitter owns the command. Needs ownership check in `CompleteCommand`.
- **`office_export_document` has no confirmation gate** — exports the full document binary with no user prompt.
- **`outlook_send_message` token validation is prefix-only** — `confirmationToken.startsWith("confirm_")` accepts any crafted string.

## License

Private / Internal Use Only
