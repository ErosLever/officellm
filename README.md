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
├── mcp-server/           # .NET MCP server
│   ├── OfficeMcpServer.csproj
│   ├── Program.cs        # Entry point
│   ├── AppBuilder.cs     # CORS, auth middleware, server config
│   ├── Models/
│   │   ├── InstanceRegistry.cs
│   │   └── CommandStore.cs
│   └── Tools/
│       └── McpToolEngine.cs  # 129 tool definitions + dispatch
├── office-addin/     # Unified Office JS Add-in (all hosts)
│   ├── manifest.xml      # Unified manifest (Presentation + Document + Workbook + Mailbox)
│   ├── package.json
│   ├── webpack.config.js
│   ├── tsconfig.json
│   └── src/
│       ├── index.html        # Task pane UI (host-adaptive)
│       ├── app.ts            # Main entry, command polling, context display
│       ├── communication.ts  # MCP registration, heartbeat, instance ID derivation
│       ├── word-commands.ts  # Word tool handlers (129 tools)
│       └── globals.d.ts      # __MCP_TOKEN__ ambient declaration
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

The single manifest (`src/office-addin/manifest.xml`) works for all Office hosts (Word, Excel, PowerPoint, Outlook).

**macOS** — `office-addin-debugging` does not support Word/Excel/PowerPoint on macOS.
Use the watched-folder method:

1. Start the dev server (generates `manifest.dev.xml`):
   ```bash
   cd src/office-addin && npm run dev
   ```

2. Copy the manifest into the Office app's watched folder:
   ```bash
   # Word
   mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
   cp src/office-addin/manifest.dev.xml \
      ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/

   # Excel
   mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
   cp src/office-addin/manifest.dev.xml \
      ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/

   # PowerPoint
   mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
   cp src/office-addin/manifest.dev.xml \
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
cd src/office-addin
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

92 xUnit tests covering models, command routing, and HTTP endpoints:

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
| `src/office-addin/webpack.config.js` | Reads token from `.env`, injects it as `__MCP_TOKEN__` via `webpack.DefinePlugin`. Fails the build if `.env` is missing. |
| `src/office-addin/src/communication.ts` | Declares `__MCP_TOKEN__`, exposes `authHeaders()` helper, spreads auth header into every `fetch` call, appends `?access_token=` to the SignalR hub URL. |
| `src/office-addin/src/globals.d.ts` | **New.** TypeScript ambient declaration for `__MCP_TOKEN__`. |
| `src/office-addin/package.json` | Added `generate-token` script; `build` and `dev` scripts call it automatically. |

**Token rotation:** delete `.env` and `src/mcp-server/GeneratedToken.cs`, re-run `node scripts/generate-token.mjs`, then rebuild both components. The old token immediately stops working.

**External MCP clients** (Claude Desktop, etc.) read the token from the platform path above (written by the server on first startup) and include it as `Authorization: Bearer <token>` — see [MCP Client Setup](#mcp-client-setup).

### Stable human-readable instance IDs

The upstream uses sequential numeric IDs (`word_1`, `word_2`, …) — predictable and meaningless to an LLM.

This fork generates IDs client-side in the add-in before registration, using the format:

```
{host}_{docSlug}_{hash6}
```

Examples: `word_report_a3f2c1`, `excel_budget_7b9e04`, `ppt_untitled_cc1d88`

- **`host`** — short host prefix (`word`, `excel`, `ppt`, `outlook`)
- **`docSlug`** — document name, lowercased, extension stripped, non-alphanumeric collapsed to `_`, max 16 chars
- **`hash6`** — 6-hex-char FNV-1a hash of the document file URL, stable across reloads of the same saved file. For unsaved documents a random suffix is generated once and persisted in `localStorage`, so it also survives reloads within the same session.

The server validates the proposed ID against `[a-z][a-z0-9_]{2,63}` and falls back to `office_<uuid>` if it is missing or invalid. Re-registering with the same ID refreshes the heartbeat in place (no duplicate entries).

**For LLM clients:** always call `office_get_active_apps` first — it returns current instance IDs with document names. Never hardcode an instance ID.

### Word formatting and comment management (Phase 19)

Seven new tools for reading and writing Word paragraph/font formatting and managing comment threads:

| Tool | Description |
|---|---|
| `word_get_formatting` | Returns paragraph and font formatting for one or more paragraphs (style, alignment, indents, line spacing, font name/size/bold/italic/color/etc.) |
| `word_set_formatting` | Applies any subset of paragraph and font formatting to a paragraph range or the current selection |
| `word_get_comments` | Returns all comment threads with top-level comment metadata (id, author, date, text, resolved, anchorText) and replies |
| `word_edit_comment` | Edits the text of an existing top-level comment |
| `word_resolve_comment` | Resolves or reopens a comment thread |
| `word_delete_comment` | Permanently deletes a comment thread and all its replies |
| `word_reply_to_comment` | Adds a new reply to an existing comment thread |
| `word_edit_reply` | Edits the text of a specific reply |
| `word_delete_reply` | Deletes a single reply from a comment thread |

Comment IDs are stable within a document session — always call `word_get_comments` first to obtain IDs before calling edit/resolve/delete tools.

### Word table management tools

Eight new tools for structural editing and formatting of Word tables, extending the original three (`word_get_tables`, `word_insert_table`, `word_update_table_cell`):

| Tool | Description |
|---|---|
| `word_add_table_rows` | Adds one or more rows, as a matrix or as objects mapped via a `headers` array |
| `word_delete_table_row` | Deletes a row by index |
| `word_add_table_column` | Adds a column; auto-adjusts outer borders if the table uses an outer-border-only pattern |
| `word_delete_table_column` | Deletes a column with the same border adjustment |
| `word_merge_table_cells` | Merges a rectangular cell range |
| `word_split_table_cell` | Splits a cell into a rowCount × columnCount grid |
| `word_copy_table_structure` | Clones a table's column count, borders, padding, and column widths into a new empty table |
| `word_set_table_format` | Sets header row count, table style, alignment, cell padding, column widths, row height, and per-column paragraph/font overrides — all parameters optional |

### Word style management tools

Four tools for inspecting and defining named styles:

| Tool | Description |
|---|---|
| `word_get_styles` | Returns all styles with font/paragraph properties; filterable by type or `inUseOnly` |
| `word_modify_style` | Updates an existing style's definition document-wide |
| `word_create_style` | Creates a new named style, optionally inheriting from a base style |
| `word_create_and_remap_style` | Clones a style with overrides and remaps existing paragraphs from the base style to the new one — useful for customizing a built-in style (e.g. `Heading 1`) without altering its definition |

Note: list numbering cannot be baked into styles via Office JS on macOS.

### word_get_image — visual analysis of inline images

Retrieves an inline image from a Word document as a base64 data URL so an LLM can inspect it visually. Indexed by position among the document's inline images (0-based); only inline images are supported — floating pictures, text boxes, and other floating shapes have no equivalent export method in the Word JS API.

**Cropping caveat:** the Word JS API exposes no crop information for inline pictures, so the returned image is always the full original picture as stored in the document, even if the user cropped it in Word — it may show more than what's visible. The response includes `displayWidth`/`displayHeight` (the visible frame's dimensions); compare their aspect ratio to the returned image's actual pixel aspect ratio to detect a mismatch, which indicates the picture has been cropped.

### SSE transport compatibility

Added a legacy Server-Sent Events transport (`GET /sse` + `POST /mcp?sessionId=<id>`) alongside the Streamable HTTP transport, for MCP clients that only support the older SSE-based protocol (e.g. Claude Desktop ≤ 0.7). The SSE handler opens a stream, emits an `endpoint` event with the session-scoped POST path, then forwards all JSON-RPC responses for that session over the stream instead of the HTTP response body.

### word_search fix

`word_search` now returns each match's `paragraphIndex` and a short context snippet (marked with `[brackets]` around the match), instead of just raw match text — fixing a regression where results could contain duplicate or fabricated matches.

### Relicensing (MIT)

The upstream repository has no license file. With the original author's written permission, this fork is relicensed under the MIT License — see [LICENSE.md](LICENSE.md).

### Pending security hardening (not yet implemented)

The following findings from an internal security review remain open:

- **SignalR group join unauthenticated** — any connected WebSocket can join any instance group. Needs a per-instance join secret.
- **Result injection** — `POST /instances/{id}/result` does not verify the submitter owns the command. Needs ownership check in `CompleteCommand`.
- **`office_export_document` has no confirmation gate** — exports the full document binary with no user prompt.
- **`outlook_send_message` token validation is prefix-only** — `confirmationToken.startsWith("confirm_")` accepts any crafted string.

## License

MIT — see [LICENSE.md](LICENSE.md).

The upstream [volkermauel/officellm](https://github.com/volkermauel/officellm) has no license file. The original author granted written permission to relicense this fork under the MIT License.
