---
name: add-office-host-tools
description: Procedure for adding a new Office host (Word, Excel, Outlook) to the unified Office LLM Harness. Covers backend tool definitions, frontend command handlers, mock framework, tests, and wiring.
version: 8
created: 2026-05-16
updated: 2026-05-16
---
## When to Use

When adding a new Office application host (e.g., Outlook, OneNote) to the vsto project's MCP server + unified add-in architecture.

## Procedure

### 1. Backend — McpToolEngine.cs tool definitions

1. Add tool definition objects to the `GetToolDefinitions()` array (before the closing `];`). Follow the existing pattern: `name`, `description`, `inputSchema` with `type`, `properties`, `required`.
2. Add tool names to the `AddInCommands` HashSet (e.g., `"outlook_get_current_item"`).
3. Update tool count tests:
   - `McpToolEngineTests.GetToolDefinitions_ReturnsNTools` — update count and add `Assert.Contains` for each new tool name
   - `HttpEndpointTests.Mcp_ToolsList_ReturnsNTools` — update count
4. Run `dotnet build src/mcp-server/` and `dotnet test tests/mcp-server.Tests/`

### 2. Frontend — Command handlers

1. Create `src/powerpoint-addin/src/{host}-commands.ts`
   - Use `runIn{Host}<T>()` helper for Excel.run()/Word.run() patterns, or callback-based API for Outlook
   - Use `globalThis` not `window` for API access
   - Export `processCommand()` that switches on command name
2. Create `src/powerpoint-addin/src/{host}-mock.ts`
   - All mock objects that the handler calls `.load()` on must have a `load()` method
   - Table `columns` must be `{ items: [...], load() {} }` not a plain array
   - Use `globalThis` for install/uninstall, set `(globalThis as any).window = globalThis`
   - Export: `installXMock`, `uninstallXMock`, `setWorkbookData`/`setMockData`, `getMutations`, `resetMutations`
3. Create `src/powerpoint-addin/src/{host}-commands.test.ts`
   - Import from mock, NOT the class (e.g., `ExcelMock` is NOT exported, use `installExcelMock`)
4. Wire into `app.ts`:
   - Add import: `import { processCommand as process{Host}Command } from "./{host}-commands"`
   - Add to `HOST_DISPATCH`: `{host}_: process{Host}Command`
   - **WARNING**: Use `append` after existing import line, not `replace` — replacing loses the adjacent import

### 3. Verification

```bash
dotnet build src/mcp-server/ && cd src/powerpoint-addin && npx webpack --mode production && cd ../../ && dotnet test tests/mcp-server.Tests/
cd src/powerpoint-addin && npx vitest run
```

## Pitfalls
## Pitfalls

- **`window` vs `globalThis`**: Mock files and command handlers must use `(globalThis as any).Excel` / `.Word` / `.Office`, NOT `(window as any)`. Node.js test env (vitest) has no `window`. Also set `(globalThis as any).window = globalThis` in mock install.
- **Mock `load()` methods**: Every mock object that the real API calls `.load()` on must include a no-op `load() {}` method — worksheets, ranges, columns collections, tables, table ranges.
- **Mutations tracking**: The `_mutations` array must live in the MOCK (not commands). Export `_getMutationsArray()` from mock, import in commands to push. Tests import `getMutations`/`resetMutations` from mock. Do NOT create a separate `_mutations` in the commands file.
- **Tool count in 3 places**: After adding N tools, update count in: (1) McpToolEngineTests test name + Assert.Equal, (2) HttpEndpointTests test name + Assert.Equal, (3) AGENTS.md tool inventory table header + rows.
- **Outlook is callback-based**: Outlook JS API uses `Office.context.mailbox.item` with callback pattern (getAsync(callback)), NOT `Excel.run()`/`Word.run()` promise pattern.
- **Biome auto-fix**: After write, biome may convert `import` to `import type` — must re-read and fix back to regular import if the symbol is used as a value.
- **Duplicate closing brackets**: When inserting tool definitions into the array in McpToolEngine.cs, ensure old closing `]; } ];` is replaced, not duplicated.
- **Edit anchor ranges**: When editing import sections in app.ts, ensure the anchor range covers exactly the intended lines — replacing a range starting at one import can delete adjacent imports.