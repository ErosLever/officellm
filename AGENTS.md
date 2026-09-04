# Office LLM Harness — Development Rules

## ⚠️ MANDATORY: Tests Must Pass Before Commit

**Before any commit, ALL automated tests must pass.** This is non-negotiable.

```bash
# Run all tests (C# unit + integration tests)
dotnet test tests/mcp-server.Tests/

# Build everything first
dotnet build src/mcp-server/
cd src/office-addin && npx webpack --mode production
```

If tests fail — **do not commit**. Fix the root cause, re-run, and only commit when green.

## Implementation Workflow (TDD)

When implementing new features or fixing bugs, follow this order:

1. **Architecture** — Understand the codebase, identify affected modules, plan the approach
2. **Tests first** — Write tests that express the expected behavior _before_ implementation
3. **Implementation** — Write the code to make those tests pass
4. **Verification** — Run the full test suite to confirm nothing is broken

```bash
# 1. Understand existing tests
rg "test|Test" tests/ --include='*.cs'

# 2. Add new tests (red)
#    Write failing tests first, verify they fail

# 3. Implement (green)
#    Write code to make tests pass

# 4. Verify full suite
dotnet test tests/mcp-server.Tests/ --verbosity normal
```

## Project Structure

```
src/
├── mcp-server/           # .NET 8 MCP server (hub)
│   ├── AppBuilder.cs     # HTTP app factory (testable)
│   ├── Program.cs        # Entry point + stdio transport
│   ├── Models/
│   │   ├── InstanceRegistry.cs  # Thread-safe instance tracking
│   │   ├── CommandStore.cs      # Command queue with wait/timeout
│   │   └── McpResponse.cs       # Standard MCP response envelope
│   └── Tools/
│       ├── OfficeTools.cs     # Tool implementations
│       └── McpToolEngine.cs   # Tool dispatch + definitions
├── office-addin/     # Office JS PowerPoint Add-in
│   ├── src/
│   │   ├── app.ts            # Main entry point
│   │   └── communication.ts  # MCP client (register, poll, heartbeat)
│   └── package.json          # webpack build
tests/
└── mcp-server.Tests/         # xUnit test suite
    ├── InstanceRegistryTests.cs
    ├── CommandStoreTests.cs
    ├── McpResponseTests.cs
    ├── McpToolEngineTests.cs
    └── HttpEndpointTests.cs      # WebApplicationFactory integration tests
specs/                        # Speckit specifications
```

## Available Endpoints

| Transport       | Endpoint                        | Purpose                                           |
| --------------- | -------------------------------- | -------------------------------------------------- |
| Streamable HTTP | `POST /mcp`                     | MCP protocol (initialize, tools/list, tools/call) |
| REST Bridge     | `GET /openapi.json`             | OpenAPI 3.0 spec for non-MCP clients              |
| REST Bridge     | `POST /api/{toolName}`          | Direct REST calls to any MCP tool                 |
| Swagger UI      | `GET /docs`                     | Interactive API documentation                     |
| Instance mgmt   | `POST /instances/register`      | Add-in registration                               |
| Instance mgmt   | `POST /instances/:id/heartbeat` | Keep-alive (every 10s)                            |
| Instance mgmt   | `GET /instances/:id/commands`   | Poll for pending commands                         |
| Instance mgmt   | `POST /instances/:id/result`    | Report command results                            |
| Health          | `GET /health`                   | Server health check                               |
| SignalR Hub     | `/hubs/commands`                | WebSocket real-time command delivery              |
| Stdio           | stdin/stdout                    | MCPo-compatible JSON-RPC transport                |

## Build Commands

```bash
# Build MCP server
dotnet build src/mcp-server/

# Build PowerPoint add-in
cd src/office-addin && npx webpack --mode production

# Run tests
dotnet test tests/mcp-server.Tests/

# Full pipeline (build + test)
dotnet build src/mcp-server/ && cd src/office-addin && npx webpack --mode production && cd ../../ && dotnet test tests/mcp-server.Tests/
```

## Running the Server

```bash
# Default port 3000
dotnet run --project src/mcp-server/

# Custom port
dotnet run --project src/mcp-server/ 8080
```

## Word JS API Rules

When working with the Word JavaScript API in the add-in:

- **`Word.run()` is lowercase** — `Word.run(async (context) => { ... })`, not `Word.Run()`.
- **`context.sync()` is required before reading properties** — After loading objects with `.load()`, you MUST call `await context.sync()` before accessing loaded properties. Failing to sync results in "property not loaded" errors.
- **changeTrackingMode pattern for mutations** — Word tracked changes use `document.changeTrackingMode`:
  ```typescript
  // Standard mutation pattern:
  // 1. Save current mode
  // 2. Set TrackMineOnly
  // 3. Perform mutation
  // 4. Restore original mode
  // 5. Return { tracked: true }
  const originalMode = context.document.changeTrackingMode;
  context.document.changeTrackingMode = Word.ChangeTrackingMode.trackMineOnly;
  // ... perform mutations ...
  context.document.changeTrackingMode = originalMode;
  await context.sync();
  ```
- **`changeTrackingMode` values**: `"Off"` | `"TrackAll"` | `"TrackMineOnly"` (requires WordApi 1.4+)
- **`getTextFrameOrNullObject` equivalent** — Word has no direct equivalent. Use `range.getHtml()` or `range.getText()` to read content. For null-safe patterns, check `range.isNullObject` after `context.sync()`.
- **Range-based operations** — Word operates on `Range` objects. Get the current selection via `context.document.getSelection()`, then manipulate it as a range.

### Tool Inventory (131 tools)

| Tool Name                            | Host       | Category        |
| ------------------------------------- | ---------- | --------------- |
| `office_get_active_apps`             | Shared     | Read            |
| `office_get_document_context`        | Shared     | Read            |
| `office_get_document_stats`          | Shared     | Read            |
| `office_batch_call`                  | Shared     | Batch           |
| `office_suggest_tools`               | Shared     | Discovery       |
| `powerpoint_get_deck_outline`        | PowerPoint | Read            |
| `powerpoint_get_slide`               | PowerPoint | Read            |
| `powerpoint_get_slide_image`         | PowerPoint | Read            |
| `powerpoint_get_shape_image`         | PowerPoint | Read            |
| `powerpoint_get_table`               | PowerPoint | Read            |
| `powerpoint_get_selection`           | PowerPoint | Read            |
| `powerpoint_get_speaker_notes`       | PowerPoint | Read            |
| `powerpoint_update_shape_text`       | PowerPoint | Write           |
| `powerpoint_update_shape_properties` | PowerPoint | Write           |
| `powerpoint_update_speaker_notes`    | PowerPoint | Write           |
| `powerpoint_add_textbox`             | PowerPoint | Write           |
| `powerpoint_add_image`               | PowerPoint | Write           |
| `powerpoint_add_table`               | PowerPoint | Write           |
| `powerpoint_delete_shape`            | PowerPoint | Write           |
| `powerpoint_add_slide`               | PowerPoint | Write           |
| `powerpoint_set_slide_layout`        | PowerPoint | Write           |
| `powerpoint_delete_slide`            | PowerPoint | Write           |
| `powerpoint_move_slide`              | PowerPoint | Write           |
| `powerpoint_duplicate_slide`         | PowerPoint | Write (cross-doc via `targetInstanceId`) |
| `word_get_outline`                   | Word       | Read            |
| `word_get_paragraphs`                | Word       | Read            |
| `word_get_selection`                 | Word       | Read            |
| `word_search`                        | Word       | Read            |
| `word_replace_text`                  | Word       | Write (tracked) |
| `word_insert_text`                   | Word       | Write (tracked) |
| `word_add_comment`                   | Word       | Write           |
| `word_delete_paragraph`              | Word       | Write (tracked) |
| `word_get_tracked_changes`           | Word       | Read            |
| `word_accept_all_changes`            | Word       | Write           |
| `word_reject_all_changes`            | Word       | Write           |
| `word_get_tables`                    | Word       | Read            |
| `word_insert_table`                  | Word       | Write (tracked) |
| `word_update_table_cell`             | Word       | Write (tracked) |
| `word_add_table_rows`                | Word       | Write           |
| `word_delete_table_row`              | Word       | Write           |
| `word_delete_table_column`           | Word       | Write           |
| `word_add_table_column`              | Word       | Write           |
| `word_merge_table_cells`             | Word       | Write           |
| `word_split_table_cell`              | Word       | Write           |
| `word_copy_table_structure`          | Word       | Write           |
| `word_set_table_format`              | Word       | Write           |
| `word_get_headers_footers`           | Word       | Read            |
| `word_set_header_footer`             | Word       | Write (tracked) |
| `word_replace_selection`             | Word       | Write (tracked) |
| `word_insert_image`                  | Word       | Write           |
| `word_apply_style`                   | Word       | Write           |
| `word_get_sections`                  | Word       | Read            |
| `word_insert_list`                   | Word       | Write (tracked) |
| `excel_get_workbook_map`             | Excel      | Read            |
| `excel_read_range`                   | Excel      | Read            |
| `excel_write_range`                  | Excel      | Write           |
| `excel_write_formula`                | Excel      | Write           |
| `excel_create_table`                 | Excel      | Write           |
| `excel_add_sheet`                    | Excel      | Write           |
| `excel_delete_sheet`                 | Excel      | Write           |
| `excel_rename_sheet`                 | Excel      | Write           |
| `excel_sort_range`                   | Excel      | Write           |
| `excel_filter_range`                 | Excel      | Write           |
| `excel_create_chart`                 | Excel      | Write           |
| `excel_get_charts`                   | Excel      | Read            |
| `excel_format_range`                 | Excel      | Write           |
| `excel_apply_conditional_formatting` | Excel      | Write           |
| `excel_create_pivottable`            | Excel      | Write           |
| `outlook_get_current_item`           | Outlook    | Read            |
| `outlook_summarize_thread`           | Outlook    | Read            |
| `outlook_draft_reply`                | Outlook    | Write (draft)   |
| `outlook_apply_category`             | Outlook    | Write           |
| `outlook_send_message`               | Outlook    | Write (gated)   |
| `outlook_get_user_profile`           | Outlook    | Read            |
| `outlook_get_master_categories`      | Outlook    | Read            |
| `outlook_create_category`            | Outlook    | Write           |
| `outlook_remove_categories`          | Outlook    | Write           |
| `outlook_display_new_message`        | Outlook    | Write (gated)   |
| `outlook_display_new_appointment`    | Outlook    | Write (gated)   |
| `outlook_get_attachments`            | Outlook    | Read            |
| `office_export_document`             | **All**    | Read (export)   |
| `word_find_replace`                  | Word       | Write (tracked) |
| `excel_freeze_panes`                 | Excel      | Write           |
| `excel_get_named_ranges`             | Excel      | Read            |
| `excel_add_named_range`              | Excel      | Write           |
| `excel_add_data_validation`          | Excel      | Write           |
| `excel_remove_data_validation`       | Excel      | Write           |
| `powerpoint_get_tags`                | PowerPoint | Read            |
| `powerpoint_set_tag`                 | PowerPoint | Write           |
| `powerpoint_delete_slides_by_tag`    | PowerPoint | Write           |
| `powerpoint_set_shape_fill`          | PowerPoint | Write           |
| `powerpoint_set_shape_line`          | PowerPoint | Write           |
| `powerpoint_set_shape_rotation`      | PowerPoint | Write           |
| `powerpoint_add_geometric_shape`     | PowerPoint | Write           |
| `powerpoint_add_line`                | PowerPoint | Write           |
| `powerpoint_insert_slides_from_file` | PowerPoint | Write           |
| `powerpoint_get_layouts`             | PowerPoint | Read            |
| `powerpoint_get_theme_colors`        | PowerPoint | Read            |
| `powerpoint_group_shapes`            | PowerPoint | Write           |
| `powerpoint_ungroup_shape`           | PowerPoint | Write           |
| `word_get_bookmarks`                 | Word       | Read            |
| `word_insert_bookmark`               | Word       | Write (tracked) |
| `word_delete_bookmark`               | Word       | Write           |
| `word_goto_bookmark`                 | Word       | Read            |
| `word_get_properties`                | Word       | Read            |
| `word_set_properties`                | Word       | Write           |
| `word_get_hyperlinks`                | Word       | Read            |
| `word_insert_hyperlink`              | Word       | Write (tracked) |
| `word_insert_footnote`               | Word       | Write (tracked) |
| `word_insert_endnote`                | Word       | Write (tracked) |
| `word_insert_field`                  | Word       | Write (tracked) |
| `word_get_content_controls`          | Word       | Read            |
| `word_insert_content_control`        | Word       | Write (tracked) |
| `word_get_formatting`                | Word       | Read            |
| `word_set_formatting`                | Word       | Write           |
| `word_get_comments`                  | Word       | Read            |
| `word_edit_comment`                  | Word       | Write           |
| `word_resolve_comment`               | Word       | Write           |
| `word_delete_comment`                | Word       | Write           |
| `word_reply_to_comment`              | Word       | Write           |
| `word_edit_reply`                    | Word       | Write           |
| `word_delete_reply`                  | Word       | Write           |
| `word_get_styles`                    | Word       | Read            |
| `word_modify_style`                  | Word       | Write           |
| `word_create_style`                  | Word       | Write           |
| `word_create_and_remap_style`        | Word       | Write           |
| `word_get_image`                     | Word       | Read            |
| `excel_protect_sheet`                | Excel      | Write           |
| `excel_unprotect_sheet`              | Excel      | Write           |
| `excel_set_page_layout`              | Excel      | Write           |
| `excel_get_page_layout`              | Excel      | Read            |

**Mutation modes by host**:

- **Word**: Tracked changes (`changeTrackingMode: "TrackMineOnly"`) — user accepts/rejects via Word Review ribbon or tracked change tools. No confirmation gate needed.
- **PowerPoint**: Direct write with undo group per `PowerPoint.run()` batch — no tracked changes API exists. Undo (Ctrl+Z) reverses the entire batch.
- **Excel**: Direct write with native undo (Ctrl+Z). No tracked changes API. Formula validation rejects invalid syntax before writing.
- **Outlook**: Drafts created in Drafts folder — NEVER auto-sent. `outlook_send_message` requires explicit confirmation token from Outlook task pane. Policy filters can block sends.

## Speckit Workflow

For feature development, follow the speckit workflow:

1. `/speckit.specify` — Refine the specification
2. `/speckit.plan` — Create implementation plan
3. `/speckit.tasks` — Generate actionable tasks
4. `/speckit.implement` — Execute implementation

See `specs/README.md` for phase details.

## Word JS API Quick Reference

```
changeTrackingMode: "Off" | "TrackAll" | "TrackMineOnly"  (WordApi 1.4+)

Mutation pattern:
  save mode → set TrackMineOnly → mutate → restore mode → return { tracked: true }

PowerPoint: NO tracked changes API. Direct-write with undo grouped per PowerPoint.run() batch.
```
