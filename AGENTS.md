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

### Tool Inventory (137 tools)

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
| `powerpoint_get_shape_paragraphs`    | PowerPoint | Read            |
| `powerpoint_get_shape_text_markdown` | PowerPoint | Read            |
| `powerpoint_get_selection`           | PowerPoint | Read            |
| `powerpoint_get_speaker_notes`       | PowerPoint | Read            |
| `powerpoint_update_shape_text`       | PowerPoint | Write           |
| `powerpoint_update_shape_properties` | PowerPoint | Write           |
| `powerpoint_update_text_range_properties` | PowerPoint | Write     |
| `powerpoint_insert_paragraph`        | PowerPoint | Write           |
| `powerpoint_delete_paragraph`        | PowerPoint | Write           |
| `powerpoint_set_shape_text_markdown` | PowerPoint | Write           |
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

`powerpoint_update_shape_properties` also supports font decoration (underline/strikethrough/allCaps/smallCaps/subscript/superscript), paragraph/bullet formatting (horizontalAlignment/indentLevel/bulletType/bulletStyle/bulletVisible), and text frame layout (margins/autoSizeSetting/wordWrap/verticalAlignment). `powerpoint_get_shape_paragraphs` and `powerpoint_update_text_range_properties` extend this to a single character sub-range (typically one paragraph) instead of the shape's whole text frame. `powerpoint_insert_paragraph` adds a brand-new paragraph (positioned via `start`/`end`/`before`/`after`) without touching any existing paragraph's text or formatting, optionally applying the same font/paragraph/bullet formatting params in the same call — if none are given, the new paragraph inherits the formatting of whichever paragraph it's spliced next to. `powerpoint_delete_paragraph` removes a paragraph identified by its `paragraphStart`/`paragraphLength` span (as returned by `powerpoint_get_shape_paragraphs`), consuming exactly one adjacent paragraph delimiter so the paragraph count actually drops by one — other paragraphs' text and formatting are untouched; deleting the last remaining paragraph just empties its text instead. `indentLevel`, `bulletType`, and `bulletStyle` require Office JS API set 1.10; confirmed working on Mac (live-tested via `powerpoint_get_shape_paragraphs`/`powerpoint_update_text_range_properties`), verify on other target platforms before relying on it there. `powerpoint_get_shape_text_markdown` renders a shape's paragraphs as an indented markdown-like list built on top of `powerpoint_get_shape_paragraphs`'s output: each paragraph becomes one line indented `"  " * indentLevel`, using `bulletChar` (`-` or `*`, default `-`) for `bulletType: "Unnumbered"` paragraphs, and for `bulletType: "Numbered"` paragraphs, the shape's *actual* bullet marker for the 16 Latin-representable `BulletStyle` values (Arabic numeral × plain/period/parenRight/parenBoth; alphabet lowercase/uppercase × period/parenRight/parenBoth; roman lowercase/uppercase × period/parenRight/parenBoth) — e.g. `a.`, `(iv)`, `III)` — counting contiguously per indent level and restarting the count whenever the indent level changes *or* the `bulletStyle` changes at the same level. The other 25 non-Latin/script-specific styles (and unset/unsupported styles) fall back to the previous `1.`/`2.` behavior. The rendering logic lives in a standalone exported function, `paragraphsToMarkdown`, backed by a `BULLET_STYLE_INFO` lookup table (style name → `{family, wrapper}`) plus `toAlphabetCounter`/`toRomanNumeral`/`formatBulletMarker` helpers — the same table and helpers the decoder below reuses, so encoder and decoder can't drift apart.

`powerpoint_set_shape_text_markdown` is the inverse: it replaces a shape's entire text frame from a markdown string in this same flavor, parsing each line into `{text, indentLevel, bulletType, bulletStyle, bulletVisible}` via the exported pure function `markdownToParagraphSpecs`, then writing the joined text and applying each paragraph's formatting via `applyRangeProperties` (the same helper `powerpoint_insert_paragraph` uses). `-`/`*` markers become `Unnumbered` bullets; everything else is classified into one of the 16 supported styles by unwrapping the marker (`plain`/`period`/`parenRight`/`parenBoth`) and inspecting the core token. Classification is sequence-aware to resolve the one genuinely ambiguous marker, `i`/`I` (Roman 1 vs. alphabetic position 9): it's alphabetic only when the immediately preceding numbered line at the same indent level (in the same contiguous run — severed by a bulleted line or a style change) was `h`/`H`; otherwise Roman. `v`/`V` and `x`/`X` are always Roman (never alphabetic position 22/24); `l`/`c`/`d`/`m` (any case) are always alphabetic (never Roman 50/100/500/1000). Any line or marker that doesn't parse — no marker+space separator, an unmatched wrapper shape, an unrecognized core, or a mixed-case multi-letter Roman token like `iV.` — is a hard error with no mutation, consistent with `powerpoint_insert_paragraph`/`powerpoint_delete_paragraph`'s validate-and-error style.

**Live Office.js enum-string quirk (confirmed on Mac desktop, PowerPointApi 1.10):** reading `paragraphFormat.horizontalAlignment`, `bulletFormat.type`, or `bulletFormat.style` off a live `TextRange` returns the 0-based index of the enum member's declaration order in `@types/office-js` as a numeric string (e.g. `"2"` for `BulletType.numbered`) instead of the named string (`"Numbered"`) the type declarations promise — verified by explicitly setting `bulletType: "Numbered"`/`"Unnumbered"`/`"None"` via `powerpoint_update_text_range_properties` and reading back `"2"`/`"3"`/`"1"` respectively. The mock test suite never caught this because its fixtures use the named strings directly. `extractRangeProperties` in `powerpoint-commands.ts` now normalizes these three fields back to their named form via `normalizeEnumString` (passes through anything that isn't a pure digit string, so mock/write-path values are unaffected) before they reach any getter output — this is why `powerpoint_get_shape_paragraphs`/`powerpoint_get_shape_text_markdown` always show `"Numbered"`/`"Left"`/etc., never raw digits.

**PowerPoint slide indices are 1-based**: every slide-index-shaped parameter (`slideIndex`, `atIndex`, `fromIndex`/`toIndex`, `targetIndex`, `startSlide`/`endSlide`, `insertAfterSlideIndex`, `slideIndexes`) counts from 1, matching the slide numbers shown in the PowerPoint UI. Internally, OfficeJS's `pres.slides.items`/`getItemAt` are 0-based, so `powerpoint-commands.ts` converts via `toZeroBasedSlideIndex`/`toOneBasedSlideIndex` only at the point of native indexing — everywhere else (config reads, result fields) the value stays 1-based.

**No PowerPoint comment tools**: Office.js has no Comment API for PowerPoint (`PowerPoint.Slide`/`Shape` expose no `comments` property, unlike `Word.Comment`/`CommentCollection` and `Excel.Comment`/`CommentCollection`). This is a platform limitation, not a missing feature in this repo — do not attempt to add `powerpoint_get_comments`/`add_comment`/etc. until Microsoft ships the underlying API.

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
