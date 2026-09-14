# Specifications — Office LLM Harness

This directory contains speckit feature specifications for the Office LLM Harness project, a Windows desktop Office add-in harness that exposes controlled document interaction tools to Open WebUI through a local MCP server.

## Architecture

```
Open WebUI                    MCP Server (port 3000)              Office Add-ins
    │                                │                               │
    │── tools/call ─────────────────►│                               │
    │   { tool, { instanceId?, ... } │                               │
    │                                │── push command (SignalR) ────►│
    │◄── result ─────────────────────│◄── push result (SignalR) ──────│
    │                                │◄── /instances/register ──────│  (on load)
    │                                │◄── SignalR connect ──────────│  (on load)
    │                                │── /instances/:id/heartbeat ─│  (every 10s, HTTP fallback)
    │                                │── /instances/:id/commands ──│  (HTTP fallback only)
    │                                │◄── /instances/:id/result ────│  (HTTP fallback only)
```

**Key design decisions:**

- **Central hub**: Single MCP server process, one port (3000)
- **Unified add-in**: One Office JS Add-in, one manifest — auto-detects host via `Office.onReady(info.host)` and presents host-relevant tools
- **SignalR transport**: Commands pushed in real-time via WebSocket. HTTP polling available as fallback
- **Host-aware tool routing**: Instance registration includes host type (`PowerPoint`, `Word`, `Excel`, `Outlook`); MCP server filters available tools by host
- **Multiple instances**: Each Office session gets its own registered instance (e.g., `powerpoint_1`, `word_2`)
- **Transport**: SignalR (WebSocket) + Streamable HTTP fallback + stdio (MCPo compatible) + OpenAPI REST bridge

## Phases

| #   | Spec                                          | Branch                  | Scope                                                                                                                        |
| --- | --------------------------------------------- | ----------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| 0   | [Spike](001-spike/)                           | `001-spike`             | MCP server hub + unified Office JS Add-in (auto-detects host) + Open WebUI integration validation                            |
| 1   | [PowerPoint MVP](002-powerpoint-mvp/)         | `002-powerpoint-mvp`    | Deck outline, slide read, shape update, speaker notes, audit log, task pane                                                  |
| 2   | [Word MVP](003-word-mvp/)                     | `003-word-mvp`          | Outline, paragraphs, rewrite selection (tracked changes), review comments, shared context abstraction                        |
| 3   | [Excel MVP](004-excel-mvp/)                   | `004-excel-mvp`         | Workbook map, read/write range, write formula, create table, range limits, formula validation                                |
| 4   | [Outlook MVP](005-outlook-mvp/)               | `005-outlook-mvp`       | Current item read, thread summary, draft reply (never auto-send), category apply, policy filter, send confirmation gate      |
| 5   | [PowerPoint v2](006-powerpoint-v2/)           | `006-powerpoint-v2`     | Full shape properties, image export, table reading, selection context, direct write operations, shape CRUD, slide management |
| 8   | [SignalR Transport](008-signalr-transport/)   | `008-signalr-transport` | WebSocket transport upgrade, fix instance ID naming, concurrent commands, HTTP fallback                                      |
| 9   | [Outlook Graph](009-outlook-graph/)           | `009-outlook-graph`     | Folder listing, email search, calendar events, shared mailboxes, compose email, Graph API via NAA proxy                      |
| 10  | [Excel Analysis](010-excel-analysis/)         | `010-excel-analysis`    | Sheet management, sort, filter, charts, conditional formatting, cell formatting, pivot tables                                |
| 11  | [Word Structure](011-word-structure/)         | `011-word-structure`    | Tables CRUD, headers/footers, replace selection, images, styles, sections, lists                                             |
| 12  | [Cross-Cutting](012-cross-cutting/)           | `012-cross-cutting`     | Unified document context, consistent error codes, document stats, batch operations                                           |
| 13  | [Document Export](013-document-export/)       | `013-document-export`   | PDF/native export via getFileAsync, visual verification, document backup                                                     |
| 14  | [Word Find & Replace](014-word-find-replace/) | `014-word-find-replace` | Regex/wildcard find & replace, scoped replacement, preview mode, tracked changes                                             |
| 15  | [Change Log](015-change-log/)                 | `015-change-log`        | Operation audit trail, undo operations, change summaries                                                                     |
| 16  | [Excel Navigation](016-excel-navigation/)     | `016-excel-navigation`  | Freeze panes, named ranges (get/add), workbook navigation aids                                                               |
| 17  | [Excel Validation](017-excel-validation/)     | `017-excel-validation`  | Data validation (dropdowns, number/date constraints, custom formulas), input messages, error alerts                          |
| 18  | [Cross-Host Expansion](018-cross-host-expansion/) | `018-cross-host` | 30 tools: PPT tags/shape formatting/geometric shapes/merge, Word bookmarks/properties/hyperlinks/footnotes/fields, Excel protection/page layout |
| 19  | [Dynamic Tool Filtering](019-dynamic-tool-filtering/) | `019-dynamic-filter` | Filter tools/list by active hosts, reduce context from 107 to only connected host tools                             |

## Progression

Each phase is **independently completable** and builds on the shared infrastructure established by earlier phases:

```
Phase 0 (Spike) → Phase 1 (PPT) → Phase 2 (Word) → Phase 3 (Excel) → Phase 4 (Outlook)
     │                │               │                │                 │
     └── MCP hub     ├── PPT tools   ├── Word tools   ├── Excel tools   ├── Outlook tools
         + unified      + audit log     + shared ctx     + range limits    + policy filter
           add-in        + task pane
           (auto-detect
            host)
```

## Tool Design

Tools are organized by host. The MCP server registers a **host-aware tool list** — only tools relevant to the connected instance's host type are exposed. All tools accept an optional `instanceId` parameter to target specific registered instances.

### Host Detection

When the add-in loads, `Office.onReady((info) => ...)` provides `info.host`:

- `HostType.PowerPoint` → registers as `powerpoint_N`, exposes PowerPoint tools
- `HostType.Word` → registers as `word_N`, exposes Word tools
- `HostType.Excel` → registers as `excel_N`, exposes Excel tools
- `HostType.Outlook` → registers as `outlook_N`, exposes Outlook tools

The instance registration (`POST /instances/register`) includes the `appName` field, and the MCP server uses this to filter which tools are available for each instance.

### Unified vs Host-Specific Tools

| Tool                              | Host       | Parameters                                                    |
| --------------------------------- | ---------- | ------------------------------------------------------------- |
| `office_get_active_app`           | **All**    | _(none)_ — lists all instances                                |
| `powerpoint_get_deck_outline`     | PowerPoint | `instanceId?`, `includeSpeakerNotes?`, `includeHiddenSlides?` |
| `powerpoint_get_slide`            | PowerPoint | `instanceId?`, `slideIndex` (required)                        |
| `powerpoint_update_shape_text`    | PowerPoint | `instanceId?`, `slideIndex`, `shapeId`, `text`                |
| `powerpoint_update_speaker_notes` | PowerPoint | `instanceId?`, `slideIndex`, `notes`                          |
| `word_get_outline`                | Word       | `instanceId?`, `maxDepth?`                                    |
| `word_rewrite_selection`          | Word       | `instanceId?`, `tone`                                         |
| `excel_get_workbook_map`          | Excel      | `instanceId?`                                                 |
| `excel_read_range`                | Excel      | `instanceId?`, `sheetName`, `address`                         |
| `excel_write_range`               | Excel      | `instanceId?`, `range`, `values`                              |
| `outlook_get_current_item`        | Outlook    | `instanceId?`                                                 |
| `outlook_summarize_thread`        | Outlook    | `instanceId?`                                                 |
| `outlook_draft_reply`             | Outlook    | `instanceId?`, `tone`, `keyPoints`                            |

`slideIndex` and other slide-index-shaped parameters (e.g. `atIndex`, `fromIndex`/`toIndex`, `targetIndex`) are 1-based, matching the slide numbers shown in the PowerPoint UI.

### Planned (Phases 8–12)

| Tool                           | Host    | Phase | Scope                                    |
| ------------------------------ | ------- | ----- | ---------------------------------------- |
| `outlook_list_folders`         | Outlook | 9     | List mail folders with unread counts     |
| `outlook_list_emails`          | Outlook | 9     | Paginated email listing in folder        |
| `outlook_search_emails`        | Outlook | 9     | Full-text search across mailbox          |
| `outlook_get_email`            | Outlook | 9     | Read specific email by ID                |
| `outlook_list_calendar_events` | Outlook | 9     | Upcoming calendar appointments           |
| `outlook_compose_email`        | Outlook | 9     | Create new email draft (never auto-send) |
| `outlook_move_email`           | Outlook | 9     | Move email to folder                     |
| `excel_add_sheet`              | Excel   | 10    | Add/rename/delete worksheets             |
| `excel_sort_range`             | Excel   | 10    | Multi-column sort                        |
| `excel_filter_range`           | Excel   | 10    | Autofilter with criteria                 |
| `excel_create_chart`           | Excel   | 10    | Create chart from data range             |
| `excel_format_range`           | Excel   | 10    | Font, fill, borders, alignment           |
| `excel_create_pivottable`      | Excel   | 10    | Pivot table from data range              |
| `word_get_tables`              | Word    | 11    | Read tables with cell content            |
| `word_insert_table`            | Word    | 11    | Create table at location                 |
| `word_replace_selection`       | Word    | 11    | Replace current selection (tracked)      |
| `word_get_headers_footers`     | Word    | 11    | Read header/footer content               |
| `word_insert_image`            | Word    | 11    | Insert inline image                      |
| `office_get_document_context`  | **All** | 12    | Unified context for any host             |

### Planned (Phases 13–17)

| Tool                           | Host    | Phase | Scope                                              |
| ------------------------------ | ------- | ----- | -------------------------------------------------- |
| `office_export_document`       | **All** | 13    | Export as PDF or native (PPTX/DOCX/XLSX)           |
| `word_find_replace`            | Word    | 14    | Regex/wildcard find & replace with tracked changes |
| `word_find_all`                | Word    | 14    | Preview matches without replacing                  |
| `office_get_change_log`        | **All** | 15    | Query audit trail for all operations               |
| `office_undo_last`             | **All** | 15    | Undo last operation via Ctrl+Z                     |
| `excel_freeze_panes`           | Excel   | 16    | Freeze/unfreeze rows and columns                   |
| `excel_get_named_ranges`       | Excel   | 16    | List all named ranges                              |
| `excel_add_named_range`        | Excel   | 16    | Create or update a named range                     |
| `excel_add_data_validation`    | Excel   | 17    | Dropdown, number, date, custom validation          |
| `excel_remove_data_validation` | Excel   | 17    | Clear validation from a range                      |

## Workflow

For each phase, follow the speckit workflow:

1. `/speckit.specify` — Refine the specification (already done)
2. `/speckit.plan` — Create implementation plan
3. `/speckit.tasks` — Generate actionable tasks
4. `/speckit.implement` — Execute implementation

## Constitution

All phases must comply with the project constitution at [`.specify/memory/constitution.md`](../.specify/memory/constitution.md).

Core principles:

- **Safety-First**: No arbitrary COM/shell/macro execution as model-callable tools
- **Local-Only**: MCP binds to `127.0.0.1` only
- **User Control**: Model proposes, user approves — all mutations require confirmation
- **Minimal Surface Area**: Narrow, intentional tools with typed inputs
- **Phased Delivery**: Each host is an independent MVP phase
