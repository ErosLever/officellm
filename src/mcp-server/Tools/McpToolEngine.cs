using System.Text.Json;
using Microsoft.AspNetCore.SignalR;
using OfficeMcpServer.Models;
using OfficeMcpServer.Hubs;

namespace OfficeMcpServer.Tools;

/// <summary>
/// Handles MCP tool definition generation and command dispatch.
/// Uses instance-aware routing with a fixed tool list.
/// </summary>
public static class McpToolEngine
{
    private static InstanceRegistry _registry = new();
    private static CommandStore _commandStore = new();
    private static AuditLog _auditLog = new();
    private static IHubContext<CommandHub>? _hubContext;

    /// <summary>
    /// Sets the SignalR hub context for real-time command push.
    /// Called during app startup after the hub is mapped.
    /// </summary>
    public static void SetHubContext(IHubContext<CommandHub> hubContext)
    {
        _hubContext = hubContext;
    }
    /// <summary>
    /// Gets all MCP tool definitions (fixed list, instance selected via parameter).
    /// </summary>
    public static object[] GetToolDefinitions() => [
        // ── Shared ──────────────────────────────────────────────
        new
        {
            name = "office_get_active_apps",
            description = "Returns a list of all registered Office instances (PowerPoint, Word, Excel, Outlook) with their document names. Call this FIRST so the user can choose which document to work with. Then pass the chosen instanceId to all other tools.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>(),
                required = Array.Empty<string>()
            }
        },

        // ── Cross-cutting tools ───────────────────────────────────
        new
        {
            name = "office_get_document_context",
            description = "Returns a rich summary of the current Office environment: which app is active, document metadata, selection state, and host-specific details. One call gives you everything you need to understand the user's environment without calling host-specific tools first.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "The instance ID to get context for. If omitted, returns context for all active instances." }
                },
                required = Array.Empty<string>()
            }
        },
        new
        {
            name = "office_get_document_stats",
            description = "Returns quantifiable metrics about the current document: Word (word/page/paragraph count), Excel (sheet/cell count), PowerPoint (slide/shape count), Outlook (folder item/unread count). Use for 'How long is this document?' type questions.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "office_batch_call",
            description = "Execute multiple tool calls in parallel in a single request. Reduces latency from O(N roundTrips) to O(1 roundTrip). Returns per-operation results preserving input order. Max 10 operations per batch.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["calls"] = new
                    {
                        type = "array",
                        description = "Array of tool invocations. Each entry: { toolName, args }. Max 10.",
                        items = new
                        {
                            type = "object",
                            properties = new Dictionary<string, object>
                            {
                                ["toolName"] = new { type = "string", description = "Tool to call" },
                                ["args"] = new { type = "object", description = "Arguments for the tool" }
                            },
                            required = new[] { "toolName" }
                        }
                    }
                },
                required = new[] { "calls" }
            }
        },
        new
        {
            name = "office_suggest_tools",
            description = "Returns relevant tool suggestions for the current host with descriptions and example invocations. Use when the user asks 'What can you do?' or 'What tools are available for this document?'",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "The instance ID to suggest tools for." },
                    ["category"] = new { type = "string", description = "Filter by category: 'Read', 'Write', or omit for all." }
                },
                required = Array.Empty<string>()
            }
        },

        // ── Read tools ──────────────────────────────────────────
        new
        {
            name = "powerpoint_get_deck_outline",
            description = "Returns the full slide deck outline. Each slide lists its shapes with type, position, size, and text content. Use office_get_active_apps first to find the right instanceId.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps (e.g. 'powerpoint_1')." },
                    ["includeSpeakerNotes"] = new { type = "boolean", description = "Include speaker notes in the outline", @default = false },
                    ["includeHiddenSlides"] = new { type = "boolean", description = "Include hidden slides in the outline", @default = false }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "powerpoint_get_slide",
            description = "Returns all shapes on a single slide with full properties: type, position (left/top), size (width/height), rotation, text content, font styling (name/size/bold/italic/color), paragraph alignment, and fill color. Use powerpoint_get_deck_outline first to see slide indices.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" }
                },
                required = new[] { "instanceId", "slideIndex" }
            }
        },
        new
        {
            name = "powerpoint_get_slide_image",
            description = "Renders a slide as a PNG image and returns it as base64. The LLM can use this to visually understand the slide layout and content. Optionally specify width/height to control image size (preserves aspect ratio).",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["width"] = new { type = "integer", description = "Max image width in pixels. Default: 800", @default = 800 },
                    ["height"] = new { type = "integer", description = "Max image height in pixels. Default: omitted (auto)" }
                },
                required = new[] { "instanceId", "slideIndex" }
            }
        },
        new
        {
            name = "powerpoint_get_shape_image",
            description = "Renders a single shape as a PNG image and returns it as base64. Useful for describing images, charts, icons, or diagrams on a slide.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["shapeId"] = new { type = "string", description = "Shape ID or name to render" },
                    ["width"] = new { type = "integer", description = "Max image width in pixels. Default: 400", @default = 400 },
                    ["height"] = new { type = "integer", description = "Max image height in pixels. Default: omitted (auto)" }
                },
                required = new[] { "instanceId", "slideIndex", "shapeId" }
            }
        },
        new
        {
            name = "powerpoint_get_table",
            description = "Reads all cell text from a table shape on a slide. Returns rowCount, columnCount, and a 2D string array of cell contents.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["shapeId"] = new { type = "string", description = "Shape ID of the table" }
                },
                required = new[] { "instanceId", "slideIndex", "shapeId" }
            }
        },
        new
        {
            name = "powerpoint_get_selection",
            description = "Returns what the user currently has selected in PowerPoint: selected text (with formatting and parent shape), selected shapes (with IDs and properties), or empty selection. Use this to understand user intent context.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "powerpoint_get_speaker_notes",
            description = "Returns speaker notes for one slide or a range of slides. Specify slideIndex for a single slide, or slideRange (e.g. '2-5') for multiple.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index for a single slide" },
                    ["slideRange"] = new { type = "string", description = "Range of slides, e.g. '2-5'" }
                },
                required = new[] { "instanceId" }
            }
        },

        // ── Write tools (direct, no confirmation gate) ──────────
        new
        {
            name = "powerpoint_update_shape_text",
            description = "Updates the text content of a specific shape on a slide. Applies directly — no confirmation required. Users should create backups before enabling write access.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["shapeId"] = new { type = "string", description = "Shape ID or name" },
                    ["text"] = new { type = "string", description = "New text content" }
                },
                required = new[] { "instanceId", "slideIndex", "shapeId", "text" }
            }
        },
        new
        {
            name = "powerpoint_update_shape_properties",
            description = "Updates position, size, rotation, and/or font properties of a shape. Only specified properties are changed. Use this to resize, reposition, or restyle shapes.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["shapeId"] = new { type = "string", description = "Shape ID or name" },
                    ["left"] = new { type = "number", description = "X position in points" },
                    ["top"] = new { type = "number", description = "Y position in points" },
                    ["width"] = new { type = "number", description = "Width in points" },
                    ["height"] = new { type = "number", description = "Height in points" },
                    ["rotation"] = new { type = "number", description = "Rotation in degrees" },
                    ["fontName"] = new { type = "string", description = "Font family name (e.g. 'Arial')" },
                    ["fontSize"] = new { type = "number", description = "Font size in points" },
                    ["bold"] = new { type = "boolean", description = "Bold on/off" },
                    ["italic"] = new { type = "boolean", description = "Italic on/off" },
                    ["color"] = new { type = "string", description = "Font color as HTML hex (e.g. '#FF0000')" }
                },
                required = new[] { "instanceId", "slideIndex", "shapeId" }
            }
        },
        new
        {
            name = "powerpoint_update_speaker_notes",
            description = "Sets speaker notes for a specific slide. Replaces any existing notes. Applies directly.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["notes"] = new { type = "string", description = "Speaker notes text" }
                },
                required = new[] { "instanceId", "slideIndex", "notes" }
            }
        },

        // ── Shape CRUD ──────────────────────────────────────────
        new
        {
            name = "powerpoint_add_textbox",
            description = "Creates a new text box on a slide with the specified text, position, and size. Returns the new shape's ID.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["text"] = new { type = "string", description = "Initial text content" },
                    ["left"] = new { type = "number", description = "X position in points" },
                    ["top"] = new { type = "number", description = "Y position in points" },
                    ["width"] = new { type = "number", description = "Width in points" },
                    ["height"] = new { type = "number", description = "Height in points" }
                },
                required = new[] { "instanceId", "slideIndex", "text", "left", "top", "width", "height" }
            }
        },
        new
        {
            name = "powerpoint_add_image",
            description = "Inserts an image onto a slide from base64-encoded data. Returns the new shape's ID.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["imageBase64"] = new { type = "string", description = "Base64-encoded image data (PNG, JPG, etc.)" },
                    ["left"] = new { type = "number", description = "X position in points" },
                    ["top"] = new { type = "number", description = "Y position in points" },
                    ["width"] = new { type = "number", description = "Width in points" },
                    ["height"] = new { type = "number", description = "Height in points" }
                },
                required = new[] { "instanceId", "slideIndex", "imageBase64", "left", "top" }
            }
        },
        new
        {
            name = "powerpoint_add_table",
            description = "Creates a new table on a slide with specified rows and columns. Returns the new shape's ID.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["rows"] = new { type = "integer", description = "Number of rows" },
                    ["columns"] = new { type = "integer", description = "Number of columns" },
                    ["left"] = new { type = "number", description = "X position in points" },
                    ["top"] = new { type = "number", description = "Y position in points" },
                    ["width"] = new { type = "number", description = "Width in points" },
                    ["height"] = new { type = "number", description = "Height in points" }
                },
                required = new[] { "instanceId", "slideIndex", "rows", "columns", "left", "top" }
            }
        },
        new
        {
            name = "powerpoint_delete_shape",
            description = "Deletes a shape from a slide. Irreversible — users should create backups.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index" },
                    ["shapeId"] = new { type = "string", description = "Shape ID or name to delete" }
                },
                required = new[] { "instanceId", "slideIndex", "shapeId" }
            }
        },

        // ── Slide management ────────────────────────────────────
        new
        {
            name = "powerpoint_add_slide",
            description = "Inserts a new blank slide at the specified position. If atIndex is omitted, adds at the end.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["atIndex"] = new { type = "integer", description = "Zero-based index to insert at. Default: end of deck." }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "powerpoint_delete_slide",
            description = "Deletes a slide from the presentation. Irreversible — users should create backups.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["slideIndex"] = new { type = "integer", description = "Zero-based slide index to delete" }
                },
                required = new[] { "instanceId", "slideIndex" }
            }
        },
        new
        {
            name = "powerpoint_move_slide",
            description = "Moves a slide from one position to another. Other slides shift accordingly.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["fromIndex"] = new { type = "integer", description = "Current zero-based position of the slide" },
                    ["toIndex"] = new { type = "integer", description = "Target zero-based position" }
                },
                required = new[] { "instanceId", "fromIndex", "toIndex" }
            }
        },

        // ── Word Read tools ───────────────────────────────────────
        new
        {
            name = "word_get_outline",
            description = "Returns the document outline (headings with levels, styles, and text). Use office_get_active_apps first to find the right instanceId.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["maxDepth"] = new { type = "integer", description = "Maximum heading level to include (1-9). Default: 3", @default = 3 }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "word_get_paragraphs",
            description = "Returns paragraphs from the document body, optionally filtered by range (startIndex, count).",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["startIndex"] = new { type = "integer", description = "Zero-based paragraph index to start from. Default: 0", @default = 0 },
                    ["count"] = new { type = "integer", description = "Maximum paragraphs to return. Default: 50", @default = 50 }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "word_get_selection",
            description = "Returns the currently selected text in Word, with paragraph context and formatting.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "word_search",
            description = "Searches the document for text and returns matching paragraphs with their locations.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["searchText"] = new { type = "string", description = "Text to search for" },
                    ["matchCase"] = new { type = "boolean", description = "Case-sensitive search. Default: false", @default = false }
                },
                required = new[] { "instanceId", "searchText" }
            }
        },

        // ── Word Write tools ──────────────────────────────────────
        new
        {
            name = "word_replace_text",
            description = "Replaces text in a specific paragraph by index. Applies directly — users should create backups.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["paragraphIndex"] = new { type = "integer", description = "Zero-based paragraph index" },
                    ["oldText"] = new { type = "string", description = "Text to find and replace within the paragraph" },
                    ["newText"] = new { type = "string", description = "Replacement text" }
                },
                required = new[] { "instanceId", "paragraphIndex", "oldText", "newText" }
            }
        },
        new
        {
            name = "word_insert_text",
            description = "Inserts text at a specific location (before/after a paragraph, or at the end of the document).",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["text"] = new { type = "string", description = "Text to insert" },
                    ["insertLocation"] = new { type = "string", description = "Where to insert: 'end' (default), 'afterParagraph', 'beforeParagraph'", @default = "end" },
                    ["paragraphIndex"] = new { type = "integer", description = "Zero-based paragraph index (required for afterParagraph/beforeParagraph)" }
                },
                required = new[] { "instanceId", "text" }
            }
        },
        new
        {
            name = "word_add_comment",
            description = "Adds a Word comment. Three anchor modes, checked in order: (1) searchText — anchors to the first case-insensitive match of a phrase anywhere in the document; (2) fromParagraph / toParagraph — anchors to a range of whole paragraphs (toParagraph defaults to fromParagraph for a single paragraph; paragraphIndex is a shorthand for both); (3) no anchor — attaches to the current user selection. Prefer searchText for sub-paragraph precision.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["commentText"] = new { type = "string", description = "REQUIRED. The comment text." },
                    ["searchText"] = new { type = "string", description = "Exact phrase to anchor the comment to (case-insensitive, first match). Takes priority over paragraph params." },
                    ["paragraphIndex"] = new { type = "integer", description = "Zero-based paragraph index. Shorthand for fromParagraph=toParagraph=this value." },
                    ["fromParagraph"] = new { type = "integer", description = "Zero-based index of the first paragraph in a multi-paragraph span." },
                    ["toParagraph"] = new { type = "integer", description = "Zero-based index of the last paragraph in the span (inclusive). Defaults to fromParagraph." }
                },
                required = new[] { "instanceId", "commentText" }
            }
        },
        new
        {
            name = "word_delete_paragraph",
            description = "Deletes a paragraph by index. Creates a tracked deletion if change tracking is enabled.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["paragraphIndex"] = new { type = "integer", description = "Zero-based paragraph index to delete" }
                },
                required = new[] { "instanceId", "paragraphIndex" }
            }
        },
        new
        {
            name = "word_get_tracked_changes",
            description = "Returns the current change tracking mode and count of pending revisions in the Word document.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "word_accept_all_changes",
            description = "Accepts all tracked changes in the Word document. Finalizes all pending insertions, deletions, and formatting changes.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "word_reject_all_changes",
            description = "Rejects all tracked changes in the Word document. Reverts all pending insertions, deletions, and formatting changes.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." }
                },
                required = new[] { "instanceId" }
            }
        },

        // ── Word Structure: Tables ──────────────────────────────
        new { name = "word_get_tables", description = "Returns all tables in the document with row/column counts and cell text for each table.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["includeCellText"] = new { type = "boolean", description = "Include cell text content. Default: true.", @default = true }, ["maxRows"] = new { type = "number", description = "Max rows to read per table. Default: 50.", @default = 50 } }, required = new[] { "instanceId" } } },
        new { name = "word_insert_table", description = "Inserts a table at a specified location. Supports header row population. Uses tracked changes.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["rows"] = new { type = "number", description = "Number of rows (excluding header)." }, ["columns"] = new { type = "number", description = "Number of columns." }, ["afterParagraphIndex"] = new { type = "number", description = "Insert after this 0-based paragraph index. Use -1 for end of document.", @default = -1 }, ["headerRow"] = new { type = "array", description = "Optional header row values.", items = new { type = "string" } } }, required = new[] { "instanceId", "rows", "columns" } } },
        new { name = "word_update_table_cell", description = "Updates a single cell in a table. Uses tracked changes. Coordinates are 0-based.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["tableIndex"] = new { type = "number", description = "0-based table index." }, ["row"] = new { type = "number", description = "0-based row index." }, ["column"] = new { type = "number", description = "0-based column index." }, ["text"] = new { type = "string", description = "New cell text." } }, required = new[] { "instanceId", "tableIndex", "row", "column", "text" } } },

        // ── Word Structure: Headers/Footers ────────────────────
        new { name = "word_get_headers_footers", description = "Returns header and footer content for each section, including default, first page, and odd/even variants.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sectionIndex"] = new { type = "number", description = "0-based section index. Default: all sections." } }, required = new[] { "instanceId" } } },
        new { name = "word_set_header_footer", description = "Sets header or footer text for a section. Uses tracked changes.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sectionIndex"] = new { type = "number", description = "0-based section index. Default: 0.", @default = 0 }, ["type"] = new { type = "string", description = "'header' or 'footer'." }, ["variant"] = new { type = "string", description = "'default', 'firstPage', or 'oddEven'. Default: 'default'.", @default = "default" }, ["text"] = new { type = "string", description = "Content text to set." } }, required = new[] { "instanceId", "type", "text" } } },

        // ── Word Structure: Selection & Insert ──────────────────
        new { name = "word_replace_selection", description = "Replaces the current selection with new text. Uses tracked changes — user accepts/rejects via Review ribbon.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["text"] = new { type = "string", description = "Replacement text." } }, required = new[] { "instanceId", "text" } } },
        new { name = "word_insert_image", description = "Inserts an inline image at a specified location. Accepts base64-encoded image data (PNG, JPEG, GIF, BMP, SVG). Max 10MB.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["imageBase64"] = new { type = "string", description = "Base64-encoded image data." }, ["afterParagraphIndex"] = new { type = "number", description = "Insert after this paragraph. -1 = end. Default: -1.", @default = -1 }, ["width"] = new { type = "number", description = "Optional width in points." }, ["height"] = new { type = "number", description = "Optional height in points." } }, required = new[] { "instanceId", "imageBase64" } } },

        // ── Word Structure: Styles & Lists ──────────────────────
        new { name = "word_apply_style", description = "Applies a named paragraph style or character style to a paragraph or range. For paragraph styles use styleName + paragraphIndex. For character styles (applies to a text run/selection) use characterStyle + searchText or applyToSelection.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["paragraphIndex"] = new { type = "number", description = "0-based paragraph index. Used for paragraph styles." }, ["styleName"] = new { type = "string", description = "Paragraph style name (e.g. 'Heading 1', 'Normal', 'Title')." }, ["characterStyle"] = new { type = "string", description = "Character style name (e.g. 'Strong', 'Emphasis', 'Code Char'). Applied to searchText match or current selection." }, ["searchText"] = new { type = "string", description = "Phrase to apply the character style to (first match). Used with characterStyle." }, ["applyToSelection"] = new { type = "boolean", description = "Apply characterStyle to the current user selection. Used with characterStyle." } }, required = new[] { "instanceId" } } },
        new { name = "word_get_sections", description = "Returns document sections with page layout info and header/footer configuration.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." } }, required = new[] { "instanceId" } } },
        new { name = "word_insert_list", description = "Inserts a bulleted or numbered list at a specified location. Uses tracked changes.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["type"] = new { type = "string", description = "'bulleted' or 'numbered'. Default: 'bulleted'.", @default = "bulleted" }, ["items"] = new { type = "array", description = "List item texts.", items = new { type = "string" } }, ["afterParagraphIndex"] = new { type = "number", description = "Insert after this paragraph. -1 = end. Default: -1.", @default = -1 } }, required = new[] { "instanceId", "items" } } },

        // ── Excel Read tools ──────────────────────────────────────
        new
        {
            name = "excel_get_workbook_map",
            description = "Returns the workbook structure: sheet names, used range dimensions, tables (name, range, sheet), and named ranges. Use office_get_active_apps first to find the right instanceId.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps (e.g. 'excel_1')." },
                    ["includeTables"] = new { type = "boolean", description = "Include table definitions in the response. Default: true", @default = true },
                    ["includeNamedRanges"] = new { type = "boolean", description = "Include named ranges in the response. Default: true", @default = true }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "excel_read_range",
            description = "Reads cell values, formulas, and optional number formats from a bounded range. Specify sheet name and A1-style address (e.g. 'B2:G24'). Returns a 2D grid with row/column indices.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["sheetName"] = new { type = "string", description = "Name of the worksheet to read from" },
                    ["address"] = new { type = "string", description = "A1-style range address (e.g. 'A1:D10') or single cell (e.g. 'B2')" },
                    ["includeFormulas"] = new { type = "boolean", description = "Include formula strings for formula cells. Default: true", @default = true },
                    ["includeNumberFormats"] = new { type = "boolean", description = "Include number format strings. Default: false", @default = false }
                },
                required = new[] { "instanceId", "sheetName", "address" }
            }
        },

        // ── Excel Write tools ─────────────────────────────────────
        new
        {
            name = "excel_write_range",
            description = "Writes values to a bounded range on a worksheet. Applies directly — users should create backups. Values is a 2D array (rows of columns). The range dimensions must match the data dimensions.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["sheetName"] = new { type = "string", description = "Name of the worksheet to write to" },
                    ["address"] = new { type = "string", description = "A1-style range address to write to (e.g. 'A1:C3')" },
                    ["values"] = new { type = "array", description = "2D array of values (rows of columns). Numbers, strings, or null for empty.", items = new { type = "array", items = new { type = "string" } } }
                },
                required = new[] { "instanceId", "sheetName", "address", "values" }
            }
        },
        new
        {
            name = "excel_write_formula",
            description = "Writes a formula to a cell or range on a worksheet. Validates formula syntax (must start with '='). Applies directly — users should create backups.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["sheetName"] = new { type = "string", description = "Name of the worksheet" },
                    ["address"] = new { type = "string", description = "A1-style cell or range address for the formula (e.g. 'D2' or 'D2:D24')" },
                    ["formula"] = new { type = "string", description = "Excel formula string (must start with '='). e.g. '=SUM(B2:B24)'" }
                },
                required = new[] { "instanceId", "sheetName", "address", "formula" }
            }
        },
        new
        {
            name = "excel_create_table",
            description = "Creates a formatted Excel table (ListObject) from a specified range. Returns the table name and range. Applies directly.", 
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["sheetName"] = new { type = "string", description = "Name of the worksheet containing the data" },
                    ["address"] = new { type = "string", description = "A1-style range address including headers (e.g. 'A1:D25')" },
                    ["tableName"] = new { type = "string", description = "Optional name for the table. Default: auto-generated by Excel." },
                    ["hasHeaders"] = new { type = "boolean", description = "First row contains headers. Default: true", @default = true }
                },
                required = new[] { "instanceId", "sheetName", "address" }
            }
        },

        // ── Excel Sheet Management ──────────────────────────────
        new { name = "excel_add_sheet", description = "Adds a new worksheet to the workbook. Returns the new sheet name and position. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." }, ["name"] = new { type = "string", description = "Name for the new sheet. Default: auto-generated." }, ["position"] = new { type = "number", description = "0-based tab position. Default: end of tab list." } }, required = new[] { "instanceId" } } },
        new { name = "excel_delete_sheet", description = "Deletes a worksheet by name. Cannot delete the last sheet. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." }, ["sheetName"] = new { type = "string", description = "Name of the worksheet to delete." } }, required = new[] { "instanceId", "sheetName" } } },
        new { name = "excel_rename_sheet", description = "Renames an existing worksheet. Returns old and new names. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." }, ["sheetName"] = new { type = "string", description = "Current name of the worksheet." }, ["newName"] = new { type = "string", description = "New name for the worksheet." } }, required = new[] { "instanceId", "sheetName", "newName" } } },

        // ── Excel Sort & Filter ──────────────────────────────────
        new { name = "excel_sort_range", description = "Sorts a data range by one or more columns. Supports multi-column sort with ascending/descending order. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sheetName"] = new { type = "string", description = "Worksheet name." }, ["address"] = new { type = "string", description = "A1-style range address." }, ["criteria"] = new { type = "array", description = "Sort criteria: [{ column: 0-based index, ascending: bool }].", items = new { type = "object" } }, ["hasHeader"] = new { type = "boolean", description = "First row is header. Default: true.", @default = true } }, required = new[] { "instanceId", "sheetName", "address", "criteria" } } },
        new { name = "excel_filter_range", description = "Applies autofilter to a data range with value-based criteria. Use clearFilters to remove all filters. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sheetName"] = new { type = "string", description = "Worksheet name." }, ["address"] = new { type = "string", description = "A1-style range address including headers." }, ["column"] = new { type = "number", description = "0-based column index to filter on." }, ["criteria"] = new { type = "object", description = "Filter criteria: { value: 'Active' } or { operator: 'greaterThan', value: 100 }." }, ["clearFilters"] = new { type = "boolean", description = "Remove all filters. Default: false.", @default = false } }, required = new[] { "instanceId", "sheetName", "address" } } },

        // ── Excel Charts ─────────────────────────────────────────
        new { name = "excel_create_chart", description = "Creates a chart from a data range. Types: Column, Bar, Line, Pie, Scatter, Area, Doughnut. Auto-positioned adjacent to data. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sheetName"] = new { type = "string", description = "Worksheet containing data." }, ["dataRange"] = new { type = "string", description = "A1-style data range (e.g. 'A1:B10')." }, ["chartType"] = new { type = "string", description = "Chart type. Default: Column.", @default = "Column" }, ["title"] = new { type = "string", description = "Optional chart title." } }, required = new[] { "instanceId", "sheetName", "dataRange" } } },
        new { name = "excel_get_charts", description = "Returns a list of all charts with type, title, data range, and position.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sheetName"] = new { type = "string", description = "Filter to specific sheet. Default: all sheets." } }, required = new[] { "instanceId" } } },

        // ── Excel Formatting ─────────────────────────────────────
        new { name = "excel_format_range", description = "Applies formatting to a range: font, fill, borders, alignment, numberFormat. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sheetName"] = new { type = "string", description = "Worksheet name." }, ["address"] = new { type = "string", description = "A1-style range address." }, ["font"] = new { type = "object", description = "Font: { name, size, bold, italic, color, underline }." }, ["fill"] = new { type = "object", description = "Fill: { color }." }, ["borders"] = new { type = "object", description = "Borders: { style, color }." }, ["alignment"] = new { type = "object", description = "Alignment: { horizontal, vertical, wrapText }." }, ["numberFormat"] = new { type = "string", description = "Number format (e.g. '#,##0.00')." } }, required = new[] { "instanceId", "sheetName", "address" } } },
        new { name = "excel_apply_conditional_formatting", description = "Applies conditional formatting. Types: cellValue, dataBar, colorScale, iconSet. Last applied wins on conflict.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sheetName"] = new { type = "string", description = "Worksheet name." }, ["address"] = new { type = "string", description = "A1-style range address." }, ["ruleType"] = new { type = "string", description = "Rule type: cellValue, dataBar, colorScale, iconSet." }, ["operator"] = new { type = "string", description = "For cellValue: greaterThan, lessThan, equalTo, between." }, ["value"] = new { type = "string", description = "Comparison value for cellValue." }, ["format"] = new { type = "object", description = "For cellValue: { fillColor }." }, ["minColor"] = new { type = "string", description = "For colorScale: min color." }, ["maxColor"] = new { type = "string", description = "For colorScale: max color." }, ["iconSet"] = new { type = "string", description = "For iconSet: 3TrafficLights, 3Arrows, etc." } }, required = new[] { "instanceId", "sheetName", "address", "ruleType" } } },

        // ── Excel Pivot Table ────────────────────────────────────
        new { name = "excel_create_pivottable", description = "Creates a pivot table from source data on a new sheet. Specify row/column/value fields with aggregation. Requires headers.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sourceRange"] = new { type = "string", description = "Source range with headers (e.g. 'Data!A1:E1000')." }, ["name"] = new { type = "string", description = "Pivot table name." }, ["rows"] = new { type = "array", description = "Row field names.", items = new { type = "string" } }, ["columns"] = new { type = "array", description = "Column field names.", items = new { type = "string" } }, ["values"] = new { type = "array", description = "Value fields: [{ field, aggregation: sum|count|average|min|max }].", items = new { type = "object" } }, ["destinationSheet"] = new { type = "string", description = "Destination sheet name. Default: 'PivotTable'." } }, required = new[] { "instanceId", "sourceRange", "rows", "values" } } },

        // ── Outlook Read tools ──────────────────────────────────────
        new
        {
            name = "outlook_get_current_item",
            description = "Returns metadata (sender, recipients, subject, received date) and bounded body text for the currently selected Outlook item. Includes attachment metadata (name, size, type) but not file contents.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps (e.g. 'outlook_1')." },
                    ["includeBody"] = new { type = "boolean", description = "Include the email body text. Default: true", @default = true },
                    ["bodyFormat"] = new { type = "string", description = "Body format: 'text' or 'html'. Default: 'text'", @default = "text" }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "outlook_summarize_thread",
            description = "Returns a structured summary of the selected email thread with timeline, key decisions, action items, and unresolved questions. Processes up to 50 messages by default.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["maxMessages"] = new { type = "integer", description = "Maximum messages to process. Default: 50", @default = 50 }
                },
                required = new[] { "instanceId" }
            }
        },

        // ── Outlook Write tools (never auto-send) ──────────────────
        new
        {
            name = "outlook_draft_reply",
            description = "Creates a draft reply in Outlook's Drafts folder for the selected email. The draft is NEVER sent automatically — the user must review and send from Outlook. Supports tone control and key points.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["tone"] = new { type = "string", description = "Reply tone: 'concise', 'formal', 'friendly', or 'technical'. Default: 'concise'", @default = "concise", @enum = new[] { "concise", "formal", "friendly", "technical" } },
                    ["keyPoints"] = new { type = "array", description = "Key points to address in the reply", items = new { type = "string" } },
                    ["includeThreadSummary"] = new { type = "boolean", description = "Include a thread summary in the draft. Default: false", @default = false }
                },
                required = new[] { "instanceId" }
            }
        },
        new
        {
            name = "outlook_apply_category",
            description = "Applies an Outlook category (color-coded label) to the currently selected email(s). Returns available categories if the requested one doesn't exist.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["categoryName"] = new { type = "string", description = "Name of the Outlook category to apply" }
                },
                required = new[] { "instanceId", "categoryName" }
            }
        },
        new
        {
            name = "outlook_send_message",
            description = "Sends a drafted message. REQUIRES explicit user confirmation via the Outlook task pane — a confirmation token must be provided. NEVER auto-sends. Subject to policy filters.",
            inputSchema = new
            {
                type = "object",
                properties = new Dictionary<string, object>
                {
                    ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID from office_get_active_apps." },
                    ["confirmationToken"] = new { type = "string", description = "REQUIRED. Confirmation token obtained from explicit user approval in the Outlook task pane." },
                    ["messageId"] = new { type = "string", description = "Optional. ID of the draft message to send. If omitted, sends the currently open draft." }
                },
                required = new[] { "instanceId", "confirmationToken" }
            }
        },

        // ── Outlook Extended (Office.js only, no Graph) ──────────
        new { name = "outlook_get_user_profile", description = "Returns the current user's profile: display name, email address, and timezone.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." } }, required = new[] { "instanceId" } } },
        new { name = "outlook_get_master_categories", description = "Returns the mailbox's master category list with names and colors.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." } }, required = new[] { "instanceId" } } },
        new { name = "outlook_create_category", description = "Adds a new category to the mailbox's master list.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["name"] = new { type = "string", description = "Category name." }, ["color"] = new { type = "string", description = "Color preset name (e.g. 'preset0'-'preset25'). Default: 'preset0'.", @default = "preset0" } }, required = new[] { "instanceId", "name" } } },
        new { name = "outlook_remove_categories", description = "Removes categories from the current item.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["categories"] = new { type = "array", description = "Category names to remove.", items = new { type = "string" } } }, required = new[] { "instanceId", "categories" } } },
        new { name = "outlook_display_new_message", description = "Opens a new message compose form in Outlook. User must review and send manually. NEVER auto-sends.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["to"] = new { type = "array", description = "Recipient email addresses.", items = new { type = "string" } }, ["cc"] = new { type = "array", description = "CC email addresses.", items = new { type = "string" } }, ["bcc"] = new { type = "array", description = "BCC email addresses.", items = new { type = "string" } }, ["subject"] = new { type = "string", description = "Email subject." }, ["body"] = new { type = "string", description = "Email body text." } }, required = new[] { "instanceId" } } },
        new { name = "outlook_display_new_appointment", description = "Opens a new appointment form in Outlook. User must save manually.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["subject"] = new { type = "string", description = "Appointment subject." }, ["location"] = new { type = "string", description = "Location." }, ["start"] = new { type = "string", description = "Start time (ISO 8601)." }, ["end"] = new { type = "string", description = "End time (ISO 8601)." }, ["requiredAttendees"] = new { type = "array", description = "Required attendee emails.", items = new { type = "string" } }, ["optionalAttendees"] = new { type = "array", description = "Optional attendee emails.", items = new { type = "string" } } }, required = new[] { "instanceId" } } },
        new { name = "outlook_get_attachments", description = "Returns detailed attachment info for the current item: name, size, type, ID, inline status.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." } }, required = new[] { "instanceId" } } },

        // ── Phase 13: Document Export ──────────────────────────────────
        new { name = "office_export_document", description = "Exports the current document as PDF or native format (PPTX/DOCX/XLSX). Returns base64-encoded file data. Use for visual verification, backups, and sharing.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["format"] = new { type = "string", description = "'pdf' or 'native'. Default: 'pdf'.", @default = "pdf" }, ["maxSizeMB"] = new { type = "number", description = "Max file size in MB. Default: 50.", @default = 50 } }, required = new[] { "instanceId" } } },

        // ── Phase 14: Word Find & Replace ──────────────────────────────
        new { name = "word_find_replace", description = "Finds and replaces text in a Word document. Supports wildcards, case matching, whole-word mode, and scoped replacement. Uses tracked changes.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["findText"] = new { type = "string", description = "Text or wildcard pattern to find." }, ["replaceText"] = new { type = "string", description = "Replacement text. Use empty string to delete." }, ["matchCase"] = new { type = "boolean", description = "Case-sensitive match. Default: false.", @default = false }, ["matchWholeWord"] = new { type = "boolean", description = "Match whole words only. Default: false.", @default = false }, ["useWildcards"] = new { type = "boolean", description = "Use Word wildcard syntax (?, *, [a-z]). Default: false.", @default = false }, ["previewOnly"] = new { type = "boolean", description = "List matches without replacing. Default: false.", @default = false }, ["scopeFromParagraph"] = new { type = "number", description = "0-based start paragraph for scoped search." }, ["scopeToParagraph"] = new { type = "number", description = "0-based end paragraph for scoped search." } }, required = new[] { "instanceId", "findText" } } },

        // ── Phase 16: Excel Navigation ─────────────────────────────────
        new { name = "excel_freeze_panes", description = "Freezes rows above and columns left of the specified cell for scroll navigation. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sheetName"] = new { type = "string", description = "Worksheet name." }, ["at"] = new { type = "string", description = "Cell address to freeze above/left of (e.g. 'A2' freezes row 1)." }, ["action"] = new { type = "string", description = "'freeze' or 'unfreeze'. Default: 'freeze'.", @default = "freeze" } }, required = new[] { "instanceId", "sheetName", "at" } } },
        new { name = "excel_get_named_ranges", description = "Returns all named ranges in the workbook with their addresses, scopes, and comments.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." } }, required = new[] { "instanceId" } } },
        new { name = "excel_add_named_range", description = "Creates or updates a named range. Makes formulas readable (e.g. 'Q1Sales' instead of 'B2:D100'). Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["name"] = new { type = "string", description = "Name for the range (e.g. 'Q1Sales')." }, ["sheetName"] = new { type = "string", description = "Worksheet containing the range." }, ["address"] = new { type = "string", description = "A1-style range address (e.g. 'B2:D100')." }, ["comment"] = new { type = "string", description = "Optional comment describing the range." } }, required = new[] { "instanceId", "name", "sheetName", "address" } } },

        // ── Phase 17: Excel Data Validation ────────────────────────────
        new { name = "excel_add_data_validation", description = "Adds data validation to a range: dropdown lists, number/date constraints, text length limits, or custom formulas. Includes input messages and error alerts. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sheetName"] = new { type = "string", description = "Worksheet name." }, ["address"] = new { type = "string", description = "A1-style range address." }, ["type"] = new { type = "string", description = "Validation type: 'list', 'wholeNumber', 'decimal', 'date', 'textLength', 'custom'." }, ["operator"] = new { type = "string", description = "Operator: 'between', 'notBetween', 'equalTo', 'greaterThan', 'lessThan', etc. Default: 'between'.", @default = "between" }, ["formula1"] = new { type = "string", description = "First formula/value (required). For list: 'A,B,C' or '=Sheet!$A$1:$A$10'." }, ["formula2"] = new { type = "string", description = "Second formula for 'between'/'notBetween' operators." }, ["showInputMessage"] = new { type = "boolean", description = "Show input tooltip on cell select. Default: true.", @default = true }, ["inputTitle"] = new { type = "string", description = "Tooltip title." }, ["inputMessage"] = new { type = "string", description = "Tooltip message." }, ["showErrorMessage"] = new { type = "boolean", description = "Show error alert on invalid input. Default: true.", @default = true }, ["errorTitle"] = new { type = "string", description = "Error alert title." }, ["errorMessage"] = new { type = "string", description = "Error alert message." }, ["errorStyle"] = new { type = "string", description = "'stop', 'warning', or 'information'. Default: 'stop'.", @default = "stop" } }, required = new[] { "instanceId", "sheetName", "address", "type" } } },
        new { name = "excel_remove_data_validation", description = "Removes data validation from a range. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sheetName"] = new { type = "string", description = "Worksheet name." }, ["address"] = new { type = "string", description = "A1-style range address." } }, required = new[] { "instanceId", "sheetName", "address" } } },

        // ── Phase 18: PowerPoint Tags & Metadata ──────────────────────
        new { name = "powerpoint_get_tags", description = "Returns all tags (key-value metadata) on a slide, shape, or presentation. Tags enable audience-specific content filtering.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["target"] = new { type = "string", description = "'presentation', 'slide', or 'shape'. Default: 'presentation'.", @default = "presentation" }, ["slideIndex"] = new { type = "number", description = "0-based slide index (required when target='slide' or 'shape')." }, ["shapeId"] = new { type = "string", description = "Shape ID (required when target='shape')." } }, required = new[] { "instanceId" } } },
        new { name = "powerpoint_set_tag", description = "Sets a tag (key-value metadata) on a slide, shape, or presentation. Use uppercase keys for consistency.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["key"] = new { type = "string", description = "Tag key (use UPPERCASE, e.g. 'AUDIENCE')." }, ["value"] = new { type = "string", description = "Tag value." }, ["target"] = new { type = "string", description = "'presentation', 'slide', or 'shape'. Default: 'presentation'.", @default = "presentation" }, ["slideIndex"] = new { type = "number", description = "0-based slide index (required for slide/shape targets)." }, ["shapeId"] = new { type = "string", description = "Shape ID (required when target='shape')." } }, required = new[] { "instanceId", "key", "value" } } },
        new { name = "powerpoint_delete_slides_by_tag", description = "Deletes all slides that have a specific tag key-value pair. Use for audience-specific content filtering (e.g. remove 'Premium' slides).", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["key"] = new { type = "string", description = "Tag key to match." }, ["value"] = new { type = "string", description = "Tag value to match. If omitted, deletes slides that have ANY value for this key." } }, required = new[] { "instanceId", "key" } } },

        // ── Phase 18: PowerPoint Shape Formatting ─────────────────────
        new { name = "powerpoint_set_shape_fill", description = "Sets the fill color, transparency, or image fill on a shape. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["slideIndex"] = new { type = "number", description = "0-based slide index." }, ["shapeId"] = new { type = "string", description = "Shape ID." }, ["fillType"] = new { type = "string", description = "'solid', 'image', or 'none'. Default: 'solid'.", @default = "solid" }, ["color"] = new { type = "string", description = "HTML color (#RRGGBB or named color like 'blue')." }, ["transparency"] = new { type = "number", description = "Transparency 0.0 (opaque) to 1.0 (clear)." }, ["imageBase64"] = new { type = "string", description = "Base64-encoded image for image fill." } }, required = new[] { "instanceId", "slideIndex", "shapeId" } } },
        new { name = "powerpoint_set_shape_line", description = "Sets the line/border format on a shape: color, width, style. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["slideIndex"] = new { type = "number", description = "0-based slide index." }, ["shapeId"] = new { type = "string", description = "Shape ID." }, ["color"] = new { type = "string", description = "Line color (#RRGGBB or named)." }, ["width"] = new { type = "number", description = "Line width in points." }, ["style"] = new { type = "string", description = "'solid', 'dash', 'dashDot', 'dashDotDot', 'longDash', 'roundDot'." }, ["visible"] = new { type = "boolean", description = "Show/hide the line. Default: true." } }, required = new[] { "instanceId", "slideIndex", "shapeId" } } },
        new { name = "powerpoint_set_shape_rotation", description = "Rotates a shape by the specified degrees. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["slideIndex"] = new { type = "number", description = "0-based slide index." }, ["shapeId"] = new { type = "string", description = "Shape ID." }, ["degrees"] = new { type = "number", description = "Rotation in degrees (0-360)." } }, required = new[] { "instanceId", "slideIndex", "shapeId", "degrees" } } },

        // ── Phase 18: PowerPoint Geometric Shapes & Lines ─────────────
        new { name = "powerpoint_add_geometric_shape", description = "Adds a geometric shape (rectangle, circle, arrow, etc.) to a slide. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["slideIndex"] = new { type = "number", description = "0-based slide index." }, ["shapeType"] = new { type = "string", description = "GeometricShapeType: 'rectangle', 'oval', 'triangle', 'rightArrow', 'chevron', 'star5', 'heart', etc." }, ["left"] = new { type = "number", description = "Left position in points." }, ["top"] = new { type = "number", description = "Top position in points." }, ["width"] = new { type = "number", description = "Width in points." }, ["height"] = new { type = "number", description = "Height in points." } }, required = new[] { "instanceId", "slideIndex", "shapeType" } } },
        new { name = "powerpoint_add_line", description = "Adds a line/connector to a slide. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["slideIndex"] = new { type = "number", description = "0-based slide index." }, ["startX"] = new { type = "number", description = "Start X in points." }, ["startY"] = new { type = "number", description = "Start Y in points." }, ["endX"] = new { type = "number", description = "End X in points." }, ["endY"] = new { type = "number", description = "End Y in points." }, ["connectorType"] = new { type = "string", description = "'straight', 'elbow', or 'curve'. Default: 'straight'.", @default = "straight" } }, required = new[] { "instanceId", "slideIndex", "startX", "startY", "endX", "endY" } } },

        // ── Phase 18: PowerPoint Slide Merge ───────────────────────────
        new { name = "powerpoint_insert_slides_from_file", description = "Inserts slides from another PowerPoint file (base64) into the current presentation. Enables template merging.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["base64File"] = new { type = "string", description = "Base64-encoded PPTX file to insert slides from." }, ["insertAfterSlideIndex"] = new { type = "number", description = "0-based slide index to insert after. Default: end of presentation." }, ["slideIndexes"] = new { type = "string", description = "Comma-separated 0-based indexes of slides to insert (e.g. '0,2,5'). Default: all slides." }, ["formatting"] = new { type = "string", description = "'useDestinationTheme' or 'keepSourceFormatting'. Default: 'useDestinationTheme'.", @default = "useDestinationTheme" } }, required = new[] { "instanceId", "base64File" } } },

        // ── Phase 18: PowerPoint Layouts & Theme ────────────────────────
        new { name = "powerpoint_get_layouts", description = "Returns all available slide layouts from the slide master with names and IDs.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." } }, required = new[] { "instanceId" } } },
        new { name = "powerpoint_get_theme_colors", description = "Returns the theme color scheme from the slide master (10 named colors).", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." } }, required = new[] { "instanceId" } } },
        new { name = "powerpoint_group_shapes", description = "Groups multiple shapes into a single group shape. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["slideIndex"] = new { type = "number", description = "0-based slide index." }, ["shapeIds"] = new { type = "string", description = "Comma-separated shape IDs to group." } }, required = new[] { "instanceId", "slideIndex", "shapeIds" } } },
        new { name = "powerpoint_ungroup_shape", description = "Ungroups a grouped shape into its individual shapes. Undoable via Ctrl+Z.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["slideIndex"] = new { type = "number", description = "0-based slide index." }, ["shapeId"] = new { type = "string", description = "Shape ID of the group to ungroup." } }, required = new[] { "instanceId", "slideIndex", "shapeId" } } },

        // ── Phase 18: Word Bookmarks ────────────────────────────────────
        new { name = "word_get_bookmarks", description = "Returns all bookmarks in the document with names and ranges.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." } }, required = new[] { "instanceId" } } },
        new { name = "word_insert_bookmark", description = "Creates or updates a bookmark at the specified paragraph range. Enables navigation and cross-references.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["name"] = new { type = "string", description = "Bookmark name (no spaces)." }, ["fromParagraph"] = new { type = "number", description = "0-based start paragraph index." }, ["toParagraph"] = new { type = "number", description = "0-based end paragraph index (inclusive). Default: same as fromParagraph." } }, required = new[] { "instanceId", "name", "fromParagraph" } } },
        new { name = "word_delete_bookmark", description = "Deletes a bookmark by name. Does not delete the text content.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["name"] = new { type = "string", description = "Bookmark name to delete." } }, required = new[] { "instanceId", "name" } } },
        new { name = "word_goto_bookmark", description = "Navigates to and selects the content of a bookmark.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["name"] = new { type = "string", description = "Bookmark name to navigate to." } }, required = new[] { "instanceId", "name" } } },

        // ── Phase 18: Word Properties & Hyperlinks ─────────────────────
        new { name = "word_get_properties", description = "Returns built-in document properties: title, author, subject, keywords, company, category, manager, creation date, etc.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." } }, required = new[] { "instanceId" } } },
        new { name = "word_set_properties", description = "Updates document properties (title, author, subject, keywords, etc.). Undoable.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["title"] = new { type = "string", description = "Document title." }, ["author"] = new { type = "string", description = "Document author." }, ["subject"] = new { type = "string", description = "Document subject." }, ["keywords"] = new { type = "string", description = "Document keywords (comma-separated)." }, ["category"] = new { type = "string", description = "Document category." }, ["company"] = new { type = "string", description = "Company name." }, ["manager"] = new { type = "string", description = "Manager name." }, ["comments"] = new { type = "string", description = "Document comments (metadata, not inline comments)." } }, required = new[] { "instanceId" } } },
        new { name = "word_get_hyperlinks", description = "Returns all hyperlinks in the document with their text and URLs.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." } }, required = new[] { "instanceId" } } },
        new { name = "word_insert_hyperlink", description = "Inserts a hyperlink at the specified location. Uses tracked changes.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["text"] = new { type = "string", description = "Display text for the hyperlink." }, ["url"] = new { type = "string", description = "URL for the hyperlink." }, ["paragraphIndex"] = new { type = "number", description = "0-based paragraph index to insert at. Default: end of document." } }, required = new[] { "instanceId", "text", "url" } } },

        // ── Phase 18: Word Footnotes/Endnotes & Fields ──────────────────
        new { name = "word_insert_footnote", description = "Inserts a footnote after the specified paragraph. Uses tracked changes.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["paragraphIndex"] = new { type = "number", description = "0-based paragraph index." }, ["text"] = new { type = "string", description = "Footnote text." } }, required = new[] { "instanceId", "paragraphIndex", "text" } } },
        new { name = "word_insert_endnote", description = "Inserts an endnote after the specified paragraph. Uses tracked changes.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["paragraphIndex"] = new { type = "number", description = "0-based paragraph index." }, ["text"] = new { type = "string", description = "Endnote text." } }, required = new[] { "instanceId", "paragraphIndex", "text" } } },
        new { name = "word_insert_field", description = "Inserts a field (TOC, page number, date, etc.) at the specified location.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["fieldType"] = new { type = "string", description = "Field type: 'TableOfContents', 'PageNumber', 'NumPages', 'Date', 'Time', 'FileName', 'Author', 'StyleRef'." }, ["paragraphIndex"] = new { type = "number", description = "0-based paragraph index to insert at. Default: end of document." } }, required = new[] { "instanceId", "fieldType" } } },

        // ── Phase 18: Word Content Controls ─────────────────────────────
        new { name = "word_get_content_controls", description = "Returns all content controls in the document with their titles, tags, and content.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." } }, required = new[] { "instanceId" } } },
        new { name = "word_insert_content_control", description = "Wraps a paragraph range in a content control for structured editing. Uses tracked changes.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["title"] = new { type = "string", description = "Content control title." }, ["tag"] = new { type = "string", description = "Content control tag for identification." }, ["fromParagraph"] = new { type = "number", description = "0-based start paragraph index." }, ["toParagraph"] = new { type = "number", description = "0-based end paragraph index (inclusive). Default: same as fromParagraph." } }, required = new[] { "instanceId", "title", "fromParagraph" } } },

        // ── Word Formatting ─────────────────────────────────────────────
        new { name = "word_get_formatting", description = "Returns both paragraph and font formatting for one or more paragraphs. Paragraph properties: style, alignment, indents, line spacing, space before/after, outline level. Font properties: name, size, bold, italic, underline, color (hex), strikeThrough, doubleStrikeThrough, subscript, superscript, highlightColor.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["paragraphIndex"] = new { type = "number", description = "0-based paragraph index. If omitted, uses 0." }, ["fromParagraph"] = new { type = "number", description = "0-based start paragraph for a range. Overrides paragraphIndex." }, ["toParagraph"] = new { type = "number", description = "0-based end paragraph (inclusive). Defaults to fromParagraph." } }, required = new[] { "instanceId" } } },
        new { name = "word_set_formatting", description = "Applies paragraph and/or font formatting to one or more paragraphs, or to the current selection. Pass only the properties you want to change — omitted properties are left unchanged. Paragraph props: style, alignment ('Left','Center','Right','Justified'), firstLineIndent, leftIndent, rightIndent (points), lineSpacing, spaceBefore, spaceAfter (points). Font props: fontName, fontSize (pt), bold, italic, underline ('None','Single','Double','Thick','Dotted','DottedHeavy','DashLine','DashLineHeavy','DashLineLong','DashLineLongHeavy','DotDashLine','DotDashLineHeavy','Dot2DashLine','Dot2DashLineHeavy','Wave','WaveHeavy','WaveDouble'), color (hex e.g. '#FF0000'), strikeThrough, doubleStrikeThrough, subscript, superscript, highlightColor ('Yellow','BrightGreen','Turquoise','Pink','Blue','Red','DarkBlue','Teal','Green','Violet','DarkRed','DarkYellow','Gray50','Gray25','Black','White','None').", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["paragraphIndex"] = new { type = "number", description = "0-based paragraph index. Mutually exclusive with applyToSelection." }, ["fromParagraph"] = new { type = "number", description = "0-based start paragraph for a range." }, ["toParagraph"] = new { type = "number", description = "0-based end paragraph (inclusive)." }, ["applyToSelection"] = new { type = "boolean", description = "Apply to the current user selection instead of a paragraph range. Default: false." }, ["style"] = new { type = "string", description = "Named paragraph style (e.g. 'Heading 1', 'Normal', 'Title')." }, ["alignment"] = new { type = "string", description = "Paragraph alignment: 'Left', 'Center', 'Right', 'Justified'." }, ["firstLineIndent"] = new { type = "number", description = "First-line indent in points. Negative = hanging indent." }, ["leftIndent"] = new { type = "number", description = "Left indent in points." }, ["rightIndent"] = new { type = "number", description = "Right indent in points." }, ["lineSpacing"] = new { type = "number", description = "Line spacing in points (e.g. 24 = double-spaced at 12pt base)." }, ["spaceBefore"] = new { type = "number", description = "Space before paragraph in points." }, ["spaceAfter"] = new { type = "number", description = "Space after paragraph in points." }, ["fontName"] = new { type = "string", description = "Font family name (e.g. 'Calibri', 'Times New Roman')." }, ["fontSize"] = new { type = "number", description = "Font size in points." }, ["bold"] = new { type = "boolean", description = "Bold on/off." }, ["italic"] = new { type = "boolean", description = "Italic on/off." }, ["underline"] = new { type = "string", description = "Underline style. 'None' to remove." }, ["color"] = new { type = "string", description = "Font color as hex string, e.g. '#FF0000'. Use '' or '#000000' for black." }, ["strikeThrough"] = new { type = "boolean", description = "Single strikethrough on/off." }, ["doubleStrikeThrough"] = new { type = "boolean", description = "Double strikethrough on/off." }, ["subscript"] = new { type = "boolean", description = "Subscript on/off." }, ["superscript"] = new { type = "boolean", description = "Superscript on/off." }, ["highlightColor"] = new { type = "string", description = "Highlight color name. 'None' to remove." } }, required = new[] { "instanceId" } } },

        // ── Word Comments ────────────────────────────────────────────────
        new { name = "word_get_comments", description = "Returns all comment threads in the document. Each thread contains the top-level comment (id, author, date, text, resolved, anchorText — the text the comment is anchored to) and its replies. Use commentId and replyId values from this response for all other comment tools.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["includeResolved"] = new { type = "boolean", description = "Include resolved comment threads. Default: true.", @default = true } }, required = new[] { "instanceId" } } },
        new { name = "word_edit_comment", description = "Edits the text of an existing top-level comment. Use word_get_comments to obtain the commentId.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["commentId"] = new { type = "string", description = "REQUIRED. The comment ID from word_get_comments." }, ["text"] = new { type = "string", description = "REQUIRED. New comment text." } }, required = new[] { "instanceId", "commentId", "text" } } },
        new { name = "word_resolve_comment", description = "Resolves or reopens a comment thread. Resolving marks the thread as done; reopening restores it for further discussion.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["commentId"] = new { type = "string", description = "REQUIRED. The comment ID from word_get_comments." }, ["resolved"] = new { type = "boolean", description = "true to resolve, false to reopen. Default: true.", @default = true } }, required = new[] { "instanceId", "commentId" } } },
        new { name = "word_delete_comment", description = "Deletes a comment thread and all its replies. This is permanent and cannot be undone.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["commentId"] = new { type = "string", description = "REQUIRED. The comment ID from word_get_comments." } }, required = new[] { "instanceId", "commentId" } } },
        new { name = "word_reply_to_comment", description = "Adds a new reply to an existing comment thread.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["commentId"] = new { type = "string", description = "REQUIRED. The comment ID to reply to." }, ["text"] = new { type = "string", description = "REQUIRED. Reply text." } }, required = new[] { "instanceId", "commentId", "text" } } },
        new { name = "word_edit_reply", description = "Edits the text of a specific reply in a comment thread. Use word_get_comments to obtain both commentId and replyId.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["commentId"] = new { type = "string", description = "REQUIRED. The parent comment ID." }, ["replyId"] = new { type = "string", description = "REQUIRED. The reply ID from word_get_comments replies array." }, ["text"] = new { type = "string", description = "REQUIRED. New reply text." } }, required = new[] { "instanceId", "commentId", "replyId", "text" } } },
        new { name = "word_delete_reply", description = "Deletes a single reply from a comment thread. The parent comment and other replies are not affected.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["commentId"] = new { type = "string", description = "REQUIRED. The parent comment ID." }, ["replyId"] = new { type = "string", description = "REQUIRED. The reply ID from word_get_comments replies array." } }, required = new[] { "instanceId", "commentId", "replyId" } } },
        new { name = "word_get_styles", description = "Returns all styles defined in the document with their font and paragraph format properties. Filter by type ('Paragraph', 'Character', 'Table', 'List') or set inUseOnly=true to return only styles currently applied in the document. Use this before word_modify_style or word_create_style to inspect existing definitions.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["type"] = new { type = "string", description = "Filter by style type: 'Paragraph', 'Character', 'Table', or 'List'. Omit to return all." }, ["inUseOnly"] = new { type = "boolean", description = "Return only styles currently applied somewhere in the document. Default: false." } }, required = new[] { "instanceId" } } },
        new { name = "word_modify_style", description = "Modifies the definition of an existing named style. Pass only the properties you want to change — omitted properties are left unchanged. Changes apply document-wide. Font props: fontName, fontSize, bold, italic, underline, color, strikeThrough, doubleStrikeThrough, subscript, superscript. Paragraph format: alignment, firstLineIndent, leftIndent, rightIndent, lineSpacing, spaceBefore, spaceAfter. Note: list numbering cannot be applied to styles via Office JS on macOS — use word_create_and_remap_style to clone a style with different formatting.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["styleName"] = new { type = "string", description = "REQUIRED. Name of the existing style to modify." }, ["fontName"] = new { type = "string", description = "Font family name." }, ["fontSize"] = new { type = "number", description = "Font size in points." }, ["bold"] = new { type = "boolean", description = "Bold on/off." }, ["italic"] = new { type = "boolean", description = "Italic on/off." }, ["underline"] = new { type = "string", description = "Underline style, e.g. 'Single', 'None'." }, ["color"] = new { type = "string", description = "Hex color, e.g. '#FF0000'." }, ["strikeThrough"] = new { type = "boolean", description = "Single strikethrough on/off." }, ["doubleStrikeThrough"] = new { type = "boolean", description = "Double strikethrough on/off." }, ["subscript"] = new { type = "boolean", description = "Subscript on/off." }, ["superscript"] = new { type = "boolean", description = "Superscript on/off." }, ["alignment"] = new { type = "string", description = "'Left', 'Center', 'Right', 'Justified'." }, ["firstLineIndent"] = new { type = "number", description = "First-line indent in points." }, ["leftIndent"] = new { type = "number", description = "Left indent in points." }, ["rightIndent"] = new { type = "number", description = "Right indent in points." }, ["lineSpacing"] = new { type = "number", description = "Line spacing in points." }, ["spaceBefore"] = new { type = "number", description = "Space before paragraph in points." }, ["spaceAfter"] = new { type = "number", description = "Space after paragraph in points." } }, required = new[] { "instanceId", "styleName" } } },
        new { name = "word_create_style", description = "Creates a new named style in the document (requires WordApi 1.5, desktop only). Specify styleType as 'Paragraph' (default) or 'Character'. Optionally inherit from a baseStyle (e.g. 'Normal'). Then set font and paragraph format properties the same way as word_modify_style.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["styleName"] = new { type = "string", description = "REQUIRED. Name for the new style. Must be unique in the document." }, ["styleType"] = new { type = "string", description = "'Paragraph' (default) or 'Character'." }, ["baseStyle"] = new { type = "string", description = "Name of an existing style to inherit from (e.g. 'Normal', 'Heading 1')." }, ["fontName"] = new { type = "string", description = "Font family name." }, ["fontSize"] = new { type = "number", description = "Font size in points." }, ["bold"] = new { type = "boolean", description = "Bold on/off." }, ["italic"] = new { type = "boolean", description = "Italic on/off." }, ["underline"] = new { type = "string", description = "Underline style, e.g. 'Single', 'None'." }, ["color"] = new { type = "string", description = "Hex color, e.g. '#FF0000'." }, ["strikeThrough"] = new { type = "boolean", description = "Single strikethrough on/off." }, ["doubleStrikeThrough"] = new { type = "boolean", description = "Double strikethrough on/off." }, ["subscript"] = new { type = "boolean", description = "Subscript on/off." }, ["superscript"] = new { type = "boolean", description = "Superscript on/off." }, ["alignment"] = new { type = "string", description = "'Left', 'Center', 'Right', 'Justified'." }, ["firstLineIndent"] = new { type = "number", description = "First-line indent in points." }, ["leftIndent"] = new { type = "number", description = "Left indent in points." }, ["rightIndent"] = new { type = "number", description = "Right indent in points." }, ["lineSpacing"] = new { type = "number", description = "Line spacing in points." }, ["spaceBefore"] = new { type = "number", description = "Space before paragraph in points." }, ["spaceAfter"] = new { type = "number", description = "Space after paragraph in points." } }, required = new[] { "instanceId", "styleName" } } },
        new { name = "word_create_and_remap_style", description = "Creates a new paragraph style derived from an existing one, applies font and paragraph format overrides, then remaps all existing paragraphs that use the base style to the new style. Useful for customising a built-in style (e.g. Heading 1) without modifying its definition. Note: list numbering cannot be baked into styles via Office JS on macOS.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["baseStyleName"] = new { type = "string", description = "REQUIRED. The existing style to clone (e.g. 'Heading 1')." }, ["newStyleName"] = new { type = "string", description = "REQUIRED. Name for the new style." }, ["fontName"] = new { type = "string", description = "Font family name override." }, ["fontSize"] = new { type = "number", description = "Font size in points." }, ["bold"] = new { type = "boolean", description = "Bold on/off." }, ["italic"] = new { type = "boolean", description = "Italic on/off." }, ["underline"] = new { type = "string", description = "Underline style." }, ["color"] = new { type = "string", description = "Hex color, e.g. '#FF0000'." }, ["alignment"] = new { type = "string", description = "'Left', 'Center', 'Right', 'Justified'." }, ["firstLineIndent"] = new { type = "number", description = "First-line indent in points." }, ["leftIndent"] = new { type = "number", description = "Left indent in points." }, ["rightIndent"] = new { type = "number", description = "Right indent in points." }, ["lineSpacing"] = new { type = "number", description = "Line spacing in points." }, ["spaceBefore"] = new { type = "number", description = "Space before paragraph in points." }, ["spaceAfter"] = new { type = "number", description = "Space after paragraph in points." } }, required = new[] { "instanceId", "baseStyleName", "newStyleName" } } },

        // ── Phase 18: Excel Protection & Page Layout ───────────────────
        new { name = "excel_protect_sheet", description = "Protects a worksheet with optional password and granular permissions (allow sort, filter, insert rows, etc.).", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sheetName"] = new { type = "string", description = "Worksheet name." }, ["password"] = new { type = "string", description = "Optional protection password." }, ["allowSort"] = new { type = "boolean", description = "Allow sorting. Default: false." }, ["allowAutoFilter"] = new { type = "boolean", description = "Allow autofilter. Default: false." }, ["allowInsertRows"] = new { type = "boolean", description = "Allow inserting rows. Default: false." }, ["allowInsertColumns"] = new { type = "boolean", description = "Allow inserting columns. Default: false." }, ["allowDeleteRows"] = new { type = "boolean", description = "Allow deleting rows. Default: false." }, ["allowDeleteColumns"] = new { type = "boolean", description = "Allow deleting columns. Default: false." }, ["allowFormatCells"] = new { type = "boolean", description = "Allow formatting cells. Default: false." }, ["allowPivotTables"] = new { type = "boolean", description = "Allow pivot tables. Default: false." } }, required = new[] { "instanceId", "sheetName" } } },
        new { name = "excel_unprotect_sheet", description = "Removes protection from a worksheet. Requires password if one was set.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sheetName"] = new { type = "string", description = "Worksheet name." }, ["password"] = new { type = "string", description = "Password (required if sheet is password-protected)." } }, required = new[] { "instanceId", "sheetName" } } },
        new { name = "excel_set_page_layout", description = "Sets page layout options: orientation, paper size, margins, print area, title rows. Undoable.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sheetName"] = new { type = "string", description = "Worksheet name." }, ["orientation"] = new { type = "string", description = "'Portrait' or 'Landscape'." }, ["paperSize"] = new { type = "string", description = "'Letter', 'A4', 'A3', 'Legal', 'Tabloid', etc." }, ["printArea"] = new { type = "string", description = "A1-style range for print area." }, ["printTitleRows"] = new { type = "string", description = "Row range to repeat on every page (e.g. '$1:$1')." }, ["centerHorizontally"] = new { type = "boolean", description = "Center page horizontally. Default: false." }, ["centerVertically"] = new { type = "boolean", description = "Center page vertically. Default: false." } }, required = new[] { "instanceId", "sheetName" } } },
        new { name = "excel_get_page_layout", description = "Returns current page layout settings for a worksheet.", inputSchema = new { type = "object", properties = new Dictionary<string, object> { ["instanceId"] = new { type = "string", description = "REQUIRED. The instance ID." }, ["sheetName"] = new { type = "string", description = "Worksheet name." } }, required = new[] { "instanceId", "sheetName" } } }
    ];

    /// <summary>
    /// All command names that should be dispatched to the add-in.
    /// </summary>
    private static readonly HashSet<string> AddInCommands = new(StringComparer.OrdinalIgnoreCase)
    {
        "powerpoint_get_deck_outline",
        "powerpoint_get_slide",
        "powerpoint_get_slide_image",
        "powerpoint_get_shape_image",
        "powerpoint_get_table",
        "powerpoint_get_selection",
        "powerpoint_get_speaker_notes",
        "powerpoint_update_shape_text",
        "powerpoint_update_shape_properties",
        "powerpoint_update_speaker_notes",
        "powerpoint_add_textbox",
        "powerpoint_add_image",
        "powerpoint_add_table",
        "powerpoint_delete_shape",
        "powerpoint_add_slide",
        "powerpoint_delete_slide",
        "powerpoint_move_slide",

        // Word
        "word_get_outline",
        "word_get_paragraphs",
        "word_get_selection",
        "word_search",
        "word_replace_text",
        "word_insert_text",
        "word_add_comment",
        "word_delete_paragraph",

        // Word tracked changes
        "word_get_tracked_changes",
        "word_accept_all_changes",
        "word_reject_all_changes",

        // Word Structure
        "word_get_tables",
        "word_insert_table",
        "word_update_table_cell",
        "word_get_headers_footers",
        "word_set_header_footer",
        "word_replace_selection",
        "word_insert_image",
        "word_apply_style",
        "word_get_sections",
        "word_insert_list",

        // Excel
        "excel_get_workbook_map",
        "excel_read_range",
        "excel_write_range",
        "excel_write_formula",
        "excel_create_table",

        // Excel Analysis
        "excel_add_sheet",
        "excel_delete_sheet",
        "excel_rename_sheet",
        "excel_sort_range",
        "excel_filter_range",
        "excel_create_chart",
        "excel_get_charts",
        "excel_format_range",
        "excel_apply_conditional_formatting",
        "excel_create_pivottable",

        // Outlook
        "outlook_get_current_item",
        "outlook_summarize_thread",
        "outlook_draft_reply",
        "outlook_apply_category",
        "outlook_send_message",

        // Outlook Extended (Office.js)
        "outlook_get_user_profile",
        "outlook_get_master_categories",
        "outlook_create_category",
        "outlook_remove_categories",
        "outlook_display_new_message",
        "outlook_display_new_appointment",
        "outlook_get_attachments",

        // Phase 13: Document Export
        "office_export_document",

        // Phase 14: Word Find & Replace
        "word_find_replace",

        // Phase 16: Excel Navigation
        "excel_freeze_panes",
        "excel_get_named_ranges",
        "excel_add_named_range",

        // Phase 17: Excel Validation
        "excel_add_data_validation",
        "excel_remove_data_validation",

        // Phase 18: PowerPoint Tags & Shape Formatting
        "powerpoint_get_tags",
        "powerpoint_set_tag",
        "powerpoint_delete_slides_by_tag",
        "powerpoint_set_shape_fill",
        "powerpoint_set_shape_line",
        "powerpoint_set_shape_rotation",
        "powerpoint_add_geometric_shape",
        "powerpoint_add_line",
        "powerpoint_insert_slides_from_file",
        "powerpoint_get_layouts",
        "powerpoint_get_theme_colors",
        "powerpoint_group_shapes",
        "powerpoint_ungroup_shape",

        // Phase 18: Word Bookmarks, Properties, Hyperlinks, Notes, Fields
        "word_get_bookmarks",
        "word_insert_bookmark",
        "word_delete_bookmark",
        "word_goto_bookmark",
        "word_get_properties",
        "word_set_properties",
        "word_get_hyperlinks",
        "word_insert_hyperlink",
        "word_insert_footnote",
        "word_insert_endnote",
        "word_insert_field",
        "word_get_content_controls",
        "word_insert_content_control",
        "word_get_formatting",
        "word_set_formatting",
        "word_get_comments",
        "word_edit_comment",
        "word_resolve_comment",
        "word_delete_comment",
        "word_reply_to_comment",
        "word_edit_reply",
        "word_delete_reply",
        "word_get_styles",
        "word_modify_style",
        "word_create_style",
        "word_create_and_remap_style",

        // Phase 18: Excel Protection & Page Layout
        "excel_protect_sheet",
        "excel_unprotect_sheet",
        "excel_set_page_layout",
        "excel_get_page_layout",
    };


    /// <summary>
    /// Executes an MCP tool call. Routes to the appropriate add-in instance.
    /// </summary>
    public static async Task<object> ExecuteTool(string name, JsonElement? args)
    {
        // office_get_active_apps is handled server-side
        if (name is "office_get_active_apps" or "office_get_active_app")
            return HandleGetActiveApps();

        // Cross-cutting tools handled server-side
        if (name == "office_get_document_context")
            return HandleGetDocumentContext(instanceId: args.HasValue && args.Value.TryGetProperty("instanceId", out var ctxIid) ? ctxIid.GetString() : null);

        if (name == "office_get_document_stats")
        {
            string? statsInstanceId = null;
            if (args.HasValue && args.Value.TryGetProperty("instanceId", out var statsIid))
                statsInstanceId = statsIid.GetString();
            if (string.IsNullOrEmpty(statsInstanceId))
                return new ToolError("Missing required parameter: instanceId.", ErrorCodes.MISSING_PARAMETER, new { parameter = "instanceId" }).ToMcpResponse();
            return await DispatchToAddIn(statsInstanceId, name, args!);
        }

        if (name == "office_batch_call")
            return await HandleBatchCall(args);

        if (name == "office_suggest_tools")
            return HandleSuggestTools(args);

        // All other tools require an explicit instanceId
        string? instanceId = null;
        if (args.HasValue && args.Value.TryGetProperty("instanceId", out var iid))
            instanceId = iid.GetString();

        return await DispatchToAddIn(instanceId, name, args);
    }

    /// <summary>
    /// Validates instanceId + dispatches a command to the add-in instance.
    /// Returns error responses for missing/invalid instanceId or unknown tools.
    /// </summary>
    private static async Task<object> DispatchToAddIn(string? instanceId, string name, JsonElement? args)
    {
        if (string.IsNullOrEmpty(instanceId))
        {
            return new ToolError(
                "Missing required parameter: instanceId. Call office_get_active_apps first to get the list of available instances.",
                ErrorCodes.MISSING_PARAMETER,
                new { parameter = "instanceId" }
            ).ToMcpResponse();
        }

        // Check if instance exists
        var instance = _registry.GetInstance(instanceId);
        if (instance == null)
        {
            return new ToolError(
                $"Instance '{instanceId}' is not registered or has timed out. Call office_get_active_apps to see current instances.",
                ErrorCodes.INSTANCE_NOT_FOUND,
                new { instanceId }
            ).ToMcpResponse();
        }

        // Validate tool name
        if (!AddInCommands.Contains(name))
        {
            return new ToolError(
                $"Unknown tool: {name}",
                ErrorCodes.INVALID_PARAMETER,
                new { parameter = "toolName", value = name }
            ).ToMcpResponse();
        }

        // Dispatch command to add-in
        var inputs = args.HasValue ? JsonSerializer.Serialize(args.Value) : "{}";
        string commandId = Guid.NewGuid().ToString("N")[..8];
        var cmd = new PendingCommand
        {
            Id = commandId,
            InstanceId = instanceId,
            Command = name,
            Args = args.HasValue ? JsonSerializer.Deserialize<object>(args.Value) : null,
            CreatedAt = DateTime.UtcNow,
        };
        _commandStore.AddCommand(cmd);

        // Push command via SignalR if hub is available
        if (_hubContext != null)
        {
            try
            {
                await _hubContext.Clients.Group(instanceId)
                    .SendAsync("ExecuteCommand", commandId, name, cmd.Args);
                // Mark claimed so HTTP polling doesn't double-execute within the stale window
                _commandStore.MarkClaimed(commandId);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SignalR push failed for {instanceId}: {ex.Message}");
                // Don't mark claimed — let HTTP polling pick it up immediately
            }
        }

        _auditLog.Log(new AuditEntry
        {
            ToolName = name,
            InstanceId = instanceId,
            Inputs = inputs,
            Outcome = "pending"
        });

        var result = await _commandStore.WaitForResult(commandId, timeoutSeconds: IsImageTool(name) ? 120 : 60);
        return BuildToolResult(result, name, instanceId, inputs);
    }

    private static object HandleGetActiveApps()
    {
        var instances = _registry.GetActiveInstances();

        _auditLog.Log(new AuditEntry
        {
            ToolName = "office_get_active_apps",
            InstanceId = "",
            Inputs = "",
            Outcome = "success"
        });

        var appList = instances.Select(i => new
        {
            instanceId = i.InstanceId,
            appName = i.AppName,
            documentName = i.DocumentName
        }).ToList();

        return new
        {
            content = new[] { new { type = "text", text = JsonSerializer.Serialize(new { apps = appList, total = appList.Count }, new JsonSerializerOptions { WriteIndented = true }) } },
            isError = false
        };
    }

    /// <summary>
    /// Returns unified document context for one or all active instances.
    /// </summary>
    private static object HandleGetDocumentContext(string? instanceId)
    {
        var instances = _registry.GetActiveInstances();

        List<OfficeInstance> targets;
        if (!string.IsNullOrEmpty(instanceId))
        {
            var single = _registry.GetInstance(instanceId);
            if (single == null)
            {
                return new ToolError(
                    $"Instance '{instanceId}' is not registered or has timed out.",
                    ErrorCodes.INSTANCE_NOT_FOUND,
                    new { instanceId }
                ).ToMcpResponse();
            }
            targets = [single];
        }
        else
        {
            targets = instances;
        }

        var contexts = targets.Select(i =>
        {
            var hostType = GetHostType(i.AppName);
            return new
            {
                instanceId = i.InstanceId,
                appName = i.AppName,
                documentName = i.DocumentName,
                hostType,
                connectedAt = i.RegisteredAt.ToString("o"),
                lastHeartbeat = i.LastHeartbeat.ToString("o"),
                isAlive = i.IsAlive
            };
        }).ToList();

        return new
        {
            content = new[] { new { type = "text", text = JsonSerializer.Serialize(new
            {
                timestamp = DateTime.UtcNow.ToString("o"),
                totalInstances = contexts.Count,
                contexts
            }, new JsonSerializerOptions { WriteIndented = true }) } },
            isError = false
        };
    }

    /// <summary>
    /// Returns tool suggestions for the active host.
    /// </summary>
    private static object HandleSuggestTools(JsonElement? args)
    {
        string? instanceId = null;
        string? category = null;
        if (args.HasValue)
        {
            if (args.Value.TryGetProperty("instanceId", out var iid))
                instanceId = iid.GetString();
            if (args.Value.TryGetProperty("category", out var cat))
                category = cat.GetString();
        }

        // Determine host type from instance
        string hostType = "all";
        if (!string.IsNullOrEmpty(instanceId))
        {
            var instance = _registry.GetInstance(instanceId);
            if (instance != null)
                hostType = GetHostType(instance.AppName);
        }

        var tools = GetToolDefinitions();
        var suggestions = new List<object>();

        foreach (var tool in tools)
        {
            var json = JsonSerializer.Serialize(tool);
            var doc = JsonDocument.Parse(json);
            var name = doc.RootElement.GetProperty("name").GetString()!;
            var desc = doc.RootElement.GetProperty("description").GetString()!;

            // Filter by host
            if (hostType != "all")
            {
                var toolPrefix = name.Split('_')[0];
                if (toolPrefix != "office" && toolPrefix != hostType) continue;
            }

            // Filter by category
            if (!string.IsNullOrEmpty(category))
            {
                var isRead = name.Contains("get") || name == "office_get_active_apps" || name == "office_get_document_context" || name == "office_get_document_stats" || name == "office_suggest_tools";
                if (category.Equals("read", StringComparison.OrdinalIgnoreCase) && !isRead) continue;
                if (category.Equals("write", StringComparison.OrdinalIgnoreCase) && isRead) continue;
            }

            suggestions.Add(new { name, description = desc.Length > 100 ? desc[..100] + "..." : desc });
        }

        return new
        {
            content = new[] { new { type = "text", text = JsonSerializer.Serialize(new
            {
                hostType,
                category = category ?? "all",
                totalTools = suggestions.Count,
                tools = suggestions
            }, new JsonSerializerOptions { WriteIndented = true }) } },
            isError = false
        };
    }

    /// <summary>
    /// Executes multiple tool calls in parallel.
    /// </summary>
    private static async Task<object> HandleBatchCall(JsonElement? args)
    {
        if (!args.HasValue || !args.Value.TryGetProperty("calls", out var callsArr))
        {
            return new ToolError("Missing required parameter: calls.", ErrorCodes.MISSING_PARAMETER, new { parameter = "calls" }).ToMcpResponse();
        }

        var callsList = callsArr.EnumerateArray().ToList();
        if (callsList.Count == 0)
        {
            return new ToolError("Calls array must not be empty.", ErrorCodes.INVALID_PARAMETER, new { parameter = "calls", value = "empty" }).ToMcpResponse();
        }

        if (callsList.Count > 10)
        {
            return new ToolError("Batch size exceeds maximum of 10 operations.", ErrorCodes.INVALID_PARAMETER, new { parameter = "calls", value = callsList.Count, max = 10 }).ToMcpResponse();
        }

        var tasks = callsList.Select((call, index) =>
        {
            var toolName = call.GetProperty("toolName").GetString() ?? "";
            JsonElement? toolArgs = call.TryGetProperty("args", out var a) ? a : null;
            return ExecuteTool(toolName, toolArgs).ContinueWith(t => new
            {
                index,
                toolName,
                result = t.Result
            });
        }).ToList();

        var results = await Task.WhenAll(tasks);

        return new
        {
            content = new[] { new { type = "text", text = JsonSerializer.Serialize(new
            {
                totalResults = results.Length,
                results = results.OrderBy(r => r.index).Select(r => new
                {
                    r.index,
                    r.toolName,
                    r.result
                }).ToList()
            }, new JsonSerializerOptions { WriteIndented = true }) } },
            isError = false
        };
    }

    /// <summary>
    /// Maps app name to host type prefix.
    /// </summary>
    private static string GetHostType(string appName)
    {
        var lower = appName.ToLowerInvariant();
        if (lower.Contains("word")) return "word";
        if (lower.Contains("excel")) return "excel";
        if (lower.Contains("powerpoint") || lower.Contains("ppt")) return "powerpoint";
        if (lower.Contains("outlook")) return "outlook";
        return "unknown";
    }


    private static object BuildToolResult(PendingCommand? result, string toolName, string instanceId, string inputs)
    {
        if (result == null)
        {
            _auditLog.Log(new AuditEntry
            {
                ToolName = toolName,
                InstanceId = instanceId,
                Inputs = inputs,
                Outcome = "timeout"
            });

            return new ToolError(
                $"Command '{toolName}' timed out waiting for add-in response on instance '{instanceId}'.",
                ErrorCodes.TIMEOUT,
                new { toolName, instanceId, timeoutMs = (IsImageTool(toolName) ? 120 : 60) * 1000 }
            ).ToMcpResponse();
        }

        if (result.Success)
        {
            _auditLog.Log(new AuditEntry
            {
                ToolName = toolName,
                InstanceId = instanceId,
                Inputs = inputs,
                Outcome = "success"
            });

            return new
            {
                content = new[] { new { type = "text", text = JsonSerializer.Serialize(result.Payload, new JsonSerializerOptions { WriteIndented = true }) } },
                isError = false
            };
        }

        _auditLog.Log(new AuditEntry
        {
            ToolName = toolName,
            InstanceId = instanceId,
            Inputs = inputs,
            Outcome = "error",
            Error = result.Error
        });

        // Parse error from add-in to determine appropriate error code
        var errorCode = ParseErrorCode(result.Error);
        return new ToolError(
            $"Command failed: {result.Error}",
            errorCode,
            new { toolName, instanceId, addInError = result.Error }
        ).ToMcpResponse();
    }

    /// <summary>Gets the registry for instance management endpoints.</summary>
    public static InstanceRegistry GetRegistry() => _registry;

    /// <summary>Gets the command store for dispatch/wait operations.</summary>
    public static CommandStore GetCommandStore() => _commandStore;

    /// <summary>
    /// Resets all static state. Used by tests to isolate test runs.
    /// </summary>
    public static void ResetForTesting()
    {
        _registry = new InstanceRegistry();
        _commandStore = new CommandStore();
        _auditLog = new AuditLog(Path.Combine(Path.GetTempPath(), $"audit-test-{Guid.NewGuid()}").TrimEnd('/'));
        _hubContext = null; // Clear hub context for test isolation
    }

    private static bool IsImageTool(string toolName) =>
        toolName is "powerpoint_get_slide_image" or "powerpoint_get_shape_image";

    private static string ParseErrorCode(string? error)
    {
        if (string.IsNullOrEmpty(error)) return ErrorCodes.HOST_NOT_AVAILABLE;

        var lower = error.ToLowerInvariant();

        if (lower.Contains("permission") || lower.Contains("access denied") || lower.Contains("not authorized"))
            return ErrorCodes.PERMISSION_DENIED;
        if (lower.Contains("confirmation") || lower.Contains("confirm"))
            return ErrorCodes.CONFIRMATION_REQUIRED;
        if (lower.Contains("range") && (lower.Contains("too large") || lower.Contains("exceeds")))
            return ErrorCodes.RANGE_TOO_LARGE;
        if (lower.Contains("formula"))
            return ErrorCodes.INVALID_FORMULA;
        if (lower.Contains("not available") || lower.Contains("crash") || lower.Contains("unloaded"))
            return ErrorCodes.HOST_NOT_AVAILABLE;
        if (lower.Contains("timeout") || lower.Contains("timed out"))
            return ErrorCodes.TIMEOUT;
        if (lower.Contains("invalid") && lower.Contains("parameter"))
            return ErrorCodes.INVALID_PARAMETER;

        return ErrorCodes.HOST_NOT_AVAILABLE;
    }
}
