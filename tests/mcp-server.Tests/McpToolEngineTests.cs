using System.Text.Json;
using OfficeMcpServer.Tools;

namespace OfficeMcpServer.Tests;

public class McpToolEngineTests
{
	public McpToolEngineTests()
	{
		McpToolEngine.ResetForTesting();
	}

    [Fact]
    public void GetToolDefinitions_Returns120Tools()
    {
        var tools = McpToolEngine.GetToolDefinitions();

        Assert.NotNull(tools);
        Assert.Equal(120, tools.Length);
    }

    [Fact]
    public void GetToolDefinitions_ContainsAllToolNames()
    {
        var tools = McpToolEngine.GetToolDefinitions();
        var names = tools.Select(t =>
        {
            var json = JsonSerializer.Serialize(t);
            var doc = JsonDocument.Parse(json);
            return doc.RootElement.GetProperty("name").GetString()!;
        }).ToList();

        // Shared
        Assert.Contains("office_get_active_apps", names);

        // Read tools
        Assert.Contains("powerpoint_get_deck_outline", names);
        Assert.Contains("powerpoint_get_slide", names);
        Assert.Contains("powerpoint_get_slide_image", names);
        Assert.Contains("powerpoint_get_shape_image", names);
        Assert.Contains("powerpoint_get_table", names);
        Assert.Contains("powerpoint_get_selection", names);
        Assert.Contains("powerpoint_get_speaker_notes", names);

        // Write tools
        Assert.Contains("powerpoint_update_shape_text", names);
        Assert.Contains("powerpoint_update_shape_properties", names);
        Assert.Contains("powerpoint_update_speaker_notes", names);

        // Shape CRUD
        Assert.Contains("powerpoint_add_textbox", names);
        Assert.Contains("powerpoint_add_image", names);
        Assert.Contains("powerpoint_add_table", names);
        Assert.Contains("powerpoint_delete_shape", names);

        // Slide management
        Assert.Contains("powerpoint_add_slide", names);
        Assert.Contains("powerpoint_delete_slide", names);
        Assert.Contains("powerpoint_move_slide", names);

        // Word tools
        Assert.Contains("word_get_outline", names);
        Assert.Contains("word_get_paragraphs", names);
        Assert.Contains("word_get_selection", names);
        Assert.Contains("word_search", names);
        Assert.Contains("word_replace_text", names);
        Assert.Contains("word_insert_text", names);
        Assert.Contains("word_add_comment", names);
        Assert.Contains("word_delete_paragraph", names);
        Assert.Contains("word_get_tracked_changes", names);
        Assert.Contains("word_accept_all_changes", names);
        Assert.Contains("word_reject_all_changes", names);

        Assert.Contains("excel_get_workbook_map", names);
        Assert.Contains("excel_read_range", names);
        Assert.Contains("excel_write_range", names);
        Assert.Contains("excel_write_formula", names);
        Assert.Contains("excel_create_table", names);

        // Outlook tools
        Assert.Contains("outlook_get_current_item", names);
        Assert.Contains("outlook_summarize_thread", names);
        Assert.Contains("outlook_draft_reply", names);
        Assert.Contains("outlook_apply_category", names);
        Assert.Contains("outlook_send_message", names);

        // Excel Analysis
        Assert.Contains("excel_add_sheet", names);
        Assert.Contains("excel_delete_sheet", names);
        Assert.Contains("excel_rename_sheet", names);
        Assert.Contains("excel_sort_range", names);
        Assert.Contains("excel_filter_range", names);
        Assert.Contains("excel_create_chart", names);
        Assert.Contains("excel_get_charts", names);
        Assert.Contains("excel_format_range", names);
        Assert.Contains("excel_apply_conditional_formatting", names);
        Assert.Contains("excel_create_pivottable", names);

        // Word Structure
        Assert.Contains("word_get_tables", names);
        Assert.Contains("word_insert_table", names);
        Assert.Contains("word_update_table_cell", names);
        Assert.Contains("word_get_headers_footers", names);
        Assert.Contains("word_set_header_footer", names);
        Assert.Contains("word_replace_selection", names);
        Assert.Contains("word_insert_image", names);
        Assert.Contains("word_apply_style", names);
        Assert.Contains("word_get_sections", names);
        Assert.Contains("word_insert_list", names);

        // Outlook Extended
        Assert.Contains("outlook_get_user_profile", names);
        Assert.Contains("outlook_get_master_categories", names);
        Assert.Contains("outlook_create_category", names);
        Assert.Contains("outlook_remove_categories", names);
        Assert.Contains("outlook_display_new_message", names);
        Assert.Contains("outlook_display_new_appointment", names);
        Assert.Contains("outlook_get_attachments", names);

        // Phase 13: Document Export
        Assert.Contains("office_export_document", names);

        // Phase 14: Word Find & Replace
        Assert.Contains("word_find_replace", names);

        // Phase 16: Excel Navigation
        Assert.Contains("excel_freeze_panes", names);
        Assert.Contains("excel_get_named_ranges", names);
        Assert.Contains("excel_add_named_range", names);

        // Phase 17: Excel Validation
        Assert.Contains("excel_add_data_validation", names);
        Assert.Contains("excel_remove_data_validation", names);

        // Phase 18: PowerPoint Tags & Shape Formatting
        Assert.Contains("powerpoint_get_tags", names);
        Assert.Contains("powerpoint_set_tag", names);
        Assert.Contains("powerpoint_delete_slides_by_tag", names);
        Assert.Contains("powerpoint_set_shape_fill", names);
        Assert.Contains("powerpoint_set_shape_line", names);
        Assert.Contains("powerpoint_set_shape_rotation", names);
        Assert.Contains("powerpoint_add_geometric_shape", names);
        Assert.Contains("powerpoint_add_line", names);
        Assert.Contains("powerpoint_insert_slides_from_file", names);
        Assert.Contains("powerpoint_get_layouts", names);
        Assert.Contains("powerpoint_get_theme_colors", names);
        Assert.Contains("powerpoint_group_shapes", names);
        Assert.Contains("powerpoint_ungroup_shape", names);

        // Phase 18: Word Bookmarks, Properties, Hyperlinks, Notes, Fields
        Assert.Contains("word_get_bookmarks", names);
        Assert.Contains("word_insert_bookmark", names);
        Assert.Contains("word_delete_bookmark", names);
        Assert.Contains("word_goto_bookmark", names);
        Assert.Contains("word_get_properties", names);
        Assert.Contains("word_set_properties", names);
        Assert.Contains("word_get_hyperlinks", names);
        Assert.Contains("word_insert_hyperlink", names);
        Assert.Contains("word_insert_footnote", names);
        Assert.Contains("word_insert_endnote", names);
        Assert.Contains("word_insert_field", names);
        Assert.Contains("word_get_content_controls", names);
        Assert.Contains("word_insert_content_control", names);

        // Phase 18: Excel Protection & Page Layout
        Assert.Contains("excel_protect_sheet", names);
        Assert.Contains("excel_unprotect_sheet", names);
        Assert.Contains("excel_set_page_layout", names);
        Assert.Contains("excel_get_page_layout", names);

        // Style management
        Assert.Contains("word_get_styles", names);
        Assert.Contains("word_modify_style", names);
        Assert.Contains("word_create_style", names);
        Assert.Contains("word_create_and_remap_style", names);
    }

    [Fact]
    public void GetToolDefinitions_AllToolsHaveRequiredFields()
    {
        var tools = McpToolEngine.GetToolDefinitions();

        foreach (var tool in tools)
        {
            var json = JsonSerializer.Serialize(tool);
            var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            Assert.True(root.TryGetProperty("name", out var nameEl), "Tool missing 'name'");
            Assert.False(string.IsNullOrEmpty(nameEl.GetString()), "Tool has empty name");

            Assert.True(root.TryGetProperty("description", out var descEl), $"Tool {nameEl.GetString()} missing 'description'");
            Assert.True(descEl.GetString()!.Length > 20, $"Tool {nameEl.GetString()} has too-short description");

            Assert.True(root.TryGetProperty("inputSchema", out _), $"Tool {nameEl.GetString()} missing 'inputSchema'");
        }
    }

    [Fact]
    public void GetToolDefinitions_PowerPointToolsRequireInstanceId()
    {
        var tools = McpToolEngine.GetToolDefinitions();

        foreach (var tool in tools)
        {
            var json = JsonSerializer.Serialize(tool);
            var doc = JsonDocument.Parse(json);
            var name = doc.RootElement.GetProperty("name").GetString()!;

            // These tools don't require instanceId
            if (name is "office_get_active_apps" or "office_get_document_context" or "office_batch_call" or "office_suggest_tools") continue;

            var required = doc.RootElement
                .GetProperty("inputSchema")
                .GetProperty("required")
                .EnumerateArray()
                .Select(r => r.GetString()!)
                .ToList();

            Assert.Contains("instanceId", required);
        }
    }

    [Fact]
    public async Task ExecuteTool_NoInstances_ReturnsEmptyApps()
    {
        var result = await McpToolEngine.ExecuteTool("office_get_active_apps", null);

        var json = JsonSerializer.Serialize(result);
        var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.TryGetProperty("isError", out var isError));
        Assert.False(isError.GetBoolean());
    }

    [Fact]
    public async Task ExecuteTool_UnknownTool_ReturnsError()
    {
        var result = await McpToolEngine.ExecuteTool("nonexistent_tool", null);

        var json = JsonSerializer.Serialize(result);
        var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.TryGetProperty("isError", out var isError));
        Assert.True(isError.GetBoolean());
    }

    [Fact]
    public async Task ExecuteTool_MissingInstanceId_ReturnsError()
    {
        var args = JsonSerializer.Deserialize<JsonElement>("{\"slideIndex\": 0}");
        var result = await McpToolEngine.ExecuteTool("powerpoint_get_slide", args);

        var json = JsonSerializer.Serialize(result);
        var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.TryGetProperty("isError", out var isError));
        Assert.True(isError.GetBoolean());
    }

    [Fact]
    public async Task ExecuteTool_MissingInstanceId_ReturnsErrorCode()
    {
        var args = JsonSerializer.Deserialize<JsonElement>("{\"slideIndex\": 0}");
        var result = await McpToolEngine.ExecuteTool("powerpoint_get_slide", args);

        var json = JsonSerializer.Serialize(result);
        var doc = JsonDocument.Parse(json);
        var content = doc.RootElement.GetProperty("content")[0].GetProperty("text").GetString();
        var errorObj = JsonDocument.Parse(content!).RootElement;

        Assert.True(doc.RootElement.GetProperty("isError").GetBoolean());
        Assert.Equal("MISSING_PARAMETER", errorObj.GetProperty("errorCode").GetString());
        Assert.Equal("instanceId", errorObj.GetProperty("details").GetProperty("parameter").GetString());
    }

    [Fact]
    public async Task ExecuteTool_UnknownInstance_ReturnsErrorCode()
    {
        var args = JsonSerializer.Deserialize<JsonElement>("{\"instanceId\": \"nonexistent_99\"}");
        var result = await McpToolEngine.ExecuteTool("powerpoint_get_deck_outline", args);

        var json = JsonSerializer.Serialize(result);
        var doc = JsonDocument.Parse(json);
        var content = doc.RootElement.GetProperty("content")[0].GetProperty("text").GetString();
        var errorObj = JsonDocument.Parse(content!).RootElement;

        Assert.True(doc.RootElement.GetProperty("isError").GetBoolean());
        Assert.Equal("INSTANCE_NOT_FOUND", errorObj.GetProperty("errorCode").GetString());
        Assert.Equal("nonexistent_99", errorObj.GetProperty("details").GetProperty("instanceId").GetString());
    }

    [Fact]
    public async Task ExecuteTool_UnknownTool_ReturnsErrorCode()
    {
        // Register an instance so we get past MISSING_PARAMETER check
        McpToolEngine.ResetForTesting();
        var registry = McpToolEngine.GetRegistry();
        var iid = registry.RegisterInstance("PowerPoint", "test.pptx");

        var args = JsonSerializer.Deserialize<JsonElement>($"{{\"instanceId\": \"{iid}\"}}");
        var result = await McpToolEngine.ExecuteTool("nonexistent_tool", args);

        var json = JsonSerializer.Serialize(result);
        var doc = JsonDocument.Parse(json);
        var content = doc.RootElement.GetProperty("content")[0].GetProperty("text").GetString();
        var errorObj = JsonDocument.Parse(content!).RootElement;

        Assert.True(doc.RootElement.GetProperty("isError").GetBoolean());
        Assert.Equal("INVALID_PARAMETER", errorObj.GetProperty("errorCode").GetString());
    }
}
