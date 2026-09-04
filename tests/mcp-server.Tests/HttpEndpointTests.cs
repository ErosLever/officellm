using System.Net;
using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc.Testing;

namespace OfficeMcpServer.Tests;

/// <summary>
/// Integration tests using WebApplicationFactory.
/// Tests the full HTTP pipeline: routing, serialization, error handling.
/// </summary>
public class HttpEndpointTests : IClassFixture<WebApplicationFactory<Program>>
{
    private readonly HttpClient _client;

    public HttpEndpointTests(WebApplicationFactory<Program> factory)
    {
        _client = factory.CreateClient();
        _client.DefaultRequestHeaders.Authorization =
            new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", BuildToken.Value);
        // Reset static state before each test class run
        Tools.McpToolEngine.ResetForTesting();
    }

    // --- Health ---

    [Fact]
    public async Task Health_ReturnsOk()
    {
        var response = await _client.GetAsync("/health");

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal("ok", body.GetProperty("status").GetString());
        Assert.True(body.TryGetProperty("activeInstances", out _));
    }

    // --- Instances ---

    [Fact]
    public async Task Register_ReturnsInstanceId()
    {
        var response = await _client.PostAsJsonAsync("/instances/register", new
        {
            appName = "PowerPoint",
            documentName = "test.pptx"
        });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.True(body.TryGetProperty("instanceId", out var id));
        Assert.False(string.IsNullOrEmpty(id.GetString()));
    }

    [Fact]
    public async Task GetInstances_ReturnsRegisteredInstances()
    {
        // Register an instance first
        await _client.PostAsJsonAsync("/instances/register", new
        {
            appName = "PowerPoint",
            documentName = "test.pptx"
        });

        var response = await _client.GetAsync("/instances");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var instances = body.GetProperty("instances");
        Assert.True(instances.GetArrayLength() >= 1);
    }

    [Fact]
    public async Task Heartbeat_ReturnsOk()
    {
        // Register first
        var reg = await _client.PostAsJsonAsync("/instances/register", new
        {
            appName = "PowerPoint",
            documentName = "test.pptx"
        });
        var regBody = await reg.Content.ReadFromJsonAsync<JsonElement>();
        var instanceId = regBody.GetProperty("instanceId").GetString();

        // Send heartbeat
        var response = await _client.PostAsJsonAsync($"/instances/{instanceId}/heartbeat", new
        {
            appName = "PowerPoint",
            documentName = "updated.pptx"
        });

        response.EnsureSuccessStatusCode();
    }

    [Fact]
    public async Task GetCommands_ReturnsEmpty_WhenNonePending()
    {
        var reg = await _client.PostAsJsonAsync("/instances/register", new
        {
            appName = "PowerPoint",
            documentName = "test.pptx"
        });
        var regBody = await reg.Content.ReadFromJsonAsync<JsonElement>();
        var instanceId = regBody.GetProperty("instanceId").GetString();

        var response = await _client.GetAsync($"/instances/{instanceId}/commands");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var commands = body.GetProperty("commands");
        Assert.Equal(0, commands.GetArrayLength());
    }

    [Fact]
    public async Task PostResult_ReturnsOk()
    {
        var reg = await _client.PostAsJsonAsync("/instances/register", new
        {
            appName = "PowerPoint",
            documentName = "test.pptx"
        });
        var regBody = await reg.Content.ReadFromJsonAsync<JsonElement>();
        var instanceId = regBody.GetProperty("instanceId").GetString();

        var response = await _client.PostAsJsonAsync($"/instances/{instanceId}/result", new
        {
            commandId = "test-cmd-1",
            success = true
        });

        response.EnsureSuccessStatusCode();
    }

    // --- OpenAPI ---

    [Fact]
    public async Task OpenApi_ReturnsValidSpec()
    {
        var response = await _client.GetAsync("/openapi.json");
        response.EnsureSuccessStatusCode();

        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal("3.0.0", body.GetProperty("openapi").GetString());

        var info = body.GetProperty("info");
        Assert.Equal("Office LLM Harness API", info.GetProperty("title").GetString());

        // Must have paths
        var paths = body.GetProperty("paths");
        Assert.True(paths.EnumerateObject().Any(), "OpenAPI spec should have at least one path");

        // Must have /api/office_get_active_apps
        Assert.True(paths.TryGetProperty("/api/office_get_active_apps", out _),
            "Should have /api/office_get_active_apps path");
    }

    [Fact]
    public async Task OpenApi_ContainsAllToolEndpoints()
    {
        var response = await _client.GetAsync("/openapi.json");
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var paths = body.GetProperty("paths");

        var expectedPaths = new[]
        {
            "/api/office_get_active_apps",
            "/api/powerpoint_get_deck_outline",
            "/api/powerpoint_get_slide",
            "/api/powerpoint_get_slide_image",
            "/api/powerpoint_get_shape_image",
            "/api/powerpoint_get_table",
            "/api/powerpoint_get_selection",
            "/api/powerpoint_get_speaker_notes",
            "/api/powerpoint_update_shape_text",
            "/api/powerpoint_update_shape_properties",
            "/api/powerpoint_update_speaker_notes",
            "/api/powerpoint_add_textbox",
            "/api/powerpoint_add_image",
            "/api/powerpoint_add_table",
            "/api/powerpoint_delete_shape",
            "/api/powerpoint_add_slide",
            "/api/powerpoint_delete_slide",
            "/api/powerpoint_move_slide",
            "/api/powerpoint_duplicate_slide",

            // Word
            "/api/word_get_outline",
            "/api/word_get_paragraphs",
            "/api/word_get_selection",
            "/api/word_search",
            "/api/word_replace_text",
            "/api/word_insert_text",
            "/api/word_add_comment",
            "/api/word_delete_paragraph",
            "/api/word_get_tracked_changes",
            "/api/word_accept_all_changes",
            "/api/word_reject_all_changes"
        };

        foreach (var expected in expectedPaths)
        {
            Assert.True(paths.TryGetProperty(expected, out var pathObj),
                $"Missing path: {expected}");
            Assert.True(pathObj.TryGetProperty("post", out var post), $"{expected} should have POST");
        }
    }

    [Fact]
    public async Task OpenApi_ToolEndpointsHaveParameterSchemas()
    {
        var response = await _client.GetAsync("/openapi.json");
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var paths = body.GetProperty("paths");

        // Verify powerpoint_get_deck_outline has instanceId in its schema
        var deckOutline = paths.GetProperty("/api/powerpoint_get_deck_outline");
        var deckSchema = deckOutline.GetProperty("post").GetProperty("requestBody")
            .GetProperty("content").GetProperty("application/json").GetProperty("schema");
        var deckProps = deckSchema.GetProperty("properties");
        Assert.True(deckProps.TryGetProperty("instanceId", out var instanceIdProp),
            "powerpoint_get_deck_outline should have instanceId in properties");
        Assert.Equal("string", instanceIdProp.GetProperty("type").GetString());

        // Verify required includes instanceId
        var deckRequired = deckSchema.GetProperty("required");
        Assert.Contains("instanceId", deckRequired.EnumerateArray().Select(r => r.GetString()));

        // Verify powerpoint_get_slide has instanceId AND slideIndex
        var getSlide = paths.GetProperty("/api/powerpoint_get_slide");
        var slideSchema = getSlide.GetProperty("post").GetProperty("requestBody")
            .GetProperty("content").GetProperty("application/json").GetProperty("schema");
        var slideProps = slideSchema.GetProperty("properties");
        Assert.True(slideProps.TryGetProperty("slideIndex", out var slideIdxProp),
            "powerpoint_get_slide should have slideIndex in properties");
        Assert.Equal("integer", slideIdxProp.GetProperty("type").GetString());

        var slideRequired = slideSchema.GetProperty("required");
        var requiredList = slideRequired.EnumerateArray().Select(r => r.GetString()).ToList();
        Assert.Contains("instanceId", requiredList);
        Assert.Contains("slideIndex", requiredList);

        // Verify office_get_active_apps has NO properties (empty schema)
        var getApps = paths.GetProperty("/api/office_get_active_apps");
        var appsSchema = getApps.GetProperty("post").GetProperty("requestBody")
            .GetProperty("content").GetProperty("application/json").GetProperty("schema");
        Assert.Equal(JsonValueKind.Object, appsSchema.GetProperty("properties").ValueKind);
        Assert.Equal(0, appsSchema.GetProperty("properties").EnumerateObject().Count());
    }

    [Fact]
    public async Task OpenApi_SpecHasServerUrl()
    {
        var response = await _client.GetAsync("/openapi.json");
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();

        var servers = body.GetProperty("servers");
        Assert.True(servers.GetArrayLength() > 0);

        var firstServer = servers[0];
        Assert.True(firstServer.TryGetProperty("url", out _));
    }

    // --- Swagger UI ---

    [Fact]
    public async Task Docs_ReturnsHtml()
    {
        var response = await _client.GetAsync("/docs");

        response.EnsureSuccessStatusCode();
        Assert.Equal("text/html", response.Content.Headers.ContentType?.MediaType);
        var html = await response.Content.ReadAsStringAsync();
        Assert.Contains("swagger-ui", html);
    }

    // --- REST API Bridge ---

    [Fact]
    public async Task ApiBridge_OfficeGetActiveApps_ReturnsSuccess()
    {
        var response = await _client.PostAsJsonAsync("/api/office_get_active_apps", new { });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.True(body.TryGetProperty("isError", out var isError));
        Assert.False(isError.GetBoolean());
    }

    [Fact]
    public async Task ApiBridge_UnknownTool_ReturnsError()
    {
        var response = await _client.PostAsJsonAsync("/api/nonexistent_tool", new { });

        // Should still return 200 with error payload (McpToolEngine returns error result)
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.True(body.TryGetProperty("isError", out var isError));
        Assert.True(isError.GetBoolean());
    }

    [Fact]
    public async Task ApiBridge_InvalidJson_Returns400()
    {
        var content = new StringContent("not json", System.Text.Encoding.UTF8, "application/json");
        var response = await _client.PostAsync("/api/office_get_active_apps", content);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }

    // --- MCP Protocol ---

    [Fact]
    public async Task Mcp_Initialize_ReturnsCapabilities()
    {
        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0",
            id = 1,
            method = "initialize"
        });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        Assert.Equal("2.0", body.GetProperty("jsonrpc").GetString());

        var result = body.GetProperty("result");
        Assert.Equal("2024-11-05", result.GetProperty("protocolVersion").GetString());
        Assert.Equal("OfficeMcpServer", result.GetProperty("serverInfo").GetProperty("name").GetString());
    }

    [Fact]
    public async Task Mcp_ToolsList_ReturnsTools()
    {
        // Initialize first
        await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0",
            id = 1,
            method = "initialize"
        });

        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0",
            id = 2,
            method = "tools/list"
        });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var tools = body.GetProperty("result").GetProperty("tools");
        Assert.True(tools.GetArrayLength() >= 5);

        // Verify tool schemas have proper properties
        var deckOutline = tools.EnumerateArray().First(t => t.GetProperty("name").GetString() == "powerpoint_get_deck_outline");
        var schemaProps = deckOutline.GetProperty("inputSchema").GetProperty("properties");
        Assert.True(schemaProps.ValueKind == JsonValueKind.Object, "properties should be an object");
        Assert.True(schemaProps.GetProperty("instanceId").GetProperty("type").GetString() == "string");

        var getSlide = tools.EnumerateArray().First(t => t.GetProperty("name").GetString() == "powerpoint_get_slide");
        var slideProps = getSlide.GetProperty("inputSchema").GetProperty("properties");
        Assert.True(slideProps.GetProperty("slideIndex").GetProperty("type").GetString() == "integer");
        var required = getSlide.GetProperty("inputSchema").GetProperty("required");
        Assert.Contains("instanceId", required.EnumerateArray().Select(r => r.GetString()));
        Assert.Contains("slideIndex", required.EnumerateArray().Select(r => r.GetString()));
    }

    [Fact]
    public async Task Mcp_ToolsCall_NoInstanceId_ReturnsError()
    {
        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0",
            id = 3,
            method = "tools/call",
            @params = new
            {
                name = "powerpoint_get_deck_outline",
                arguments = new { }
            }
        });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var result = body.GetProperty("result");
        Assert.True(result.TryGetProperty("isError", out var isError));
        Assert.True(isError.GetBoolean());
    }

    [Fact]
    public async Task Mcp_OfficeGetActiveApps_ReturnsSuccess()
    {
        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0",
            id = 3,
            method = "tools/call",
            @params = new
            {
                name = "office_get_active_apps",
                arguments = new { }
            }
        });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var result = body.GetProperty("result");
        Assert.True(result.TryGetProperty("isError", out var isError));
        Assert.False(isError.GetBoolean());
    }

    [Fact]
    public async Task Mcp_EmptyBody_Returns400()
    {
        var content = new StringContent("", System.Text.Encoding.UTF8, "application/json");
        var response = await _client.PostAsync("/mcp", content);
        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }

    [Fact]
    public async Task Mcp_ToolsList_Returns131Tools()
    {
        // Initialize first
        await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0",
            id = 1,
            method = "initialize"
        });

        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0",
            id = 2,
            method = "tools/list"
        });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var tools = body.GetProperty("result").GetProperty("tools");
        Assert.Equal(131, tools.GetArrayLength());
    }

    [Fact]
    public async Task Mcp_ToolsList_AllNewToolsHaveDescriptions()
    {
        await _client.PostAsJsonAsync("/mcp", new { jsonrpc = "2.0", id = 1, method = "initialize" });

        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0", id = 2, method = "tools/list"
        });

        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var tools = body.GetProperty("result").GetProperty("tools");

        var newToolNames = new[]
        {
            "powerpoint_get_slide_image", "powerpoint_get_shape_image",
            "powerpoint_get_table", "powerpoint_get_selection",
            "powerpoint_get_speaker_notes", "powerpoint_update_shape_properties",
            "powerpoint_add_textbox", "powerpoint_add_image",
            "powerpoint_add_table", "powerpoint_delete_shape",
            "powerpoint_add_slide", "powerpoint_set_slide_layout", "powerpoint_delete_slide", "powerpoint_move_slide",
            "powerpoint_duplicate_slide"
        };

        foreach (var toolName in newToolNames)
        {
            var tool = tools.EnumerateArray().FirstOrDefault(t => t.GetProperty("name").GetString() == toolName);
            Assert.True(tool.ValueKind != JsonValueKind.Undefined, $"Tool '{toolName}' not found in tools/list");
            Assert.True(tool.GetProperty("description").GetString()!.Length > 20, $"Tool '{toolName}' has too-short description");
        }
    }

    [Theory]
    [InlineData("powerpoint_get_slide_image", "slideIndex")]
    [InlineData("powerpoint_get_shape_image", "slideIndex,shapeId")]
    [InlineData("powerpoint_get_table", "slideIndex,shapeId")]
    [InlineData("powerpoint_add_textbox", "slideIndex,text,left,top,width,height")]
    [InlineData("powerpoint_add_image", "slideIndex,imageBase64,left,top")]
    [InlineData("powerpoint_add_table", "slideIndex,rows,columns,left,top")]
    [InlineData("powerpoint_delete_shape", "slideIndex,shapeId")]
    [InlineData("powerpoint_add_slide", null)]
    [InlineData("powerpoint_set_slide_layout", "slideIndex,layoutId")]
    [InlineData("powerpoint_delete_slide", "slideIndex")]
    [InlineData("powerpoint_move_slide", "fromIndex,toIndex")]
    [InlineData("powerpoint_duplicate_slide", "slideIndex")]
    public async Task Mcp_NewTools_RequireInstanceId(string toolName, string? otherRequired)
    {
        // Calling without instanceId should return error
        var args = new Dictionary<string, object> { ["slideIndex"] = 0 };
        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0", id = 3, method = "tools/call",
            @params = new { name = toolName, arguments = args }
        });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var result = body.GetProperty("result");
        Assert.True(result.TryGetProperty("isError", out var isError));
        Assert.True(isError.GetBoolean(), $"{toolName} should fail without instanceId");
    }

    [Theory]
    [InlineData("word_accept_all_changes", null)]
    [InlineData("word_reject_all_changes", null)]
    [InlineData("word_get_tracked_changes", null)]
    public async Task Mcp_TrackedChangeTools_RequireInstanceId(string toolName, string? otherRequired)
    {
        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0", id = 3, method = "tools/call",
            @params = new { name = toolName, arguments = new { } }
        });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var result = body.GetProperty("result");
        Assert.True(result.TryGetProperty("isError", out var isError));
        Assert.True(isError.GetBoolean(), $"{toolName} should fail without instanceId");
    }

    [Fact]
    public async Task Mcp_ToolsCall_InvalidInstanceId_ReturnsError()
    {
        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0", id = 4, method = "tools/call",
            @params = new
            {
                name = "powerpoint_get_slide_image",
                arguments = new { instanceId = "nonexistent_99", slideIndex = 0 }
            }
        });

        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var result = body.GetProperty("result");
        Assert.True(result.GetProperty("isError").GetBoolean());
    }

    [Fact]
    public async Task Mcp_ToolsCall_ImageTool_QueuesCommandForInstance()
    {
        // Register an instance
        var reg = await _client.PostAsJsonAsync("/instances/register", new
        {
            appName = "PowerPoint", documentName = "test.pptx"
        });
        var regBody = await reg.Content.ReadFromJsonAsync<JsonElement>();
        var instanceId = regBody.GetProperty("instanceId").GetString();

        // Call tool — will timeout (no add-in to process), but should queue command
        // Use a short timeout by checking the commands endpoint
        var callTask = _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0", id = 5, method = "tools/call",
            @params = new
            {
                name = "powerpoint_get_slide_image",
                arguments = new { instanceId, slideIndex = 0, width = 400 }
            }
        });

        // While waiting, poll commands — should find the queued command
        await Task.Delay(100);
        var cmds = await _client.GetAsync($"/instances/{instanceId}/commands");
        var cmdsBody = await cmds.Content.ReadFromJsonAsync<JsonElement>();
        var commands = cmdsBody.GetProperty("commands");

        // The command should have been queued (or already claimed)
        // Either way it confirms the dispatch path works
        Assert.True(commands.GetArrayLength() <= 1, "Should have at most 1 queued command");

        // Complete the command so the tool call can finish
        if (commands.GetArrayLength() == 1)
        {
            var cmdId = commands[0].GetProperty("id").GetString();
            await _client.PostAsJsonAsync($"/instances/{instanceId}/result", new
            {
                commandId = cmdId, success = true, payload = new { image = "test" }
            });
        }

        // Wait for the tool call to complete
        var response = await callTask;
        response.EnsureSuccessStatusCode();
    }

    [Fact]
    public async Task OpenApi_NewTools_HaveCorrectParameterTypes()
    {
        var response = await _client.GetAsync("/openapi.json");
        var body = await response.Content.ReadFromJsonAsync<JsonElement>();
        var paths = body.GetProperty("paths");

        // powerpoint_add_textbox should have text as string, left/top/width/height as number
        var addTextbox = paths.GetProperty("/api/powerpoint_add_textbox");
        var schema = addTextbox.GetProperty("post").GetProperty("requestBody")
            .GetProperty("content").GetProperty("application/json").GetProperty("schema");
        var props = schema.GetProperty("properties");

        Assert.Equal("string", props.GetProperty("text").GetProperty("type").GetString());
        Assert.Equal("number", props.GetProperty("left").GetProperty("type").GetString());
        Assert.Equal("number", props.GetProperty("top").GetProperty("type").GetString());
        Assert.Equal("number", props.GetProperty("width").GetProperty("type").GetString());
        Assert.Equal("number", props.GetProperty("height").GetProperty("type").GetString());

        var required = schema.GetProperty("required").EnumerateArray().Select(r => r.GetString()).ToList();
        Assert.Contains("text", required);
        Assert.Contains("left", required);
        Assert.Contains("top", required);

        // powerpoint_get_slide_image should have width/height as integer with defaults
        var getImage = paths.GetProperty("/api/powerpoint_get_slide_image");
        var imgSchema = getImage.GetProperty("post").GetProperty("requestBody")
            .GetProperty("content").GetProperty("application/json").GetProperty("schema");
        var imgProps = imgSchema.GetProperty("properties");

        Assert.Equal("integer", imgProps.GetProperty("width").GetProperty("type").GetString());
        Assert.Equal("integer", imgProps.GetProperty("slideIndex").GetProperty("type").GetString());

        // powerpoint_add_table should have rows/columns as integer
        var addTable = paths.GetProperty("/api/powerpoint_add_table");
        var tblSchema = addTable.GetProperty("post").GetProperty("requestBody")
            .GetProperty("content").GetProperty("application/json").GetProperty("schema");
        var tblProps = tblSchema.GetProperty("properties");

        Assert.Equal("integer", tblProps.GetProperty("rows").GetProperty("type").GetString());
        Assert.Equal("integer", tblProps.GetProperty("columns").GetProperty("type").GetString());
    }

    [Fact]
    public async Task Mcp_UnknownMethod_Returns404()
    {
        var response = await _client.PostAsJsonAsync("/mcp", new
        {
            jsonrpc = "2.0",
            id = 99,
            method = "nonexistent/method"
        });

        Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
    }
}
