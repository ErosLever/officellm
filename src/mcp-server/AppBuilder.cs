using System.Collections.Concurrent;
using System.Text.Json;
using System.Threading.Channels;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.DependencyInjection;
using OfficeMcpServer.Models;
using OfficeMcpServer.Tools;
using OfficeMcpServer.Hubs;
using Microsoft.AspNetCore.SignalR;
namespace OfficeMcpServer;

/// <summary>
/// Creates and configures the MCP server WebApplication.
/// Separated from Program.cs to enable WebApplicationFactory-based integration testing.
/// </summary>
public static class AppBuilder
{
    // SSE sessions: sessionId → channel used to push JSON-RPC responses back to the client
    private static readonly ConcurrentDictionary<string, Channel<string>> _sseSessions = new();

    /// <summary>
    /// Creates a pre-configured WebApplication. The caller must call app.Run() to start it.
    /// For testing, pass mcpHost=null and mcpPort=null to use ASP.NET's default URL mechanism.
    /// </summary>
    public static WebApplication Create(string? mcpHost = "127.0.0.1", int? mcpPort = null)
    {
        // args are now parsed in Program.cs before calling Create()

        var builder = WebApplication.CreateBuilder();

        // Suppress noisy ASP.NET request logging (polling generates 6+ lines per request)
        builder.Logging.AddFilter("Microsoft.AspNetCore", LogLevel.Warning);
        builder.Logging.AddFilter("Microsoft.Extensions.Hosting.Internal", LogLevel.Warning);

        if (mcpPort.HasValue && !string.IsNullOrEmpty(mcpHost))
        {
            builder.WebHost.UseUrls($"http://{mcpHost}:{mcpPort.Value}");
        }

        // Read CORS origins from env var (comma-separated) or use defaults
        string? corsOriginsEnv = Environment.GetEnvironmentVariable("CORS_ORIGINS");
        string[] allowedOrigins;

        if (!string.IsNullOrWhiteSpace(corsOriginsEnv))
        {
            allowedOrigins = corsOriginsEnv.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        }
        else
        {
            allowedOrigins = ["http://127.0.0.1:*", "https://127.0.0.1:*"];
        }

        builder.Services.AddCors(options =>
        {
            options.AddPolicy("AllowedOrigins", policy =>
            {
                if (allowedOrigins.Length == 0 ||
                    (allowedOrigins.Length == 2 &&
                     allowedOrigins[0] == "http://127.0.0.1:*" &&
                     allowedOrigins[1] == "https://127.0.0.1:*"))
                {
                    // Default: allow any origin on localhost.
                    // WithOrigins() does not support wildcard ports, so we use
                    // SetIsOriginAllowed instead. The token enforces auth;
                    // CORS is defence-in-depth only for a loopback-only server.
                    policy.SetIsOriginAllowed(_ => true)
                          .AllowAnyHeader()
                          .AllowAnyMethod()
                          .AllowCredentials();
                }
                else
                {
                    policy.WithOrigins(allowedOrigins)
                          .AllowAnyHeader()
                          .AllowAnyMethod();
                }
            });
        });

        // SignalR for real-time command delivery
        builder.Services.AddSignalR();

        var app = builder.Build();
        app.UseCors("AllowedOrigins");

        // ============================================================
        // AUTH MIDDLEWARE
        // Validates "Authorization: Bearer <token>" on every request.
        // SignalR WebSocket connections carry the token as ?access_token=
        // because browser WebSocket APIs cannot set custom headers.
        // Exempt: /manifest.xml (fetched by Office host before add-in loads)
        // ============================================================
        app.Use(async (context, next) =>
        {
            var path = context.Request.Path.Value ?? "";

            // manifest.xml must be reachable before the add-in has a token
            if (path.Equals("/manifest.xml", StringComparison.OrdinalIgnoreCase))
            {
                await next();
                return;
            }

            string? provided = null;

            // Standard header: Authorization: Bearer <token>
            var authHeader = context.Request.Headers["Authorization"].FirstOrDefault();
            if (authHeader != null && authHeader.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
                provided = authHeader["Bearer ".Length..].Trim();

            // SignalR WebSocket fallback: ?access_token=<token>
            if (provided == null)
                provided = context.Request.Query["access_token"].FirstOrDefault();

            if (provided == null || provided != BuildToken.Value)
            {
                context.Response.StatusCode = 401;
                context.Response.Headers["WWW-Authenticate"] = "Bearer";
                await context.Response.WriteAsync("Unauthorized");
                return;
            }

            await next();
        });

        // Write token to disk on first run so external MCP clients can read it.
        // Runs on the actual host machine, not the build machine.
        WriteTokenFile();

        // --- Static file serving (PowerPoint/Word/Excel/Outlook add-in UI) ---
        string? staticFilesPath = Environment.GetEnvironmentVariable("STATIC_FILES_PATH")
            ?? Environment.GetEnvironmentVariable("STATIC_FILES_DIR");

        if (!string.IsNullOrEmpty(staticFilesPath))
        {
            app.UseStaticFiles(new StaticFileOptions
            {
                FileProvider = new PhysicalFileProvider(staticFilesPath),
                RequestPath = ""
            });
        }
        else
        {
            // Default: look for wwwroot in the current directory
            string wwwroot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
            if (Directory.Exists(wwwroot))
            {
                app.UseStaticFiles(new StaticFileOptions
                {
                    FileProvider = new PhysicalFileProvider(wwwroot),
                    RequestPath = ""
                });
            }
        }

        // --- Resolve shared state ---
        var registry = McpToolEngine.GetRegistry();
        var commandStore = McpToolEngine.GetCommandStore();

        // --- Start cleanup timer (runs every 30s) ---
        _ = Task.Run(async () =>
        {
            while (true)
            {
                await Task.Delay(30000);
                registry.CleanupTimedOut();
            }
        });

        // ============================================================
        // DYNAMIC MANIFEST XML
        // Serves manifest.xml with URLs inferred from request Host header
        // ============================================================

        app.MapGet("/manifest.xml", (HttpContext context) =>
        {
            var manifest = ManifestTemplate.Render(context.Request);
            return Results.Text(manifest, ManifestTemplate.ContentType, ManifestTemplate.Encoding);
        });

        // ============================================================
        // INSTANCE MANAGEMENT ENDPOINTS
        // ============================================================

        app.MapPost("/instances/register", async (HttpContext context) =>
        {
            using var reader = new StreamReader(context.Request.Body);
            var json = await reader.ReadToEndAsync();
            var data = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(json);
            string appName = data.GetProperty("appName").GetString() ?? "Unknown";
            string documentName = data.TryGetProperty("documentName", out var dn) ? dn.GetString() ?? "" : "";
            string? proposedId = data.TryGetProperty("instanceId", out var pid) ? pid.GetString() : null;

            string instanceId = registry.RegisterInstance(appName, documentName, proposedId);
            Console.WriteLine($"Registered instance: {instanceId} ({appName} - {documentName})");
            return Results.Json(new { instanceId, appName, documentName });
        });

        app.MapPost("/instances/{instanceId}/heartbeat", async (HttpContext context, string instanceId) =>
        {
            using var reader = new StreamReader(context.Request.Body);
            var json = await reader.ReadToEndAsync();
            var data = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(json);
            string? appName = data.TryGetProperty("appName", out var a) ? a.GetString() : null;
            string? documentName = data.TryGetProperty("documentName", out var d) ? d.GetString() : null;
            registry.UpdateHeartbeat(instanceId, appName, documentName);
            return Results.Json(new { status = "ok" });
        });

        app.MapGet("/instances", () =>
        {
            var instances = registry.GetActiveInstances().Select(i => new
            {
                i.InstanceId, i.AppName, i.DocumentName, i.IsAlive,
                RegisteredAt = i.RegisteredAt.ToString("o"),
            });
            return Results.Json(new { instances });
        });

        app.MapGet("/instances/{instanceId}/commands", (string instanceId) =>
        {
            var commands = commandStore.GetAndClaimPendingCommands(instanceId);
            return Results.Json(new { commands });
        });

        // Confirmation endpoint removed — mutation tools now apply directly.
        // Users create backups before enabling LLM write access.

        app.MapPost("/instances/{instanceId}/result", async (HttpContext context, string instanceId) =>
        {
            using var reader = new StreamReader(context.Request.Body);
            var json = await reader.ReadToEndAsync();
            var data = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(json);
            string commandId = data.GetProperty("commandId").GetString() ?? "";
            bool success = data.TryGetProperty("success", out var s) && s.GetBoolean();
            string? error = data.TryGetProperty("error", out var e) ? e.GetString() : null;
            object? payload = data.TryGetProperty("payload", out var p) ? JsonSerializer.Deserialize<object>(p) : null;
            Console.WriteLine($"Instance {instanceId} completed command {commandId}: success={success}");
            commandStore.CompleteCommand(commandId, success, error ?? "", payload);
            return Results.Json(new { status = "ok" });
        });

        // ============================================================
        // MCP PROTOCOL ENDPOINTS (Streamable HTTP)
        // ============================================================

        app.MapPost("/mcp", async (HttpContext context) =>
        {
            using var reader = new StreamReader(context.Request.Body);
            var json = await reader.ReadToEndAsync();
            if (string.IsNullOrWhiteSpace(json))
                return Results.BadRequest("Empty request body");

            // SSE session: if ?sessionId= present, push response into SSE stream
            string? sessionId = context.Request.Query["sessionId"].FirstOrDefault();
            Channel<string>? sseChannel = null;
            if (sessionId != null)
                _sseSessions.TryGetValue(sessionId, out sseChannel);

            async Task<IResult> SendResponse(object responseObj)
            {
                if (sseChannel != null)
                {
                    var serialized = JsonSerializer.Serialize(responseObj);
                    await sseChannel.Writer.WriteAsync(serialized);
                    return Results.Accepted();
                }
                return Results.Json(responseObj, new JsonSerializerOptions { WriteIndented = true });
            }

            try
            {
                var message = JsonSerializer.Deserialize<JsonElement>(json);
                string method = message.GetProperty("method").GetString() ?? "";

                switch (method)
                {
                    case "initialize":
                        return await SendResponse(new
                        {
                            jsonrpc = "2.0", id = message.GetProperty("id"),
                            result = new
                            {
                                protocolVersion = "2024-11-05",
                                serverInfo = new { name = "OfficeMcpServer", version = "0.1.0-spike" },
                                capabilities = new { tools = new { listChanged = true } }
                            }
                        });

                    case "notifications/initialized":
                        // Notifications have no id and expect no response
                        return Results.Accepted();

                    case "tools/list":
                        var tools = McpToolEngine.GetToolDefinitions();
                        return await SendResponse(new
                        { jsonrpc = "2.0", id = message.GetProperty("id"), result = new { tools } });

                    case "tools/call":
                        var toolName = message.GetProperty("params").GetProperty("name").GetString() ?? "";
                        JsonElement? toolArgs = null;
                        if (message.GetProperty("params").TryGetProperty("arguments", out var arguments))
                            toolArgs = arguments;
                        var result = await McpToolEngine.ExecuteTool(toolName, toolArgs);
                        return await SendResponse(new
                        { jsonrpc = "2.0", id = message.GetProperty("id"), result });

                    default:
                        return Results.Problem(detail: $"Method not found: {method}", statusCode: StatusCodes.Status404NotFound);
                }
            }
            catch (Exception ex)
            {
                return Results.Problem(detail: $"Parse error: {ex.Message}", statusCode: StatusCodes.Status400BadRequest);
            }
        });

        // ============================================================
        // HEALTH ENDPOINT
        // ============================================================

        app.MapGet("/health", () =>
        {
            var activeCount = registry.GetActiveInstances().Count;
            return Results.Json(new { status = "ok", activeInstances = activeCount });
        });

        // ============================================================
        // SSE TRANSPORT  (legacy MCP clients: Claude Desktop ≤ 0.7)
        // GET /sse  — opens a text/event-stream, emits one "endpoint" event,
        // then forwards all JSON-RPC responses for this session.
        // POST /mcp?sessionId=<id> — same handler as below, writes response
        // into the SSE channel instead of the HTTP response.
        // ============================================================

        app.MapGet("/sse", async (HttpContext context) =>
        {
            var sessionId = Guid.NewGuid().ToString("N");
            var channel = Channel.CreateUnbounded<string>(new UnboundedChannelOptions { SingleReader = true });
            _sseSessions[sessionId] = channel;

            context.Response.Headers["Content-Type"] = "text/event-stream";
            context.Response.Headers["Cache-Control"] = "no-cache";
            context.Response.Headers["X-Accel-Buffering"] = "no";

            // Send the endpoint event so the client knows where to POST
            var postPath = $"/mcp?sessionId={sessionId}";
            await context.Response.WriteAsync($"event: endpoint\ndata: {postPath}\n\n");
            await context.Response.Body.FlushAsync();

            var ct = context.RequestAborted;
            try
            {
                await foreach (var msg in channel.Reader.ReadAllAsync(ct))
                {
                    await context.Response.WriteAsync($"event: message\ndata: {msg}\n\n", ct);
                    await context.Response.Body.FlushAsync(ct);
                }
            }
            catch (OperationCanceledException) { }
            finally
            {
                _sseSessions.TryRemove(sessionId, out _);
            }
        });

        // ============================================================
        // SIGNALR HUB
        // ============================================================

        app.MapHub<CommandHub>("/hubs/commands");

        // Wire SignalR hub context into McpToolEngine for command push
        var hubContext = app.Services.GetRequiredService<IHubContext<CommandHub>>();
        McpToolEngine.SetHubContext(hubContext);

        // ============================================================
        // OPENAPI REST BRIDGE
        // ============================================================

        app.MapGet("/openapi.json", (HttpRequest request) =>
        {
            var tools = McpToolEngine.GetToolDefinitions();
            var baseUrl = $"{request.Scheme}://{request.Host}";

            var paths = new Dictionary<string, object>();

            foreach (var toolObj in tools)
            {
                var tool = (dynamic)toolObj;
                string toolName = tool.name;
                string description = tool.description;

                var propDict = new Dictionary<string, object>();
                var requiredList = new List<string>();

                // inputSchema is an anonymous type with .properties and .required fields
                try
                {
                    var schemaProps = (Dictionary<string, object>)tool.inputSchema.properties;
                    foreach (var kvp in schemaProps)
                    {
                        var propDef = (dynamic)kvp.Value;
                        var openApiProp = new Dictionary<string, object>
                        {
                            ["type"] = (string?)propDef.type ?? "string",
                            ["description"] = (string?)propDef.description ?? ""
                        };
                        propDict[kvp.Key] = openApiProp;
                    }
                }
                catch (InvalidCastException ex)
                {
                    Console.WriteLine($"Warning: Could not extract properties for tool {toolName}: {ex.Message}");
                }

                try
                {
                    var reqArr = (string[])tool.inputSchema.required;
                    requiredList = reqArr.ToList();
                }
                catch (InvalidCastException ex)
                {
                    Console.WriteLine($"Warning: Could not extract required for tool {toolName}: {ex.Message}");
                }

                var requestSchema = new
                {
                    type = "object",
                    properties = propDict,
                    required = requiredList
                };

                paths[$"/api/{toolName}"] = new
                {
                    post = new
                    {
                        summary = description,
                        operationId = toolName,
                        tags = new[] { "Office Tools" },
                        requestBody = new
                        {
                            required = true,
                            content = new Dictionary<string, object>
                            {
                                ["application/json"] = new { schema = requestSchema }
                            }
                        },
                        responses = new Dictionary<string, object>
                        {
                            ["200"] = new
                            {
                                description = "Tool execution result",
                                content = new Dictionary<string, object>
                                {
                                    ["application/json"] = new
                                    {
                                        schema = new
                                        {
                                            type = "object",
                                            properties = new Dictionary<string, object>
                                            {
                                                ["content"] = new { type = "array", description = "Array of content blocks" },
                                                ["isError"] = new { type = "boolean", description = "Whether the tool call resulted in an error" }
                                            }
                                        }
                                    }
                                }
                            },
                            ["400"] = new { description = "Invalid request" },
                            ["404"] = new { description = "Tool not found" }
                        }
                    }
                };
            }

            var spec = new
            {
                openapi = "3.0.0",
                info = new
                {
                    title = "Office LLM Harness API",
                    version = "0.1.0-spike",
                    description = "REST API bridge for Office MCP tools. Each MCP tool is exposed as a POST endpoint."
                },
                servers = new[] { new { url = baseUrl, description = "Local MCP server" } },
                paths,
                components = new
                {
                    schemas = new Dictionary<string, object>
                    {
                        ["ToolResult"] = new
                        {
                            type = "object",
                            properties = new Dictionary<string, object>
                            {
                                ["content"] = new { type = "array", items = new { type = "object" } },
                                ["isError"] = new { type = "boolean" }
                            }
                        }
                    }
                }
            };

            return Results.Json(spec, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
        });

        app.MapPost("/api/{toolName}", async (HttpContext context, string toolName) =>
        {
            using var reader = new StreamReader(context.Request.Body);
            var json = await reader.ReadToEndAsync();

            System.Text.Json.JsonElement? args = null;
            if (!string.IsNullOrWhiteSpace(json))
            {
                try { args = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(json); }
                catch { return Results.BadRequest(new { error = "Invalid JSON body" }); }
            }

            var result = await McpToolEngine.ExecuteTool(toolName, args);
            return Results.Json(result, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
        });

        app.MapGet("/docs", () => Results.Text(
            "<!DOCTYPE html><html><head><title>Office LLM Harness API</title>" +
            "<meta charset=\"utf-8\"/><meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">" +
            "<link rel=\"stylesheet\" type=\"text/css\" href=\"https://unpkg.com/swagger-ui-dist@5/swagger-ui.css\">" +
            "</head><body><div id=\"swagger-ui\"></div>" +
            "<script src=\"https://unpkg.com/swagger-ui-dist@5/swagger-ui-bundle.js\"></script>" +
            "<script>SwaggerUIBundle({url:'/openapi.json',dom_id:'#swagger-ui'," +
            "presets:[SwaggerUIBundle.presets.apis,SwaggerUIBundle.SwaggerUIStandalonePreset]," +
            "layout:'BaseLayout'})</script></body></html>",
            "text/html", System.Text.Encoding.UTF8));

        return app;
    }

    /// <summary>
    /// Persists the compiled-in token to a user-readable file so external MCP
    /// clients (e.g. Claude Desktop) can read it once on the actual host machine.
    ///
    /// Cross-platform paths:
    ///   Windows : %APPDATA%\OfficeMcpServer\token
    ///   macOS   : ~/.config/OfficeMcpServer/token
    ///   Linux   : ~/.config/OfficeMcpServer/token
    /// </summary>
    /// <summary>
    /// Returns the cross-platform config directory:
    ///   Windows : %APPDATA%\OfficeMcpServer
    ///   macOS   : ~/Library/Application Support/OfficeMcpServer  (SpecialFolder.ApplicationData)
    ///   Linux   : $XDG_CONFIG_HOME/OfficeMcpServer or ~/.config/OfficeMcpServer
    /// Note: SpecialFolder.ApplicationData on macOS returns ~/Library/Application Support,
    /// NOT ~/.config. This is the correct macOS convention.
    /// </summary>
    internal static string GetConfigDirectory()
        => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OfficeMcpServer");

    private static void WriteTokenFile()
    {
        try
        {
            var dir = GetConfigDirectory();
            Directory.CreateDirectory(dir);
            var tokenPath = Path.Combine(dir, "token");
            File.WriteAllText(tokenPath, BuildToken.Value);

            // chmod 600 on Unix — owner read/write only
            if (!OperatingSystem.IsWindows())
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "chmod",
                    Arguments = $"600 \"{tokenPath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                })?.WaitForExit();
            }

            Console.WriteLine($"Token file: {tokenPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: could not write token file: {ex.Message}");
        }
    }
}
