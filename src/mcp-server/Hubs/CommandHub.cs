using OfficeMcpServer.Models;
using OfficeMcpServer.Tools;
using Microsoft.AspNetCore.SignalR;
namespace OfficeMcpServer.Hubs;

/// <summary>
/// SignalR hub for real-time command dispatch to Office add-ins.
/// Replaces HTTP polling with instant push via WebSocket.
/// </summary>
public class CommandHub : Hub
{
    /// <summary>
    /// Called by the add-in after connecting. Joins the instance's SignalR group
    /// so the server can push commands to it.
    /// </summary>
    public async Task JoinGroup(string instanceId)
    {
        await Groups.AddToGroupAsync(Context.ConnectionId, instanceId);
        Console.WriteLine($"SignalR: Connection {Context.ConnectionId} joined group {instanceId}");
    }

    /// <summary>
    /// Called by the add-in when a command completes. Resolves the TaskCompletionSource
    /// in CommandStore for instant result delivery to the waiting MCP caller.
    /// </summary>
    public Task ReportResult(string commandId, bool success, string? error, object? payload)
    {
        var commandStore = McpToolEngine.GetCommandStore();
        var cmd = commandStore.TryGetCommand(commandId);
        var toolName = cmd?.Command ?? "unknown";
        var docName = cmd != null ? McpToolEngine.GetRegistry().GetInstance(cmd.InstanceId)?.DocumentName : null;
        var docLabel = string.IsNullOrEmpty(docName) ? (cmd?.InstanceId ?? "unknown") : docName;
        var outcome = success ? "succeeded" : $"failed ({error})";
        Console.WriteLine($"SignalR: '{toolName}' on '{docLabel}' {outcome}");
        commandStore.CompleteCommand(commandId, success, error ?? "", payload);
        return Task.CompletedTask;
    }

    /// <summary>
    /// Called when an add-in disconnects. Cleans up the SignalR group.
    /// </summary>
    public override async Task OnDisconnectedAsync(Exception? exception)
    {
        // Note: We don't know which instanceId this connection belonged to here.
        // The instance will time out via heartbeat cleanup if it doesn't reconnect.
        if (exception != null)
        {
            Console.WriteLine($"SignalR: Connection {Context.ConnectionId} disconnected with error: {exception.Message}");
        }
        else
        {
            Console.WriteLine($"SignalR: Connection {Context.ConnectionId} disconnected");
        }

        await base.OnDisconnectedAsync(exception);
    }
}
