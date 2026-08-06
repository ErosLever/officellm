namespace OfficeMcpServer.Models;

/// <summary>
/// A pending command dispatched to an Office add-in instance.
/// </summary>
public class PendingCommand
{
    public string Id { get; set; } = string.Empty;
    public string InstanceId { get; set; } = string.Empty;
    public string Command { get; set; } = string.Empty;
    public object? Args { get; set; }
    public DateTime CreatedAt { get; set; }
    public string? ClaimedBy { get; set; }
    public DateTime? ClaimedAt { get; set; }
    public bool Completed { get; set; }
    public bool Success { get; set; }
    public string? Error { get; set; }
    public object? Payload { get; set; }
    public DateTime? CompletedAt { get; set; }

    /// <summary>
    /// TaskCompletionSource for real-time result notification (SignalR).
    /// Resolved immediately when CompleteCommand is called, replacing the old polling loop.
    /// </summary>
    public TaskCompletionSource<PendingCommand>? CompletionSource { get; set; }
}

/// <summary>
/// Thread-safe store for commands dispatched to add-in instances.
/// Supports queuing, claiming, result waiting (via TaskCompletionSource), and cleanup.
/// </summary>
public class CommandStore
{
    private readonly Dictionary<string, PendingCommand> _commands = new();
    private readonly object _lock = new();

    /// <summary>
    /// Adds a new pending command with a TaskCompletionSource for instant result notification.
    /// </summary>
    public void AddCommand(PendingCommand command)
    {
        lock (_lock)
        {
            command.CompletionSource = new TaskCompletionSource<PendingCommand>();
            _commands[command.Id] = command;
        }
    }

    /// <summary>
    /// Atomically gets and claims pending commands for a specific instance.
    /// Used by HTTP polling fallback. SignalR pushes skip this.
    /// </summary>
    public List<PendingCommand> GetAndClaimPendingCommands(string instanceId)
    {
        lock (_lock)
        {
            var staleThreshold = DateTime.UtcNow.AddSeconds(-10);
            var pending = _commands.Values
                .Where(c => c.InstanceId == instanceId && !c.Completed &&
                    (string.IsNullOrEmpty(c.ClaimedBy) ||
                     (c.ClaimedAt.HasValue && c.ClaimedAt.Value < staleThreshold)))
                .ToList();

            foreach (var cmd in pending)
            {
                cmd.ClaimedBy = instanceId;
                cmd.ClaimedAt = DateTime.UtcNow;
            }

            return pending;
        }
    }

    /// <summary>
    /// Marks a command as claimed by its instanceId (called after a successful SignalR push).
    /// Prevents HTTP polling from picking up the same command within the stale window.
    /// </summary>
    public void MarkClaimed(string commandId)
    {
        lock (_lock)
        {
            if (_commands.TryGetValue(commandId, out var cmd))
            {
                cmd.ClaimedBy = cmd.InstanceId;
                cmd.ClaimedAt = DateTime.UtcNow;
            }
        }
    }

    /// <summary>
    /// Gets a single pending command by ID for SignalR push.
    /// Returns null if not found or already completed.
    /// </summary>
    public PendingCommand? GetPendingCommand(string instanceId, string commandId)
    {
        lock (_lock)
        {
            if (_commands.TryGetValue(commandId, out var cmd)
                && cmd.InstanceId == instanceId
                && !cmd.Completed)
            {
                cmd.ClaimedBy = instanceId;
                return cmd;
            }
            return null;
        }
    }

    /// <summary>
    /// Records the result of a completed command.
    /// Resolves the TaskCompletionSource for instant notification (SignalR path).
    /// </summary>
    public void CompleteCommand(string commandId, bool success, string error, object? payload)
    {
        TaskCompletionSource<PendingCommand>? tcs = null;
        PendingCommand? cmd = null;

        lock (_lock)
        {
            if (_commands.TryGetValue(commandId, out cmd))
            {
                cmd.Completed = true;
                cmd.Success = success;
                cmd.Error = error;
                cmd.Payload = payload;
                cmd.CompletedAt = DateTime.UtcNow;
                tcs = cmd.CompletionSource;
            }
        }

        // Resolve outside the lock to avoid deadlocks
        tcs?.TrySetResult(cmd!);
    }

    /// <summary>
    /// Waits for a command result with a timeout.
    /// Uses TaskCompletionSource for instant notification (SignalR path).
    /// Falls back to polling if TCS is not available (legacy compatibility).
    /// </summary>
    public async Task<PendingCommand?> WaitForResult(string commandId, int timeoutSeconds = 30)
    {
        TaskCompletionSource<PendingCommand>? tcs;
        lock (_lock)
        {
            _commands.TryGetValue(commandId, out var cmd);
            tcs = cmd?.CompletionSource;
        }

        if (tcs != null)
        {
            // SignalR path: wait for instant notification
            using var cts = new CancellationTokenSource(timeoutSeconds * 1000);
            cts.Token.Register(() => tcs.TrySetCanceled());

            try
            {
                return await tcs.Task;
            }
            catch (OperationCanceledException)
            {
                // Timeout — return null
                return null;
            }
        }

        // Legacy fallback: polling (shouldn't normally be hit)
        var deadline = DateTime.UtcNow.AddSeconds(timeoutSeconds);
        while (DateTime.UtcNow < deadline)
        {
            lock (_lock)
            {
                if (_commands.TryGetValue(commandId, out var cmd) && cmd.Completed)
                    return cmd;
            }
            await Task.Delay(500);
        }

        return null;
    }

    /// <summary>
    /// Cleans up old completed commands (older than 5 minutes).
    /// </summary>
    public void CleanupOldCommands()
    {
        lock (_lock)
        {
            var cutoff = DateTime.UtcNow.AddMinutes(-5);
            var toRemove = _commands.Values.Where(c => c.Completed && (c.CompletedAt ?? DateTime.MinValue) < cutoff).ToList();
            foreach (var cmd in toRemove)
            {
                cmd.CompletionSource?.TrySetCanceled();
                _commands.Remove(cmd.Id);
            }
        }
    }
}
