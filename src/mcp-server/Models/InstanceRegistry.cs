namespace OfficeMcpServer.Models;

/// <summary>
/// Represents a registered Office application instance (e.g., one PowerPoint session).
/// Each add-in registers itself with a unique instance ID.
/// </summary>
public class OfficeInstance
{
    public string InstanceId { get; set; } = string.Empty;
    public string AppName { get; set; } = string.Empty;
    public string DocumentName { get; set; } = string.Empty;
    public DateTime RegisteredAt { get; set; } = DateTime.UtcNow;
    public DateTime LastHeartbeat { get; set; } = DateTime.UtcNow;
    public bool IsAlive { get; set; } = true;

    /// <summary>
    /// Updates the heartbeat timestamp.
    /// </summary>
    public void RefreshHeartbeat()
    {
        LastHeartbeat = DateTime.UtcNow;
        IsAlive = true;
    }

    /// <summary>
    /// Checks if the instance has timed out (no heartbeat for 60 seconds).
    /// </summary>
    public bool HasTimedOut(int timeoutSeconds = 60)
    {
        return (DateTime.UtcNow - LastHeartbeat).TotalSeconds > timeoutSeconds;
    }
}

/// <summary>
/// Thread-safe registry of active Office application instances.
/// The MCP server uses this to route tool calls to the correct add-in.
/// </summary>
public class InstanceRegistry
{
    private readonly Dictionary<string, OfficeInstance> _instances = new();
    private readonly object _lock = new();

    private static readonly System.Text.RegularExpressions.Regex _safeId =
        new(@"^[a-z][a-z0-9_]{2,63}$", System.Text.RegularExpressions.RegexOptions.Compiled);

    /// <summary>
    /// Registers a new Office instance and returns its ID.
    /// If the add-in proposes an ID (Option B stable IDs), it is used directly
    /// after validation. Falls back to a random UUID-based ID if omitted or invalid.
    /// Re-registration with the same ID refreshes the heartbeat in place.
    /// </summary>
    public string RegisterInstance(string appName, string documentName, string? proposedId = null)
    {
        lock (_lock)
        {
            string instanceId = (!string.IsNullOrWhiteSpace(proposedId) && _safeId.IsMatch(proposedId))
                ? proposedId
                : $"office_{Guid.NewGuid():N}".Substring(0, 16);

            if (_instances.TryGetValue(instanceId, out var existing))
            {
                existing.AppName = appName;
                existing.DocumentName = documentName;
                existing.RefreshHeartbeat();
                return instanceId;
            }

            _instances[instanceId] = new OfficeInstance
            {
                InstanceId = instanceId,
                AppName = appName,
                DocumentName = documentName,
            };

            return instanceId;
        }
    }

    /// <summary>
    /// Updates an existing instance's heartbeat and info.
    /// </summary>
    public void UpdateHeartbeat(string instanceId, string? appName = null, string? documentName = null)
    {
        lock (_lock)
        {
            if (_instances.TryGetValue(instanceId, out var instance))
            {
                instance.RefreshHeartbeat();
                if (appName != null) instance.AppName = appName;
                if (documentName != null) instance.DocumentName = documentName;
            }
        }
    }

    /// <summary>
    /// Gets all active (non-timed-out) instances.
    /// </summary>
    public List<OfficeInstance> GetActiveInstances()
    {
        lock (_lock)
        {
            return _instances.Values.Where(i => !i.HasTimedOut()).ToList();
        }
    }

    /// <summary>
    /// Gets a specific instance by ID, or null if not found.
    /// </summary>
    public OfficeInstance? GetInstance(string instanceId)
    {
        lock (_lock)
        {
            return _instances.TryGetValue(instanceId, out var instance) && !instance.HasTimedOut()
                ? instance
                : null;
        }
    }

    /// <summary>
    /// Removes timed-out instances (no heartbeat for 60 seconds).
    /// </summary>
    public void CleanupTimedOut()
    {
        lock (_lock)
        {
            var timedOut = _instances.Values.Where(i => i.HasTimedOut(60)).ToList();
            foreach (var instance in timedOut)
            {
                _instances.Remove(instance.InstanceId);
                Console.WriteLine($"Instance {instance.InstanceId} timed out and removed");
            }
        }
    }
}
