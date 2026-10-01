namespace OfficeMcpServer.Models;

/// <summary>
/// Stable error codes for consistent error responses across all tools.
/// The LLM can learn these codes to build reliable recovery strategies.
/// </summary>
public static class ErrorCodes
{
    public const string INSTANCE_NOT_FOUND = "INSTANCE_NOT_FOUND";
    public const string MISSING_PARAMETER = "MISSING_PARAMETER";
    public const string INVALID_PARAMETER = "INVALID_PARAMETER";
    public const string TIMEOUT = "TIMEOUT";
    public const string HOST_NOT_AVAILABLE = "HOST_NOT_AVAILABLE";
    public const string PERMISSION_DENIED = "PERMISSION_DENIED";
    public const string RANGE_TOO_LARGE = "RANGE_TOO_LARGE";
    public const string INVALID_FORMULA = "INVALID_FORMULA";
    public const string CONFIRMATION_REQUIRED = "CONFIRMATION_REQUIRED";
    public const string INTERNAL_ERROR = "INTERNAL_ERROR";
}

/// <summary>
/// Structured error response with stable error code and optional details.
/// </summary>
public class ToolError
{
    public string Error { get; }
    public string ErrorCode { get; }
    public object? Details { get; }

    public ToolError(string error, string errorCode, object? details = null)
    {
        Error = error;
        ErrorCode = errorCode;
        Details = details;
    }

    /// <summary>
    /// Builds the standard MCP error response envelope.
    /// </summary>
    public object ToMcpResponse()
    {
        var errorObj = new Dictionary<string, object>
        {
            ["error"] = Error,
            ["errorCode"] = ErrorCode
        };
        if (Details != null)
            errorObj["details"] = Details;

        return new
        {
            content = new[]
            {
                new
                {
                    type = "text",
                    text = System.Text.Json.JsonSerializer.Serialize(errorObj, new System.Text.Json.JsonSerializerOptions { WriteIndented = true })
                }
            },
            isError = true
        };
    }
}
