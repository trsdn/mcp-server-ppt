using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;
using PptMcp.Core.Commands.File;

namespace PptMcp.McpServer.Tools;

/// <summary>
/// Actions for the file tool (hand-coded because session management is not generated).
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<PptFileAction>))]
public enum PptFileAction
{
    [JsonStringEnumMemberName("open")] Open,
    [JsonStringEnumMemberName("close")] Close,
    [JsonStringEnumMemberName("create")] Create,
    [JsonStringEnumMemberName("list")] List,
    [JsonStringEnumMemberName("test")] Test,
    [JsonStringEnumMemberName("save")] Save
}

/// <summary>
/// PowerPoint file and session management tool for MCP server.
/// </summary>
[McpServerToolType]
public static class PptFileTool
{
    /// <summary>
    /// File and session management for PowerPoint automation.
    ///
    /// WORKFLOW: open → use session_id with other tools → close (save=true to persist changes).
    /// NEW FILES: Use 'create' action to create file AND start session in one call.
    ///
    /// SESSION REUSE: Call 'list' first to check for existing sessions.
    /// If file is already open, reuse existing session_id instead of opening again.
    /// </summary>
    [McpServerTool(Name = "file", Title = "File Operations", Destructive = true)]
    [Description("File and session management for PowerPoint automation. WORKFLOW: open → use session_id with other tools → close (save=true to persist changes).")]
    public static string PptFile(
        PptFileAction action,
        [DefaultValue(null)] string? path,
        [DefaultValue(null)] string? session_id,
        [DefaultValue(false)] bool save,
        [DefaultValue(false)] bool show,
        [DefaultValue(300)] int timeout_seconds)
    {
        return PptToolsBase.ExecuteToolAction("file", action.ToString().ToLowerInvariant(), path, () =>
        {
            return action switch
            {
                PptFileAction.List => ListSessions(),
                PptFileAction.Open => OpenSession(path!, show, timeout_seconds),
                PptFileAction.Close => CloseSession(session_id!, save),
                PptFileAction.Create => CreateSession(path!, show, timeout_seconds),
                PptFileAction.Test => TestFile(path!),
                PptFileAction.Save => SaveSession(session_id!),
                _ => throw new ArgumentException($"Unknown action: {action}")
            };
        });
    }

    private static string OpenSession(string path, bool show, int timeoutSeconds)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("path is required for 'open' action");

        var pathError = PptToolsBase.ValidateWindowsPath(path);
        if (pathError != null) return pathError;

        if (!File.Exists(path))
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"File not found: {path}",
                filePath = path,
                isError = true
            }, PptToolsBase.JsonOptions);
        }

        var response = ServiceBridge.ServiceBridge.SendAsync(
            "session.open", null,
            new { filePath = path, show, timeoutSeconds },
            timeoutSeconds
        ).GetAwaiter().GetResult();

        if (!response.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = response.ErrorMessage ?? "Failed to open session",
                filePath = path,
                isError = true
            }, PptToolsBase.JsonOptions);
        }

        return TransformSessionResponse(response.Result, path);
    }

    private static string CloseSession(string sessionId, bool save)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
            throw new ArgumentException("session_id is required for 'close' action");

        var response = ServiceBridge.ServiceBridge.SendAsync(
            "session.close", sessionId, new { save }
        ).GetAwaiter().GetResult();

        if (!response.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                session_id = sessionId,
                errorMessage = response.ErrorMessage ?? "Failed to close session",
                isError = true
            }, PptToolsBase.JsonOptions);
        }

        return response.Result ?? JsonSerializer.Serialize(new
        {
            success = true,
            session_id = sessionId,
            saved = save
        }, PptToolsBase.JsonOptions);
    }

    private static string CreateSession(string path, bool show, int timeoutSeconds)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("path is required for 'create' action");

        var pathError = PptToolsBase.ValidateWindowsPath(path);
        if (pathError != null) return pathError;

        var response = ServiceBridge.ServiceBridge.SendAsync(
            "session.create", null,
            new { filePath = path, show, timeoutSeconds },
            timeoutSeconds
        ).GetAwaiter().GetResult();

        if (!response.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = response.ErrorMessage ?? "Failed to create session",
                filePath = path,
                isError = true
            }, PptToolsBase.JsonOptions);
        }

        return TransformSessionResponse(response.Result, path);
    }

    private static string SaveSession(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
            throw new ArgumentException("session_id is required for 'save' action");

        var response = ServiceBridge.ServiceBridge.SendAsync(
            "session.save", sessionId
        ).GetAwaiter().GetResult();

        if (!response.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                session_id = sessionId,
                errorMessage = response.ErrorMessage ?? "Failed to save",
                isError = true
            }, PptToolsBase.JsonOptions);
        }

        return JsonSerializer.Serialize(new
        {
            success = true,
            session_id = sessionId
        }, PptToolsBase.JsonOptions);
    }

    private static string ListSessions()
    {
        var response = ServiceBridge.ServiceBridge.SendAsync("session.list").GetAwaiter().GetResult();

        if (!response.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = response.ErrorMessage ?? "Failed to list sessions",
                isError = true
            }, PptToolsBase.JsonOptions);
        }

        return response.Result ?? JsonSerializer.Serialize(new
        {
            success = true,
            sessions = Array.Empty<object>(),
            count = 0
        }, PptToolsBase.JsonOptions);
    }

    private static string TestFile(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("path is required for 'test' action");

        var pathError = PptToolsBase.ValidateWindowsPath(path);
        if (pathError != null) return pathError;

        var fileCommands = new FileCommands();
        var info = fileCommands.Test(path);

        return JsonSerializer.Serialize(new
        {
            success = info.Success,
            exists = info.Exists,
            filePath = info.FilePath,
            fileName = info.FileName,
            fileSizeBytes = info.FileSizeBytes,
            isReadOnly = info.IsReadOnly,
            isMacroEnabled = info.IsMacroEnabled
        }, PptToolsBase.JsonOptions);
    }

    /// <summary>
    /// Transforms the service response to use snake_case session_id for MCP compatibility.
    /// </summary>
    private static string TransformSessionResponse(string? result, string path)
    {
        if (!string.IsNullOrEmpty(result))
        {
            try
            {
                using var doc = JsonDocument.Parse(result);
                if (doc.RootElement.TryGetProperty("sessionId", out var sessionIdProp))
                {
                    var sessionId = sessionIdProp.GetString();
                    string? filePath = doc.RootElement.TryGetProperty("filePath", out var fp) ? fp.GetString() : path;
                    return JsonSerializer.Serialize(new
                    {
                        success = true,
                        session_id = sessionId,
                        filePath
                    }, PptToolsBase.JsonOptions);
                }
            }
            catch (JsonException) { }
            return result;
        }

        return JsonSerializer.Serialize(new { success = true, filePath = path }, PptToolsBase.JsonOptions);
    }
}

