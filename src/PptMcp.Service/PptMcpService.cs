using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Text.Json;
using PptMcp.ComInterop.Session;
using PptMcp.Service.Rpc;
using StreamJsonRpc;
using PptMcp.Generated;

namespace PptMcp.Service;

/// <summary>
/// The PptMcp Service. Holds SessionManager and executes Core commands.
/// Runs in-process within the host (MCP Server or CLI), accepting commands via named pipe.
/// The named pipe enables cross-thread communication between the host's request threads
/// and the service's STA thread (required for COM interop).
/// </summary>
public sealed class PptMcpService : IDisposable
{
    private readonly SessionManager _sessionManager = new();
    private readonly CancellationTokenSource _shutdownCts = new();
    private readonly DateTime _startTime = DateTime.UtcNow;
    private string _pipeName = "";
    private TimeSpan? _idleTimeout;
    private DateTime _lastActivityTime = DateTime.UtcNow;
    private bool _disposed;

    // Core command instances, one per generated category.
    // Built from ServiceRegistry.CategoryTable so a new Core category is picked up
    // automatically instead of needing a hand-written field and switch case (GitHub #124).
    private readonly Dictionary<string, object> _commandInstances =
        ServiceRegistry.CategoryTable.ToDictionary(
            static kvp => kvp.Key,
            static kvp => kvp.Value.CreateCommands(),
            StringComparer.Ordinal);

    public PptMcpService()
    {
    }

    public DateTime StartTime => _startTime;
    public int SessionCount => _sessionManager.GetActiveSessions().Count;
    public SessionManager SessionManager => _sessionManager;

    /// <summary>
    /// Runs the service in-process, listening for commands on the named pipe.
    /// This method blocks until shutdown is requested via <see cref="RequestShutdown"/>.
    /// </summary>
    /// <param name="pipeName">The named pipe to listen on.</param>
    /// <param name="idleTimeout">Optional idle timeout. Service shuts down after this duration with no active sessions. Null = no timeout.</param>
    public async Task RunAsync(string pipeName, TimeSpan? idleTimeout = null)
    {
        _pipeName = pipeName;
        _idleTimeout = idleTimeout;
        await RunPipeServerAsync(_shutdownCts.Token);
    }

    public void RequestShutdown() => _shutdownCts.Cancel();

    // Exposed for testing — backoff parameters for pipe server accept loop error recovery
    internal static readonly TimeSpan InitialBackoff = TimeSpan.FromMilliseconds(100);
    internal static readonly TimeSpan MaxBackoff = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Records client activity to keep the idle timeout monitor alive.
    /// Called by <see cref="Rpc.DaemonRpcTarget"/> on each incoming RPC call.
    /// </summary>
    internal void RecordActivity() => _lastActivityTime = DateTime.UtcNow;

    private async Task RunPipeServerAsync(CancellationToken cancellationToken)
    {
        // Use a semaphore to limit concurrent connections (prevents resource exhaustion)
        using var connectionLimit = new SemaphoreSlim(10, 10);

        // Start idle timeout monitor if configured
        if (_idleTimeout.HasValue)
        {
            _ = Task.Run(() => MonitorIdleTimeoutAsync(cancellationToken), cancellationToken);
        }

        var currentBackoff = InitialBackoff;

        while (!cancellationToken.IsCancellationRequested)
        {
            NamedPipeServerStream? server = null;
            try
            {
                server = ServiceSecurity.CreateSecureServer(_pipeName);
                await server.WaitForConnectionAsync(cancellationToken);

                // Success — reset backoff
                currentBackoff = InitialBackoff;

                // Record activity on each connection
                _lastActivityTime = DateTime.UtcNow;

                // Capture server for the task
                var clientServer = server;
                server = null; // Prevent disposal in finally - task owns it now

                // Handle client via StreamJsonRpc — replaces hand-rolled JSON protocol
                // with standard JSON-RPC 2.0 over Content-Length-delimited framing.
                _ = Task.Run(async () =>
                {
                    await connectionLimit.WaitAsync(cancellationToken);
                    try
                    {
                        var rpcTarget = new DaemonRpcTarget(this);
                        using var rpc = JsonRpc.Attach(clientServer, rpcTarget);
                        await rpc.Completion; // Waits until client disconnects
                    }
                    finally
                    {
                        connectionLimit.Release();
                        try { if (clientServer.IsConnected) clientServer.Disconnect(); } catch { }
                        await clientServer.DisposeAsync();
                    }
                }, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception)
            {
                // Backoff to prevent CPU spin when errors repeat (e.g. pipe creation failure).
                // Doubles each iteration: 100ms → 200ms → 400ms → … → 5s cap.
                // Resets to 100ms on next successful connection.
                try { await Task.Delay(currentBackoff, cancellationToken); } catch (OperationCanceledException) { break; }
                currentBackoff = TimeSpan.FromMilliseconds(Math.Min(currentBackoff.TotalMilliseconds * 2, MaxBackoff.TotalMilliseconds));
            }
            finally
            {
                if (server != null)
                {
                    try { if (server.IsConnected) server.Disconnect(); } catch (Exception) { /* Cleanup — disconnect may fail if client already disconnected */ }
                    await server.DisposeAsync();
                }
            }
        }
    }

    private async Task MonitorIdleTimeoutAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            await Task.Delay(TimeSpan.FromSeconds(30), cancellationToken);

            var hasSessions = _sessionManager.GetActiveSessions().Count > 0;
            if (hasSessions)
            {
                _lastActivityTime = DateTime.UtcNow;
                continue;
            }

            var idleTime = DateTime.UtcNow - _lastActivityTime;
            if (idleTime >= _idleTimeout!.Value)
            {
                RequestShutdown();
                break;
            }
        }
    }

    /// <summary>
    /// Processes a service request directly (in-process, no pipe).
    /// Used by the MCP Server for direct in-process communication.
    /// </summary>
    public async Task<ServiceResponse> ProcessAsync(ServiceRequest request)
    {
        try
        {
            // Route command
            var parts = request.Command.Split('.', 2);
            var category = parts[0];
            var action = parts.Length > 1 ? parts[1] : "";

            return category switch
            {
                "service" => HandleServiceCommand(action),
                "session" => HandleSessionCommand(action, request),
                "diag" => HandleDiagCommand(action, request),
                _ => await DispatchGeneratedAsync(category, action, request)
            };
        }
        catch (Exception ex)
        {
            // Include type name so callers can distinguish exception kinds (GitHub #482, Bug 5)
            return new ServiceResponse { Success = false, ErrorMessage = $"{ex.GetType().Name}: {ex.Message}" };
        }
    }

    // === SERVICE COMMANDS ===

    private ServiceResponse HandleServiceCommand(string action)
    {
        return action switch
        {
            "ping" => new ServiceResponse { Success = true },
            "shutdown" => HandleShutdown(),
            "status" => HandleStatus(),
            _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown service action: {action}" }
        };
    }

    private ServiceResponse HandleShutdown()
    {
        _shutdownCts.Cancel();
        return new ServiceResponse { Success = true };
    }

    private ServiceResponse HandleStatus()
    {
        var status = new ServiceStatus
        {
            Running = true,
            ProcessId = Environment.ProcessId,
            SessionCount = _sessionManager.GetActiveSessions().Count,
            StartTime = _startTime
        };
        return new ServiceResponse { Success = true, Result = JsonSerializer.Serialize(status, ServiceProtocol.JsonOptions) };
    }

    // === SESSION COMMANDS ===

    // === DIAG COMMANDS ===

    private static ServiceResponse HandleDiagCommand(string action, ServiceRequest request)
    {
        return action switch
        {
            "ping" => new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new
                {
                    success = true,
                    action = "ping",
                    message = "pong",
                    timestamp = DateTime.UtcNow.ToString("o")
                }, ServiceProtocol.JsonOptions)
            },
            "echo" => HandleDiagEcho(request),
            "validate-params" => HandleDiagValidateParams(request),
            _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown diag action: {action}" }
        };
    }

    private static ServiceResponse HandleDiagEcho(ServiceRequest request)
    {
        Dictionary<string, JsonElement>? args = null;
        if (!string.IsNullOrEmpty(request.Args))
            args = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(request.Args, ServiceProtocol.JsonOptions);

        if (args == null || !args.TryGetValue("message", out var messageEl) || messageEl.ValueKind == JsonValueKind.Null)
        {
            return new ServiceResponse { Success = false, ErrorMessage = "Parameter 'message' is required for echo" };
        }

        var message = messageEl.GetString()!;
        string? tag = null;
        if (args.TryGetValue("tag", out var tagEl) && tagEl.ValueKind != JsonValueKind.Null)
            tag = tagEl.GetString();

        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new
            {
                success = true,
                action = "echo",
                message,
                tag
            }, ServiceProtocol.JsonOptions)
        };
    }

    private static ServiceResponse HandleDiagValidateParams(ServiceRequest request)
    {
        Dictionary<string, JsonElement>? args = null;
        if (!string.IsNullOrEmpty(request.Args))
            args = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(request.Args, ServiceProtocol.JsonOptions);

        if (args == null || !args.TryGetValue("name", out var nameEl) || nameEl.ValueKind == JsonValueKind.Null)
        {
            return new ServiceResponse { Success = false, ErrorMessage = "Parameter 'name' is required for validate-params" };
        }

        var count = args.TryGetValue("count", out var countEl) && countEl.ValueKind == JsonValueKind.Number ? countEl.GetInt32() : 0;
        string? label = args.TryGetValue("label", out var labelEl) && labelEl.ValueKind != JsonValueKind.Null ? labelEl.GetString() : null;
        var verbose = args.TryGetValue("verbose", out var verboseEl) && verboseEl.ValueKind != JsonValueKind.Null && verboseEl.GetBoolean();

        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new
            {
                success = true,
                action = "validate-params",
                parameters = new
                {
                    name = nameEl.GetString(),
                    count,
                    label,
                    verbose
                }
            }, ServiceProtocol.JsonOptions)
        };
    }

    // === SESSION COMMANDS ===

    private ServiceResponse HandleSessionCommand(string action, ServiceRequest request)
    {
        return action switch
        {
            "create" => HandleSessionCreate(request),
            "open" => HandleSessionOpen(request),
            "close" => HandleSessionClose(request),
            "save" => HandleSessionSave(request),
            "list" => HandleSessionList(),
            _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown session action: {action}" }
        };
    }

    private ServiceResponse HandleSessionCreate(ServiceRequest request)
    {
        var args = ServiceRegistry.DeserializeArgs<SessionOpenArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args?.FilePath))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "filePath is required" };
        }

        var fullPath = Path.GetFullPath(args.FilePath);

        if (File.Exists(fullPath))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"File already exists: {fullPath}. Use session open to open an existing presentation."
            };
        }

        var extension = Path.GetExtension(fullPath);
        if (!string.Equals(extension, ".pptx", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(extension, ".pptm", StringComparison.OrdinalIgnoreCase))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Invalid file extension '{extension}'. session create supports .pptx and .pptm only."
            };
        }

        try
        {
            // Use the combined create+open which starts PowerPoint only once
            TimeSpan? timeout = args.TimeoutSeconds.HasValue
                ? TimeSpan.FromSeconds(args.TimeoutSeconds.Value)
                : null;
            var sessionId = _sessionManager.CreateSessionForNewFile(fullPath, show: args.Show, operationTimeout: timeout, origin: SessionOrigin.CLI);

            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, sessionId, filePath = fullPath }, ServiceProtocol.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return new ServiceResponse { Success = false, ErrorMessage = $"{ex.GetType().Name}: {ex.Message}" };
        }
    }

    private ServiceResponse HandleSessionOpen(ServiceRequest request)
    {
        var args = ServiceRegistry.DeserializeArgs<SessionOpenArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args?.FilePath))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "filePath is required" };
        }

        try
        {
            TimeSpan? timeout = args.TimeoutSeconds.HasValue
                ? TimeSpan.FromSeconds(args.TimeoutSeconds.Value)
                : null;
            var sessionId = _sessionManager.CreateSession(args.FilePath, show: args.Show, operationTimeout: timeout, origin: SessionOrigin.CLI);
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, sessionId, filePath = args.FilePath }, ServiceProtocol.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return new ServiceResponse { Success = false, ErrorMessage = $"{ex.GetType().Name}: {ex.Message}" };
        }
    }

    private ServiceResponse HandleSessionClose(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        var args = ServiceRegistry.DeserializeArgs<SessionCloseArgs>(request.Args);
        var closed = _sessionManager.CloseSession(request.SessionId, save: args?.Save ?? false);

        return closed
            ? new ServiceResponse { Success = true }
            : new ServiceResponse { Success = false, ErrorMessage = $"Session '{request.SessionId}' not found" };
    }

    private ServiceResponse HandleSessionSave(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        var batch = _sessionManager.GetSession(request.SessionId);
        if (batch == null)
        {
            return new ServiceResponse { Success = false, ErrorMessage = $"Session '{request.SessionId}' not found" };
        }

        // Check if PowerPoint process is still alive before attempting save
        if (!batch.IsPowerPointProcessAlive())
        {
            _sessionManager.CloseSession(request.SessionId, save: false, force: true);
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"PowerPoint process for session '{request.SessionId}' has died. Session has been closed. Please create a new session."
            };
        }

        batch.Save();
        return new ServiceResponse { Success = true };
    }

    private ServiceResponse HandleSessionList()
    {
        var sessions = _sessionManager.GetActiveSessions()
            .Select(s => new
            {
                sessionId = s.SessionId,
                filePath = s.FilePath,
                isPowerPointVisible = _sessionManager.IsPowerPointVisible(s.SessionId),
                activeOperations = _sessionManager.GetActiveOperationCount(s.SessionId),
                canClose = _sessionManager.GetActiveOperationCount(s.SessionId) == 0
            })
            .ToList();

        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new { success = true, sessions, count = sessions.Count }, ServiceProtocol.JsonOptions)
        };
    }



    // === GENERATED DISPATCH ===

    // All command routing uses ServiceRegistry.*.DispatchToCore() generated methods.

    // See ServiceRegistry.*.Dispatch.g.cs for the generated code.



    private static ServiceResponse WrapResult(string? dispatchResult)
    {
        return dispatchResult == null
            ? new ServiceResponse { Success = true }
            : new ServiceResponse { Success = true, Result = dispatchResult };
    }

    /// <summary>
    /// Dispatches any generated service category via <see cref="ServiceRegistry.CategoryTable"/>.
    ///
    /// The table is generated from the same [ServiceCategory] interfaces that produce the MCP
    /// tools and CLI commands, so every advertised category is routable by construction.
    /// A hand-written switch previously covered only 22 of 33 categories (GitHub #124).
    /// </summary>
    private async Task<ServiceResponse> DispatchGeneratedAsync(string category, string actionString, ServiceRequest request)
    {
        if (!ServiceRegistry.CategoryTable.TryGetValue(category, out var entry))
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown command category: {category}" };

        // Validate the action before acquiring a session so an invalid action never
        // starts PowerPoint or reports a misleading session error.
        if (!entry.IsValidAction(actionString))
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown action: {actionString}" };

        var commands = _commandInstances[category];

        if (!entry.RequiresSession)
        {
            entry.Dispatch(commands, actionString, batch: null, request.Args, out var sessionlessResult);
            return WrapResult(sessionlessResult);
        }

        return await WithSessionAsync(request.SessionId, batch =>
        {
            entry.Dispatch(commands, actionString, batch, request.Args, out var result);
            return WrapResult(result);
        });
    }

    private Task<ServiceResponse> WithSessionAsync(string? sessionId, Func<IPptBatch, ServiceResponse> action)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return Task.FromResult(new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" });
        }

        var batch = _sessionManager.GetSession(sessionId);
        if (batch == null)
        {
            return Task.FromResult(new ServiceResponse { Success = false, ErrorMessage = $"Session '{sessionId}' not found" });
        }

        // Check if PowerPoint process is still alive before attempting operation
        if (!batch.IsPowerPointProcessAlive())
        {
            // PowerPoint died - clean up the dead session
            _sessionManager.CloseSession(sessionId, save: false, force: true);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"PowerPoint process for session '{sessionId}' has died. Session has been closed. Please create a new session."
            });
        }

        try
        {
            var response = action(batch);
            return Task.FromResult(response);
        }
        catch (TimeoutException ex)
        {
            // Operation timed out — PowerPoint COM call is hung.
            // Force-close the session to trigger the force-kill path in PptBatch.Dispose().
            _sessionManager.CloseSession(sessionId, save: false, force: true);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"PowerPoint operation timed out and the session has been closed: {ex.Message} " +
                               "Please reopen the file with a new session."
            });
        }
        catch (OperationCanceledException)
        {
            // Caller cancelled while a COM operation may still be running on the STA thread.
            _sessionManager.CloseSession(sessionId, save: false, force: true);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Operation was cancelled and the session has been closed. " +
                               "The PowerPoint COM thread may have been unresponsive. " +
                               "Please reopen the file with a new session."
            });
        }
        catch (COMException ex) when (
            ex.HResult == ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE ||
            ex.HResult == ResiliencePipelines.RPC_E_CALL_FAILED)
        {
            // PowerPoint process died during the operation — clean up the dead session
            _sessionManager.CloseSession(sessionId, save: false, force: true);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"PowerPoint process for session '{sessionId}' has died. " +
                               "Session has been cleaned up. Please reopen the file with a new session."
            });
        }
        catch (InvalidOperationException ex) when (
            ex.Message.Contains("no longer running", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("process", StringComparison.OrdinalIgnoreCase))
        {
            // PowerPoint process detected as dead before COM call (PptBatch pre-check)
            _sessionManager.CloseSession(sessionId, save: false, force: true);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"PowerPoint process for session '{sessionId}' is no longer running. " +
                               "Session has been cleaned up. Please reopen the file with a new session."
            });
        }
        catch (Exception ex)
        {
            return Task.FromResult(new ServiceResponse { Success = false, ErrorMessage = $"{ex.GetType().Name}: {ex.Message}" });
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _shutdownCts.Cancel();
        _sessionManager.Dispose();
        _shutdownCts.Dispose();
    }
}

// === ARGUMENT TYPES (Session only - all other args are now generated in ServiceRegistry) ===

// Session
public sealed class SessionOpenArgs
{
    public string? FilePath { get; set; }
    public bool Show { get; set; }
    public int? TimeoutSeconds { get; set; }
}
public sealed class SessionCloseArgs { public bool Save { get; set; } }
