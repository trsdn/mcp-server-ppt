using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Text.Json;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Accessibility;
using PptMcp.Core.Commands.Animation;
using PptMcp.Core.Commands.Chart;
using PptMcp.Core.Commands.Design;
using PptMcp.Core.Commands.DocumentProperty;
using PptMcp.Core.Commands.Export;
using PptMcp.Core.Commands.File;
using PptMcp.Core.Commands.Hyperlink;
using PptMcp.Core.Commands.Image;
using PptMcp.Core.Commands.Master;
using PptMcp.Core.Commands.Media;
using PptMcp.Core.Commands.Notes;
using PptMcp.Core.Commands.Proofing;
using PptMcp.Core.Commands.Section;
using PptMcp.Core.Commands.Shape;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Commands.SlideTable;
using PptMcp.Core.Commands.Slideshow;
using PptMcp.Core.Commands.Text;
using PptMcp.Core.Commands.Transition;
using PptMcp.Core.Commands.Vba;
using PptMcp.Core.Commands.Window;
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

    // Core command instances
    private readonly FileCommands _fileCommands = new();
    private readonly SlideCommands _slideCommands = new();
    private readonly ShapeCommands _shapeCommands = new();
    private readonly TextCommands _textCommands = new();
    private readonly NotesCommands _notesCommands = new();
    private readonly MasterCommands _masterCommands = new();
    private readonly ExportCommands _exportCommands = new();
    private readonly TransitionCommands _transitionCommands = new();
    private readonly ImageCommands _imageCommands = new();
    private readonly SlideTableCommands _slideTableCommands = new();
    private readonly ChartCommands _chartCommands = new();
    private readonly AnimationCommands _animationCommands = new();
    private readonly DesignCommands _designCommands = new();
    private readonly SlideshowCommands _slideshowCommands = new();
    private readonly VbaCommands _vbaCommands = new();
    private readonly WindowCommands _windowCommands = new();
    private readonly HyperlinkCommands _hyperlinkCommands = new();
    private readonly SectionCommands _sectionCommands = new();
    private readonly DocumentPropertyCommands _documentPropertyCommands = new();
    private readonly MediaCommands _mediaCommands = new();
    private readonly ProofingCommands _proofingCommands = new();
    private readonly AccessibilityCommands _accessibilityCommands = new();

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
                "file" => DispatchSessionless(action, request),
                "slide" => await DispatchSimpleAsync<SlideAction>(action, request,
                    ServiceRegistry.Slide.TryParseAction,
                    (a, batch) => ServiceRegistry.Slide.DispatchToCore(_slideCommands, a, batch, request.Args)),
                "shape" => await DispatchSimpleAsync<ShapeAction>(action, request,
                    ServiceRegistry.Shape.TryParseAction,
                    (a, batch) => ServiceRegistry.Shape.DispatchToCore(_shapeCommands, a, batch, request.Args)),
                "text" => await DispatchSimpleAsync<TextAction>(action, request,
                    ServiceRegistry.Text.TryParseAction,
                    (a, batch) => ServiceRegistry.Text.DispatchToCore(_textCommands, a, batch, request.Args)),
                "notes" => await DispatchSimpleAsync<NotesAction>(action, request,
                    ServiceRegistry.Notes.TryParseAction,
                    (a, batch) => ServiceRegistry.Notes.DispatchToCore(_notesCommands, a, batch, request.Args)),
                "master" => await DispatchSimpleAsync<MasterAction>(action, request,
                    ServiceRegistry.Master.TryParseAction,
                    (a, batch) => ServiceRegistry.Master.DispatchToCore(_masterCommands, a, batch, request.Args)),
                "export" => await DispatchSimpleAsync<ExportAction>(action, request,
                    ServiceRegistry.Export.TryParseAction,
                    (a, batch) => ServiceRegistry.Export.DispatchToCore(_exportCommands, a, batch, request.Args)),
                "transition" => await DispatchSimpleAsync<TransitionAction>(action, request,
                    ServiceRegistry.Transition.TryParseAction,
                    (a, batch) => ServiceRegistry.Transition.DispatchToCore(_transitionCommands, a, batch, request.Args)),
                "image" => await DispatchSimpleAsync<ImageAction>(action, request,
                    ServiceRegistry.Image.TryParseAction,
                    (a, batch) => ServiceRegistry.Image.DispatchToCore(_imageCommands, a, batch, request.Args)),
                "slidetable" => await DispatchSimpleAsync<SlidetableAction>(action, request,
                    ServiceRegistry.Slidetable.TryParseAction,
                    (a, batch) => ServiceRegistry.Slidetable.DispatchToCore(_slideTableCommands, a, batch, request.Args)),
                "chart" => await DispatchSimpleAsync<ChartAction>(action, request,
                    ServiceRegistry.Chart.TryParseAction,
                    (a, batch) => ServiceRegistry.Chart.DispatchToCore(_chartCommands, a, batch, request.Args)),
                "animation" => await DispatchSimpleAsync<AnimationAction>(action, request,
                    ServiceRegistry.Animation.TryParseAction,
                    (a, batch) => ServiceRegistry.Animation.DispatchToCore(_animationCommands, a, batch, request.Args)),
                "design" => await DispatchSimpleAsync<DesignAction>(action, request,
                    ServiceRegistry.Design.TryParseAction,
                    (a, batch) => ServiceRegistry.Design.DispatchToCore(_designCommands, a, batch, request.Args)),
                "slideshow" => await DispatchSimpleAsync<SlideshowAction>(action, request,
                    ServiceRegistry.Slideshow.TryParseAction,
                    (a, batch) => ServiceRegistry.Slideshow.DispatchToCore(_slideshowCommands, a, batch, request.Args)),
                "vba" => await DispatchSimpleAsync<VbaAction>(action, request,
                    ServiceRegistry.Vba.TryParseAction,
                    (a, batch) => ServiceRegistry.Vba.DispatchToCore(_vbaCommands, a, batch, request.Args)),
                "window" => await DispatchSimpleAsync<WindowAction>(action, request,
                    ServiceRegistry.Window.TryParseAction,
                    (a, batch) => ServiceRegistry.Window.DispatchToCore(_windowCommands, a, batch, request.Args)),
                "hyperlink" => await DispatchSimpleAsync<HyperlinkAction>(action, request,
                    ServiceRegistry.Hyperlink.TryParseAction,
                    (a, batch) => ServiceRegistry.Hyperlink.DispatchToCore(_hyperlinkCommands, a, batch, request.Args)),
                "section" => await DispatchSimpleAsync<SectionAction>(action, request,
                    ServiceRegistry.Section.TryParseAction,
                    (a, batch) => ServiceRegistry.Section.DispatchToCore(_sectionCommands, a, batch, request.Args)),
                "docproperty" => await DispatchSimpleAsync<DocpropertyAction>(action, request,
                    ServiceRegistry.Docproperty.TryParseAction,
                    (a, batch) => ServiceRegistry.Docproperty.DispatchToCore(_documentPropertyCommands, a, batch, request.Args)),
                "media" => await DispatchSimpleAsync<MediaAction>(action, request,
                    ServiceRegistry.Media.TryParseAction,
                    (a, batch) => ServiceRegistry.Media.DispatchToCore(_mediaCommands, a, batch, request.Args)),
                "proofing" => await DispatchSimpleAsync<ProofingAction>(action, request,
                    ServiceRegistry.Proofing.TryParseAction,
                    (a, batch) => ServiceRegistry.Proofing.DispatchToCore(_proofingCommands, a, batch, request.Args)),
                "accessibility" => await DispatchSimpleAsync<AccessibilityAction>(action, request,
                    ServiceRegistry.Accessibility.TryParseAction,
                    (a, batch) => ServiceRegistry.Accessibility.DispatchToCore(_accessibilityCommands, a, batch, request.Args)),
                _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown command category: {category}" }
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



    private delegate bool TryParseDelegate<TAction>(string action, out TAction result);



    private static ServiceResponse WrapResult(string? dispatchResult)

    {

        return dispatchResult == null

            ? new ServiceResponse { Success = true }

            : new ServiceResponse { Success = true, Result = dispatchResult };

    }



    private async Task<ServiceResponse> DispatchSimpleAsync<TAction>(

        string actionString, ServiceRequest request,

        TryParseDelegate<TAction> tryParse,

        Func<TAction, IPptBatch, string?> dispatch) where TAction : struct

    {

        if (!tryParse(actionString, out var action))

            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown action: {actionString}" };



        return await WithSessionAsync(request.SessionId, batch => WrapResult(dispatch(action, batch)));

    }

    /// <summary>
    /// Dispatches a session-less command (no PowerPoint batch required).
    /// Used for [NoSession] categories like file.
    /// </summary>
    private ServiceResponse DispatchSessionless(string actionString, ServiceRequest request)
    {
        if (!ServiceRegistry.File.TryParseAction(actionString, out var action))
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown action: {actionString}" };

        return WrapResult(ServiceRegistry.File.DispatchToCore(_fileCommands, action, request.Args));
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
