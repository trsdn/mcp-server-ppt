using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace PptMcp.ComInterop.Session;

/// <summary>
/// Manages the active PowerPoint session for the MCP server and CLI.
/// Maps user-facing sessionId to internal IPptBatch instance.
/// </summary>
/// <remarks>
/// <para><b>CRITICAL: PowerPoint COM is single-instance.</b> Unlike Excel, creating multiple
/// PowerPoint.Application objects shares the same POWERPNT.EXE process. Quitting one kills all.
/// Therefore, only ONE session can be active at a time.</para>
/// <para><b>Concurrency Model:</b></para>
/// <list type="bullet">
/// <item><b>Only ONE session at a time:</b> PowerPoint COM is single-instance</item>
/// <item><b>Within-session operations are SERIAL:</b> Operations queue on one STA thread</item>
/// <item><b>Close before opening another:</b> Must close current session before opening a new file</item>
/// </list>
/// </remarks>
public sealed class SessionManager : IDisposable
{
    private readonly ConcurrentDictionary<string, IPptBatch> _activeSessions = new();
    private readonly ConcurrentDictionary<string, string> _activeFilePaths = new();
    private readonly ConcurrentDictionary<string, string> _sessionFilePaths = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, int> _activeOperationCounts = new();
    private readonly ConcurrentDictionary<string, bool> _showPowerPointFlags = new();
    private readonly ConcurrentDictionary<string, SessionOrigin> _sessionOrigins = new();
    private readonly ConcurrentDictionary<string, DateTime> _sessionCreatedAt = new();
    private readonly Polly.ResiliencePipeline _sessionCreationPipeline = ResiliencePipelines.CreateSessionCreationPipeline();
    private readonly ILogger<SessionManager> _logger;
    private bool _disposed;

    /// <summary>
    /// Creates a new SessionManager with optional logging.
    /// </summary>
    /// <param name="logger">Optional logger for diagnostics</param>
    public SessionManager(ILogger<SessionManager>? logger = null)
    {
        _logger = logger ?? NullLogger<SessionManager>.Instance;
    }

    /// <summary>
    /// Creates a new session for the specified PowerPoint file.
    /// </summary>
    /// <param name="filePath">Path to the PowerPoint file to open</param>
    /// <param name="show">Whether to show the PowerPoint window (default: false for background automation)</param>
    /// <param name="operationTimeout">Maximum time for any operation in this session (default: 5 minutes)</param>
    /// <param name="origin">Which client is creating this session (CLI or MCP)</param>
    /// <returns>Unique session ID for this session</returns>
    /// <exception cref="FileNotFoundException">File does not exist</exception>
    /// <exception cref="InvalidOperationException">Failed to create session, session already active, or file already open</exception>
    /// <remarks>
    /// <para><b>Single-session only:</b> PowerPoint COM is single-instance. Only one session can be active at a time.
    /// Close the current session before opening another file.</para>
    /// </remarks>
    public string CreateSession(string filePath, bool show = false, TimeSpan? operationTimeout = null, SessionOrigin origin = SessionOrigin.Unknown)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        // CRITICAL: PowerPoint COM is single-instance — only one session at a time
        if (!_activeSessions.IsEmpty)
        {
            throw new InvalidOperationException("PowerPoint COM is single-instance. Only one session can be active at a time. Close the current session before opening another file.");
        }

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"PowerPoint file not found: {filePath}. To create a new file, use the 'create' action instead of 'open'.", filePath);
        }

        // Normalize file path for comparison
        string normalizedPath = Path.GetFullPath(filePath);

        // Check if file is already open in another session
        if (_activeFilePaths.ContainsKey(normalizedPath))
        {
            throw new InvalidOperationException($"File '{filePath}' is already open in another session. PowerPoint cannot open the same file multiple times.");
        }

        // Generate unique session ID
        string sessionId = Guid.NewGuid().ToString("N");

        IPptBatch? batch = null;
        try
        {
            // Create batch session using Core API with retry for transient COM failures
            // (e.g., CO_E_SERVER_EXEC_FAILURE when system resources are constrained)
            batch = _sessionCreationPipeline.Execute(() => PptSession.BeginBatch(show, operationTimeout, filePath));

            // Store in active sessions
            if (!_activeSessions.TryAdd(sessionId, batch))
            {
                throw new InvalidOperationException($"Session ID collision: {sessionId}");
            }

            // Track the file path
            if (!_activeFilePaths.TryAdd(normalizedPath, sessionId))
            {
                // Cleanup if file path tracking fails
                _activeSessions.TryRemove(sessionId, out _);
                throw new InvalidOperationException($"Failed to track file path for session: {sessionId}");
            }

            if (!_sessionFilePaths.TryAdd(sessionId, normalizedPath))
            {
                _activeSessions.TryRemove(sessionId, out _);
                _activeFilePaths.TryRemove(normalizedPath, out _);
                throw new InvalidOperationException($"Failed to record session metadata for: {sessionId}");
            }

            // Initialize operation counter and show flag
            _activeOperationCounts[sessionId] = 0;
            _showPowerPointFlags[sessionId] = show;
            _sessionOrigins[sessionId] = origin;
            _sessionCreatedAt[sessionId] = DateTime.UtcNow;

            // Success - transfer ownership to dictionary
            var result = sessionId;
            batch = null;  // Prevent disposal in finally
            return result;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to create session for '{filePath}': {ex.Message}", ex);
        }
        finally
        {
            // Dispose batch only if we didn't successfully add it to dictionary
            batch?.Dispose();
        }
    }

    /// <summary>
    /// Creates a new PowerPoint file and opens a session for it in one operation.
    /// This is the preferred method for creating new Presentations with sessions.
    /// </summary>
    /// <param name="filePath">Path for the new PowerPoint file (.pptx or .pptm)</param>
    /// <param name="show">Whether to show the PowerPoint window (default: false)</param>
    /// <param name="operationTimeout">Maximum time for any operation in this session (default: 5 minutes)</param>
    /// <param name="origin">Which client is creating this session (CLI or MCP)</param>
    /// <returns>Unique session ID for this session</returns>
    /// <exception cref="InvalidOperationException">File already exists, or failed to create session</exception>
    /// <exception cref="DirectoryNotFoundException">Target directory does not exist</exception>
    /// <remarks>
    /// <para><b>Single PowerPoint Start:</b> This method starts PowerPoint only once, creating the file and session together.</para>
    /// <para><b>File Format:</b> Determined by extension - .pptm creates macro-enabled Presentation.</para>
    /// <para><b>Directory:</b> Target directory must exist - will not be created automatically.</para>
    /// </remarks>
    public string CreateSessionForNewFile(string filePath, bool show = false, TimeSpan? operationTimeout = null, SessionOrigin origin = SessionOrigin.Unknown)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        // CRITICAL: PowerPoint COM is single-instance — only one session at a time
        if (!_activeSessions.IsEmpty)
        {
            throw new InvalidOperationException("PowerPoint COM is single-instance. Only one session can be active at a time. Close the current session before creating a new file.");
        }

        string normalizedPath = Path.GetFullPath(filePath);

        // Validate extension
        string extension = Path.GetExtension(normalizedPath).ToLowerInvariant();
        if (extension is not (".pptx" or ".pptm"))
        {
            throw new ArgumentException($"Invalid file extension '{extension}'. Only .pptx and .pptm are supported.");
        }

        // Check if file already exists
        if (File.Exists(normalizedPath))
        {
            throw new InvalidOperationException($"File already exists: {normalizedPath}. Use CreateSession to open existing files.");
        }

        // Check if file is already open in another session
        if (_activeFilePaths.ContainsKey(normalizedPath))
        {
            throw new InvalidOperationException($"File '{filePath}' is already open in another session.");
        }

        // Generate unique session ID
        string sessionId = Guid.NewGuid().ToString("N");
        bool isMacroEnabled = extension == ".pptm";

        PptBatch? batch = null;
        try
        {
            // Create new Presentation and keep session open with retry for transient COM failures
            batch = _sessionCreationPipeline.Execute(() => PptBatch.CreateNewPresentation(normalizedPath, isMacroEnabled, logger: null, show: show, operationTimeout: operationTimeout));

            // Store in active sessions
            if (!_activeSessions.TryAdd(sessionId, batch))
            {
                throw new InvalidOperationException($"Session ID collision: {sessionId}");
            }

            // Track the file path
            if (!_activeFilePaths.TryAdd(normalizedPath, sessionId))
            {
                _activeSessions.TryRemove(sessionId, out _);
                throw new InvalidOperationException($"Failed to track file path for session: {sessionId}");
            }

            if (!_sessionFilePaths.TryAdd(sessionId, normalizedPath))
            {
                _activeSessions.TryRemove(sessionId, out _);
                _activeFilePaths.TryRemove(normalizedPath, out _);
                throw new InvalidOperationException($"Failed to record session metadata for: {sessionId}");
            }

            // Initialize operation counter and show flag
            _activeOperationCounts[sessionId] = 0;
            _showPowerPointFlags[sessionId] = show;
            _sessionOrigins[sessionId] = origin;
            _sessionCreatedAt[sessionId] = DateTime.UtcNow;

            // Success - transfer ownership to dictionary
            var result = sessionId;
            batch = null;  // Prevent disposal in finally
            return result;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to create session for new file '{filePath}': {ex.Message}", ex);
        }
        finally
        {
            // Dispose batch only if we didn't successfully add it to dictionary
            batch?.Dispose();
        }
    }



    /// <summary>
    /// Gets an active session by ID.
    /// If the session exists but PowerPoint has died, it is automatically cleaned up and null is returned.
    /// </summary>
    /// <param name="sessionId">Session ID returned from CreateSession</param>
    /// <returns>IPptBatch instance, or null if session not found or PowerPoint process is dead</returns>
    public IPptBatch? GetSession(string sessionId)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return null;
        }

        if (!_activeSessions.TryGetValue(sessionId, out var batch))
        {
            return null;
        }

        // Check if PowerPoint process is still alive
        if (!batch.IsPowerPointProcessAlive())
        {
            _logger?.LogWarning("Session {SessionId} has dead PowerPoint process, auto-cleaning up", sessionId);
            CleanupDeadSession(sessionId, batch);
            return null;
        }

        return batch;
    }

    /// <summary>
    /// Cleans up a session whose PowerPoint process has died.
    /// This removes all tracking data and disposes the batch (best effort).
    /// </summary>
    private void CleanupDeadSession(string sessionId, IPptBatch batch)
    {
        // Remove from active sessions
        _activeSessions.TryRemove(sessionId, out _);

        // Remove file path metadata so it can be opened again
        if (_sessionFilePaths.TryRemove(sessionId, out var normalizedPath))
        {
            _activeFilePaths.TryRemove(normalizedPath, out _);
        }
        else
        {
            var filePathEntry = _activeFilePaths.FirstOrDefault(kvp => kvp.Value == sessionId);
            if (!filePathEntry.Equals(default(KeyValuePair<string, string>)))
            {
                _activeFilePaths.TryRemove(filePathEntry.Key, out _);
            }
        }

        // Clean up operation tracking data
        _activeOperationCounts.TryRemove(sessionId, out _);
        _showPowerPointFlags.TryRemove(sessionId, out _);

        // Clean up session origin tracking data
        _sessionOrigins.TryRemove(sessionId, out _);
        _sessionCreatedAt.TryRemove(sessionId, out _);

        // Dispose the batch (best effort - process is already dead)
        try
        {
            batch.Dispose();
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex, "Error disposing dead session {SessionId} (expected - process is dead)", sessionId);
        }
    }

    /// <summary>
    /// Increments the active operation count for a session.
    /// Call this when starting an operation on the session.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    public void BeginOperation(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return;
        _activeOperationCounts.AddOrUpdate(sessionId, 1, (_, count) => count + 1);
    }

    /// <summary>
    /// Decrements the active operation count for a session.
    /// Call this when an operation completes (success or failure).
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    public void EndOperation(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return;
        _activeOperationCounts.AddOrUpdate(sessionId, 0, (_, count) => Math.Max(0, count - 1));
    }

    /// <summary>
    /// Gets the number of active operations for a session.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>Number of active operations, or 0 if session not found</returns>
    public int GetActiveOperationCount(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return 0;
        return _activeOperationCounts.TryGetValue(sessionId, out var count) ? count : 0;
    }

    /// <summary>
    /// Gets whether PowerPoint is visible for a session.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>True if PowerPoint is visible for this session</returns>
    public bool IsPowerPointVisible(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return false;
        return _showPowerPointFlags.TryGetValue(sessionId, out var visible) && visible;
    }

    /// <summary>
    /// Updates the visibility flag for a session.
    /// Called by window management commands when PowerPoint visibility changes mid-session.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <param name="visible">New visibility state</param>
    /// <returns>True if session was found and flag updated, false if session not found</returns>
    public bool SetPowerPointVisible(string sessionId, bool visible)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return false;
        if (!_activeSessions.ContainsKey(sessionId)) return false;
        _showPowerPointFlags[sessionId] = visible;
        return true;
    }

    /// <summary>
    /// Validates whether a session can be closed safely.
    /// Returns information about blocking conditions.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>Validation result with details about any blocking conditions</returns>
    public CloseValidationResult ValidateClose(string sessionId)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return new CloseValidationResult(false, false, 0, "Session ID is required");
        }

        if (!_activeSessions.ContainsKey(sessionId))
        {
            return new CloseValidationResult(false, false, 0, $"Session '{sessionId}' not found");
        }

        var activeOps = GetActiveOperationCount(sessionId);
        var isVisible = IsPowerPointVisible(sessionId);

        if (activeOps > 0)
        {
            return new CloseValidationResult(true, isVisible, activeOps,
                $"Cannot close: {activeOps} operation(s) still running. Wait for operations to complete before closing.");
        }

        return new CloseValidationResult(true, isVisible, 0, null);
    }

    /// <summary>
    /// Closes the specified session with optional save.
    /// If save is true, saves changes before closing to ensure atomic operation.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <param name="save">Whether to save changes before closing (default: false)</param>
    /// <param name="force">Force close even if operations are running (default: false)</param>
    /// <returns>True if session was found and closed, false if session not found</returns>
    /// <exception cref="InvalidOperationException">Save operation failed or operations still running</exception>
    public bool CloseSession(string sessionId, bool save = false, bool force = false)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return false;
        }

        // Check for running operations (unless force is true)
        if (!force)
        {
            var activeOps = GetActiveOperationCount(sessionId);
            if (activeOps > 0)
            {
                throw new InvalidOperationException(
                    $"Cannot close session '{sessionId}': {activeOps} operation(s) still running. " +
                    "Wait for all operations to complete before closing, or use force=true to close anyway.");
            }
        }

        // Save first if requested (blocks until complete)
        if (save)
        {
            var batch = GetSession(sessionId);
            if (batch != null)
            {
                try
                {
                    batch.Save();
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Failed to save session '{sessionId}' before closing: {ex.Message}", ex);
                }
            }
        }

        // Then close
        return CloseSessionSync(sessionId);
    }

    private bool CloseSessionSync(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return false;
        }

        if (!_activeSessions.TryRemove(sessionId, out var batch))
        {
            return false;
        }

        // Remove file path metadata so it can be opened again
        if (_sessionFilePaths.TryRemove(sessionId, out var normalizedPath))
        {
            _activeFilePaths.TryRemove(normalizedPath, out _);
        }
        else
        {
            var filePathEntry = _activeFilePaths.FirstOrDefault(kvp => kvp.Value == sessionId);
            if (!filePathEntry.Equals(default(KeyValuePair<string, string>)))
            {
                _activeFilePaths.TryRemove(filePathEntry.Key, out _);
            }
        }

        // Clean up operation tracking data
        _activeOperationCounts.TryRemove(sessionId, out _);
        _showPowerPointFlags.TryRemove(sessionId, out _);

        // Clean up session origin tracking data
        _sessionOrigins.TryRemove(sessionId, out _);
        _sessionCreatedAt.TryRemove(sessionId, out _);

        try
        {
            batch.Dispose();
            return true;
        }
        catch (Exception)
        {
            // Best-effort — session is already removed from dictionary
            return true;
        }
    }

    /// <summary>
    /// Gets the number of active sessions.
    /// Note: This count may include dead sessions. Use <see cref="GetActiveSessions"/> for accurate count.
    /// </summary>
    public int ActiveSessionCount => _activeSessions.Count;

    /// <summary>
    /// Checks if the PowerPoint process for a session is still alive.
    /// If the session exists but PowerPoint has died, it is automatically cleaned up.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>True if session exists and PowerPoint process is alive, false otherwise</returns>
    public bool IsSessionAlive(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return false;
        if (!_activeSessions.TryGetValue(sessionId, out var batch)) return false;

        if (batch.IsPowerPointProcessAlive())
        {
            return true;
        }

        // Auto-cleanup dead session
        _logger?.LogWarning("Session {SessionId} has dead PowerPoint process, auto-cleaning up during IsSessionAlive check", sessionId);
        CleanupDeadSession(sessionId, batch);
        return false;
    }

    /// <summary>
    /// Gets all active session IDs.
    /// Note: This property does not filter dead sessions. Use <see cref="GetActiveSessions"/> for filtered results.
    /// </summary>
    public IEnumerable<string> ActiveSessionIds => _activeSessions.Keys.ToList();

    /// <summary>
    /// Returns a snapshot of active sessions with associated Presentation paths.
    /// Dead sessions (where PowerPoint process has died) are automatically cleaned up and excluded.
    /// </summary>
    public IReadOnlyList<SessionDescriptor> GetActiveSessions()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var snapshot = new List<SessionDescriptor>(_sessionFilePaths.Count);
        var deadSessions = new List<(string sessionId, IPptBatch batch)>();

        foreach (var kvp in _sessionFilePaths)
        {
            var sessionId = kvp.Key;

            // Check if session is still alive
            if (_activeSessions.TryGetValue(sessionId, out var batch))
            {
                if (batch.IsPowerPointProcessAlive())
                {
                    // Get origin and createdAt metadata (defaults for legacy sessions)
                    _sessionOrigins.TryGetValue(sessionId, out var origin);
                    _sessionCreatedAt.TryGetValue(sessionId, out var createdAt);

                    snapshot.Add(new SessionDescriptor(sessionId, kvp.Value, origin, createdAt == default ? null : createdAt));
                }
                else
                {
                    // Mark for cleanup (don't cleanup during iteration)
                    deadSessions.Add((sessionId, batch));
                }
            }
            // If not in _activeSessions but in _sessionFilePaths, skip (orphaned metadata)
        }

        // Clean up dead sessions after iteration
        foreach (var (sessionId, batch) in deadSessions)
        {
            _logger?.LogWarning("Session {SessionId} has dead PowerPoint process, auto-cleaning up during GetActiveSessions", sessionId);
            CleanupDeadSession(sessionId, batch);
        }

        return snapshot;
    }

    /// <summary>
    /// Attempts to get the Presentation path associated with a session ID.
    /// </summary>
    public bool TryGetFilePath(string sessionId, [NotNullWhen(true)] out string? filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            filePath = null;
            return false;
        }

        return _sessionFilePaths.TryGetValue(sessionId, out filePath);
    }

    /// <summary>
    /// Disposes all active sessions, auto-saving each one first to prevent data loss.
    /// </summary>
    /// <remarks>
    /// <para><b>CRITICAL:</b> Sessions are auto-saved before disposal to prevent silent data loss
    /// when the service shuts down (e.g., MCP client disconnect, process exit).</para>
    /// <para><b>CRITICAL:</b> Sessions are disposed SEQUENTIALLY to avoid COM threading issues.</para>
    /// <para>PowerPoint COM objects must be disposed on their STA threads. Parallel disposal causes deadlocks.</para>
    /// </remarks>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        // Close all active sessions SEQUENTIALLY to avoid COM threading issues
        // PowerPoint COM objects must be disposed on their STA threads, parallel disposal causes deadlocks
        var sessions = _activeSessions.Values.ToList();
        _activeSessions.Clear();
        _activeFilePaths.Clear();
        _sessionFilePaths.Clear();

        foreach (var session in sessions)
        {
            // Auto-save before disposal to prevent silent data loss.
            // This protects against the common scenario where the MCP client disconnects
            // or the service process exits, which would otherwise discard all unsaved work.
            if (session.IsPowerPointProcessAlive())
            {
                try
                {
                    using var saveTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                    session.Save(saveTimeout.Token);
                    _logger.LogInformation("Auto-saved session for {Path} before shutdown", session.PresentationPath);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to auto-save session for {Path} before shutdown (changes may be lost)", session.PresentationPath);
                }
            }

            try
            {
                // Dispose sequentially - PptBatch.Dispose() handles its own PowerPoint cleanup
                // via PptShutdownService with proper timeouts and retry logic
                session.Dispose();
            }
            catch (Exception)
            {
                // Best-effort cleanup — continue with remaining sessions
            }
        }
    }
}

/// <summary>
/// Represents a snapshot of an active PowerPoint session managed by <see cref="SessionManager"/>.
/// </summary>
/// <param name="SessionId">Public session identifier shared with clients.</param>
/// <param name="FilePath">Normalized Presentation path associated with the session.</param>
/// <param name="Origin">Which client created this session (CLI or MCP).</param>
/// <param name="CreatedAt">When the session was created.</param>
public sealed record SessionDescriptor(
    string SessionId,
    string FilePath,
    SessionOrigin Origin = SessionOrigin.Unknown,
    DateTime? CreatedAt = null);

/// <summary>
/// Indicates which client created a session.
/// </summary>
public enum SessionOrigin
{
    /// <summary>Session origin is unknown (legacy sessions).</summary>
    Unknown = 0,

    /// <summary>Session was created via the CLI.</summary>
    CLI = 1,

    /// <summary>Session was created via the MCP Server.</summary>
    MCP = 2
}

/// <summary>
/// Result of validating whether a session can be closed.
/// </summary>
/// <param name="SessionExists">Whether the session was found.</param>
/// <param name="IsPowerPointVisible">Whether PowerPoint is visible (show=true).</param>
/// <param name="ActiveOperationCount">Number of operations currently running.</param>
/// <param name="BlockingReason">Reason why close is blocked, or null if close is allowed.</param>
public sealed record CloseValidationResult(
    bool SessionExists,
    bool IsPowerPointVisible,
    int ActiveOperationCount,
    string? BlockingReason)
{
    /// <summary>
    /// Whether the session can be closed (no blocking conditions).
    /// </summary>
    public bool CanClose => SessionExists && ActiveOperationCount == 0;
}



