using System.Runtime.InteropServices;
using System.Threading.Channels;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptMcp.ComInterop.Session;

/// <summary>
/// Implementation of IPptBatch that manages a single PowerPoint instance on a dedicated STA thread.
/// Ensures proper COM interop with PowerPoint using STA apartment state and OLE message filter.
/// </summary>
/// <remarks>
/// <para><b>CRITICAL: PowerPoint COM Threading Model</b></para>
/// <list type="bullet">
/// <item>Each PptBatch runs on ONE dedicated STA (Single-Threaded Apartment) thread</item>
/// <item>Operations are queued via Channel and executed SERIALLY (never in parallel)</item>
/// <item>Multiple simultaneous Execute() calls are processed one at a time</item>
/// <item>This is a COM interop requirement, not an implementation choice</item>
/// <item>For parallel processing, create multiple sessions for DIFFERENT files</item>
/// </list>
/// <para><b>Resource Cost:</b> Each PptBatch = one PowerPoint.Application process (~50-100MB+ memory)</para>
/// </remarks>
internal sealed class PptBatch : IPptBatch
{
    // P/Invoke for getting process ID from window handle
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    private readonly string _presentationPath; // Primary presentation path
    private readonly string[] _allPresentationPaths; // All presentation paths (includes primary)
    private readonly bool _showPowerPoint; // Whether to show PowerPoint window
    private readonly bool _createNewFile; // Whether to create a new file instead of opening existing
    private readonly bool _isMacroEnabled; // For new files: whether to create .pptm (macro-enabled)
    private readonly TimeSpan _operationTimeout; // Timeout for individual operations
    private readonly ILogger<PptBatch> _logger;
    private readonly Channel<Func<Task>> _workQueue;
    private readonly Thread _staThread;
    private readonly CancellationTokenSource _shutdownCts;
    private int _disposed; // 0 = not disposed, 1 = disposed (using int for Interlocked.CompareExchange)
    private int? _powerPointProcessId; // POWERPNT.exe process ID for force-kill if needed
    private bool _operationTimedOut; // Track if an operation timed out for aggressive cleanup

    // COM state (STA thread only)
    private PowerPoint.Application? _powerPoint;
    private PowerPoint.Presentation? _presentation; // Primary presentation
    private Dictionary<string, PowerPoint.Presentation>? _presentations; // All presentations keyed by normalized path
    private PptContext? _context;

    /// <summary>
    /// Creates a new PptBatch for one or more presentations.
    /// All presentations are opened in the same PowerPoint.Application instance, enabling cross-presentation operations.
    /// </summary>
    /// <param name="presentationPaths">Paths to PowerPoint presentations. First path is the primary presentation.</param>
    /// <param name="logger">Optional logger for diagnostic output. If null, uses NullLogger (no output).</param>
    /// <param name="show">Whether to show the PowerPoint window (default: false for background automation).</param>
    /// <param name="operationTimeout">Timeout for individual operations. Default: 5 minutes.</param>
    public PptBatch(string[] presentationPaths, ILogger<PptBatch>? logger = null, bool show = false, TimeSpan? operationTimeout = null)
        : this(presentationPaths, logger, show, createNewFile: false, isMacroEnabled: false, operationTimeout: operationTimeout)
    {
    }

    /// <summary>
    /// Creates a new PptBatch that creates a new presentation file instead of opening an existing one.
    /// The file is saved immediately after creation, then kept open in the session.
    /// </summary>
    /// <param name="filePath">Path where the new PowerPoint file will be created.</param>
    /// <param name="isMacroEnabled">Whether to create .pptm (macro-enabled) format.</param>
    /// <param name="logger">Optional logger for diagnostic output.</param>
    /// <param name="show">Whether to show the PowerPoint window.</param>
    /// <param name="operationTimeout">Timeout for individual operations. Default: 5 minutes.</param>
    /// <returns>PptBatch instance with the new presentation open.</returns>
    internal static PptBatch CreateNewPresentation(string filePath, bool isMacroEnabled, ILogger<PptBatch>? logger = null, bool show = false, TimeSpan? operationTimeout = null)
    {
        return new PptBatch([filePath], logger, show, createNewFile: true, isMacroEnabled: isMacroEnabled, operationTimeout: operationTimeout);
    }

    /// <summary>
    /// Private constructor that handles both opening existing files and creating new ones.
    /// </summary>
    private PptBatch(string[] presentationPaths, ILogger<PptBatch>? logger, bool show, bool createNewFile, bool isMacroEnabled, TimeSpan? operationTimeout = null)
    {
        if (presentationPaths == null || presentationPaths.Length == 0)
            throw new ArgumentException("At least one presentation path is required", nameof(presentationPaths));

        _allPresentationPaths = presentationPaths;
        _presentationPath = presentationPaths[0]; // Primary presentation
        _showPowerPoint = show;
        _createNewFile = createNewFile;
        _isMacroEnabled = isMacroEnabled;
        _operationTimeout = operationTimeout ?? ComInteropConstants.DefaultOperationTimeout;
        _logger = logger ?? NullLogger<PptBatch>.Instance;
        _shutdownCts = new CancellationTokenSource();

        // Create unbounded channel for work items
        _workQueue = Channel.CreateUnbounded<Func<Task>>(new UnboundedChannelOptions
        {
            SingleReader = true,
            SingleWriter = false
        });

        // Start STA thread with message pump
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        _staThread = new Thread(() =>
        {
            try
            {
                // CRITICAL: Register OLE message filter on STA thread for PowerPoint busy handling
                OleMessageFilter.Register();

                // Create PowerPoint ON THIS STA THREAD
                Type? pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                if (pptType == null)
                {
                    throw new InvalidOperationException("Microsoft PowerPoint is not installed on this system.");
                }

                PowerPoint.Application tempPowerPoint = (PowerPoint.Application)Activator.CreateInstance(pptType)!;
                // PowerPoint COM does NOT allow hiding the application window (unlike Excel).
                // Setting Visible = msoFalse (0) throws "Hiding the application window is not allowed."
                // Always set Visible = msoTrue (-1). Use WindowState to minimize if needed.
                ((dynamic)tempPowerPoint).Visible = -1; // msoTrue — required by PowerPoint COM
                ((dynamic)tempPowerPoint).DisplayAlerts = 0; // ppAlertsNone
                if (!_showPowerPoint)
                {
                    // Minimize instead of hiding — PowerPoint doesn't allow Visible=false
                    ((dynamic)tempPowerPoint).WindowState = 2; // ppWindowMinimized
                }

                // Capture PowerPoint process ID for force-kill scenarios (hung PowerPoint, dead RPC connection)
                try
                {
                    // PowerPoint.Application.HWND returns the window handle
                    int hwnd = tempPowerPoint.HWND;
                    if (hwnd != 0)
                    {
                        uint processId = 0;
                        _ = GetWindowThreadProcessId(new IntPtr(hwnd), out processId);
                        if (processId != 0)
                        {
                            _powerPointProcessId = (int)processId;
                            _logger.LogDebug("Captured PowerPoint process ID via HWND: {ProcessId}", _powerPointProcessId);
                        }
                    }

                    if (!_powerPointProcessId.HasValue)
                    {
                        _logger.LogWarning(
                            "Could not determine PowerPoint process ID via HWND. " +
                            "Force-kill will be disabled for this session to avoid killing unrelated PowerPoint instances.");
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to capture PowerPoint process ID. Force-kill will not be available.");
                }

                // Open or create presentations in the same PowerPoint instance
                var tempPresentations = new Dictionary<string, PowerPoint.Presentation>(StringComparer.OrdinalIgnoreCase);
                PowerPoint.Presentation? primaryPresentation = null;

                foreach (var path in _allPresentationPaths)
                {
                    PowerPoint.Presentation pres;
                    string normalizedPath = Path.GetFullPath(path);

                    if (_createNewFile)
                    {
                        // CREATE NEW FILE: Use Add() + SaveAs() instead of Open()
                        string? directory = Path.GetDirectoryName(normalizedPath);
                        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                        {
                            throw new DirectoryNotFoundException($"Directory does not exist: '{directory}'. Create the directory first before creating PowerPoint files.");
                        }

                        // Create new presentation WITH window (required for embedded objects like charts)
                        // PowerPoint's AddChart/AddChart2 requires a DocumentWindow to exist.
                        // Using Add(msoFalse) creates without window and breaks chart/OLE operations.
                        pres = ((dynamic)tempPowerPoint).Presentations.Add();

                        // SaveAs with appropriate format
                        if (_isMacroEnabled)
                        {
                            ((dynamic)pres).SaveAs(normalizedPath, ComInteropConstants.PpSaveAsOpenXMLPresentationMacroEnabled);
                        }
                        else
                        {
                            ((dynamic)pres).SaveAs(normalizedPath, ComInteropConstants.PpSaveAsOpenXMLPresentation);
                        }
                    }
                    else
                    {
                        // OPEN EXISTING FILE: Validate and open
                        bool isIrm = FileAccessValidator.IsIrmProtected(normalizedPath);

                        if (isIrm)
                        {
                            ((dynamic)tempPowerPoint).Visible = -1 /* msoTrue */;
                            _logger.LogDebug(
                                "IRM-protected file detected: {FileName}. Forcing PowerPoint visible and opening read-only.",
                                Path.GetFileName(normalizedPath));
                        }
                        else
                        {
                            // CRITICAL: Check if file is locked at OS level BEFORE attempting PowerPoint COM open
                            FileAccessValidator.ValidateFileNotLocked(path);
                        }

                        // Open presentation with PowerPoint COM
                        try
                        {
                            pres = isIrm
                                ? ((dynamic)tempPowerPoint).Presentations.Open(normalizedPath, -1 /* msoTrue ReadOnly */)
                                : ((dynamic)tempPowerPoint).Presentations.Open(normalizedPath);
                        }
                        catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC))
                        {
                            // Error 1004 equivalent - File is already open or locked
                            throw FileAccessValidator.CreateFileLockedError(path, ex);
                        }
                    }

                    tempPresentations[normalizedPath] = pres;

                    if (path == _presentationPath)
                    {
                        primaryPresentation = pres;
                    }
                }

                _powerPoint = tempPowerPoint;
                _presentation = primaryPresentation;
                _presentations = tempPresentations;
                _context = new PptContext(_presentationPath, _powerPoint, _presentation!);

                started.SetResult();

                // Message pump - process work queue until completion or cancellation.
                // CRITICAL: Uses WaitToReadAsync() instead of polling with Thread.Sleep(10).
                //
                // Why WaitToReadAsync and not polling:
                // 1. Thread.Sleep(10) on an STA thread with registered OLE message filter is unreliable.
                //    Pending COM messages (_pptApp events during calculation) cause Sleep to return
                //    immediately via MsgWaitForMultipleObjectsEx, turning the loop into a 100% CPU spin.
                // 2. The previous outer catch(Exception){} silently bypassed Thread.Sleep when any
                //    exception occurred, causing tight spin loops with zero backoff.
                // 3. WaitToReadAsync().AsTask().GetAwaiter().GetResult() blocks the thread efficiently
                //    and wakes instantly when work arrives. No COM message pumping occurs during the
                //    block, but that's fine — we don't host COM objects or subscribe to _pptApp events,
                //    so no inbound COM messages need dispatching while idle. COM calls within work items
                //    pump messages internally via CoWaitForMultipleHandles.
                while (true)
                {
                    try
                    {
                        // Block until work is available, channel completes, or shutdown is requested.
                        if (!_workQueue.Reader.WaitToReadAsync(_shutdownCts.Token)
                                              .AsTask().GetAwaiter().GetResult())
                        {
                            // Channel completed (writer called Complete()) — exit gracefully
                            _logger.LogDebug("Channel completed, exiting message pump for {FileName}", Path.GetFileName(_presentationPath));
                            break;
                        }

                        // Drain all available work items before blocking again
                        while (_workQueue.Reader.TryRead(out var work))
                        {
                            try
                            {
                                work().GetAwaiter().GetResult();
                            }
                            catch (Exception)
                            {
                                // Individual work items may fail, but keep processing queue.
                                // The exception is already captured in the TaskCompletionSource.
                            }
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        // Shutdown requested via _shutdownCts.
                        // Drain any remaining work items so in-flight Execute() callers get their
                        // results/exceptions promptly instead of waiting for the 5-minute timeout.
                        // This is safe: _pptApp COM objects are still alive (cleaned up in the finally
                        // block below), and Writer.Complete() prevents new items from arriving.
                        while (_workQueue.Reader.TryRead(out var remainingWork))
                        {
                            try
                            {
                                remainingWork().GetAwaiter().GetResult();
                            }
                            catch (Exception)
                            {
                                // Already captured in TaskCompletionSource
                            }
                        }

                        _logger.LogDebug("Shutdown requested, exiting message pump for {FileName}", Path.GetFileName(_presentationPath));
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                started.TrySetException(ex);
            }
            finally
            {
                // Cleanup COM objects on STA thread exit
                _logger.LogDebug("STA thread cleanup starting for {FileName}", Path.GetFileName(_presentationPath));

                // For multi-Presentation batches, close all Presentations individually before quitting PowerPoint
                if (_presentations != null && _presentations.Count > 1)
                {
                    _logger.LogDebug("Closing {Count} Presentations", _presentations.Count);
                    foreach (var kvp in _presentations.ToList())
                    {
                        try
                        {
                            PowerPoint.Presentation? pres = kvp.Value;
                            // Suppress save-changes dialog
                            try { ((dynamic)pres).Saved = -1; } catch { }
                            pres.Close();
                            Marshal.ReleaseComObject(pres);
                            pres = null;
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to close Presentation {Path}", kvp.Key);
                        }
                    }
                    _presentations.Clear();

                    // Quit _pptApp after all Presentations closed
                    if (_powerPoint != null)
                    {
                        try
                        {
                            _logger.LogDebug("Quitting _pptApp application");
                            _powerPoint.Quit();
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to quit PowerPoint");
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(_powerPoint);
                            _powerPoint = null;
                        }
                    }
                }
                else
                {
                    // Single Presentation: use PptShutdownService for resilient shutdown
                    PptShutdownService.CloseAndQuit(_presentation, _powerPoint, false, _presentationPath, _logger);
                }

                _presentation = null;
                _powerPoint = null;
                _presentations = null;
                _context = null;

                OleMessageFilter.Revoke();
                _logger.LogDebug("STA thread cleanup completed for {FileName}", Path.GetFileName(_presentationPath));
            }
        })
        {
            IsBackground = true,
            Name = $"PptBatch-{Path.GetFileName(_presentationPath)}"
        };

        // CRITICAL: Set STA apartment state before starting thread
        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();

        // Wait for STA thread to initialize
        started.Task.GetAwaiter().GetResult();
    }

    public string PresentationPath => _presentationPath;

    public ILogger Logger => _logger;

    public int? PowerPointProcessId => _powerPointProcessId;

    public TimeSpan OperationTimeout => _operationTimeout;

    public bool IsPowerPointProcessAlive()
    {
        if (_disposed != 0) return false;
        if (!_powerPointProcessId.HasValue) return false;

        try
        {
            using var proc = System.Diagnostics.Process.GetProcessById(_powerPointProcessId.Value);
            return !proc.HasExited;
        }
        catch (ArgumentException)
        {
            // Process ID doesn't exist - process has terminated
            return false;
        }
    }

    public IReadOnlyDictionary<string, PowerPoint.Presentation> Presentations
    {
        get
        {
            ObjectDisposedException.ThrowIf(_disposed != 0, nameof(PptBatch));
            return _presentations ?? throw new InvalidOperationException("Presentations not initialized");
        }
    }

    public PowerPoint.Presentation GetPresentation(string filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, nameof(PptBatch));

        if (_presentations == null)
            throw new InvalidOperationException("Presentations not initialized");

        string normalizedPath = Path.GetFullPath(filePath);
        if (_presentations.TryGetValue(normalizedPath, out var Presentation))
        {
            return Presentation;
        }

        throw new KeyNotFoundException($"Presentation '{filePath}' is not open in this batch.");
    }

    /// <summary>
    /// Executes a void COM operation on the STA thread.
    /// Use this overload for operations that don't need to return values.
    /// All _pptApp COM operations are synchronous.
    /// </summary>
    public void Execute(
        Action<PptContext, CancellationToken> operation,
        CancellationToken cancellationToken = default)
    {
        // Delegate to generic Execute<T> with dummy return
        Execute((ctx, ct) =>
        {
            operation(ctx, ct);
            return 0;
        }, cancellationToken);
    }

    /// <summary>
    /// Executes a COM operation on the STA thread.
    /// All _pptApp COM operations are synchronous.
    /// </summary>
    public T Execute<T>(
        Func<PptContext, CancellationToken, T> operation,
        CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, nameof(PptBatch));

        // Fail fast if a previous operation timed out or was cancelled while the STA thread
        // was stuck in IDispatch.Invoke. The STA thread cannot process new work items until
        // the hung COM call returns (which may be never). Without this check, new callers
        // would queue work and block until their own timeout expires.
        if (_operationTimedOut)
        {
            throw new TimeoutException(
                $"A previous operation timed out or was cancelled for '{Path.GetFileName(_presentationPath)}'. " +
                "The _pptApp COM thread may be unresponsive. Please close this session and create a new one.");
        }

        // Check if _pptApp process is still alive before attempting operation
        if (!IsPowerPointProcessAlive())
        {
            _logger.LogError("_pptApp process is no longer running for Presentation {FileName}", Path.GetFileName(_presentationPath));
            throw new InvalidOperationException(
                $"_pptApp process is no longer running for Presentation '{Path.GetFileName(_presentationPath)}'. " +
                "The _pptApp application may have been closed manually or crashed. " +
                "Please close this session and create a new one.");
        }

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

        // Post operation to STA thread synchronously
        // RACE CONDITION NOTE: Dispose() may call Writer.Complete() between our _disposed check
        // above and this WriteAsync() call. ChannelClosedException means the session is shutting
        // down — convert to ObjectDisposedException for a clean caller experience.
        try
        {
            var writeTask = _workQueue.Writer.WriteAsync(() =>
            {
                try
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var result = operation(_context!, cancellationToken);
                    tcs.SetResult(result);
                }
                catch (OperationCanceledException oce)
                {
                    tcs.TrySetCanceled(oce.CancellationToken);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
                return Task.CompletedTask;
            }, cancellationToken);

            // ValueTask is completed synchronously in normal case
            if (writeTask.IsCompleted)
            {
                writeTask.GetAwaiter().GetResult();
            }
            else
            {
                // Fallback: should not normally occur with unbounded channel
                writeTask.AsTask().GetAwaiter().GetResult();
            }
        }
        catch (ChannelClosedException)
        {
            // Dispose() completed the channel between our _disposed check and WriteAsync.
            // The session is shutting down — report as disposed.
            throw new ObjectDisposedException(nameof(PptBatch),
                $"Session for '{Path.GetFileName(_presentationPath)}' was disposed while submitting an operation.");
        }

        // Wait for operation to complete with timeout.
        // When the caller provides a cancellation token (e.g., PowerQuery refresh with its own timeout),
        // respect it exclusively and don't layer the session _operationTimeout on top.
        // This prevents a double-cap where min(callerTimeout, sessionTimeout) is always the shorter one —
        // which caused heavy Power Query refreshes (~8+ min) to always fail against the 5-min default.
        try
        {
            if (cancellationToken.CanBeCanceled)
            {
                // Caller controls the timeout — use their token exclusively
                return tcs.Task.WaitAsync(cancellationToken).GetAwaiter().GetResult();
            }
            else
            {
                // No caller timeout — apply session-level operation timeout as safety net
                using var timeoutCts = new CancellationTokenSource(_operationTimeout);
                return tcs.Task.WaitAsync(timeoutCts.Token).GetAwaiter().GetResult();
            }
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            // Session timeout occurred (not caller cancellation) — only happens in the else branch
            _logger.LogError("Operation timed out after {Timeout} for {FileName}", _operationTimeout, Path.GetFileName(_presentationPath));
            _operationTimedOut = true; // Mark timeout for aggressive cleanup during disposal
            throw new TimeoutException(
                $"_pptApp operation timed out after {_operationTimeout.TotalSeconds} seconds for '{Path.GetFileName(_presentationPath)}'. " +
                "_pptApp may be unresponsive or the operation is taking longer than expected. " +
                "Consider increasing timeoutSeconds when opening the session.");
        }
        catch (OperationCanceledException)
        {
            _logger.LogDebug("Operation cancelled or timed out for {FileName}", Path.GetFileName(_presentationPath));
            _operationTimedOut = true; // STA thread may still be blocked — session is unusable
            throw;
        }
    }

    public void Save(CancellationToken cancellationToken = default)
    {
        Execute((ctx, ct) =>
        {
            PptShutdownService.SavePresentationWithTimeout(
                _presentation!,
                Path.GetFileName(_presentationPath),
                _logger,
                ct);
            return 0;
        }, cancellationToken);
    }

    public void Dispose()
    {
        var callingThread = Environment.CurrentManagedThreadId;

        // Use Interlocked.CompareExchange for thread-safe disposal check
        // Returns 0 if exchange succeeded (was not disposed), 1 if already disposed
        if (Interlocked.CompareExchange(ref _disposed, 1, 0) != 0)
        {
            _logger.LogDebug("[Thread {CallingThread}] Dispose skipped - already disposed for {FileName}", callingThread, Path.GetFileName(_presentationPath));
            return; // Already disposed
        }

        _logger.LogDebug("[Thread {CallingThread}] Dispose starting for {FileName}", callingThread, Path.GetFileName(_presentationPath));

        // Cancel the shutdown token FIRST to wake up the message pump
        _logger.LogDebug("[Thread {CallingThread}] Cancelling shutdown token for {FileName}", callingThread, Path.GetFileName(_presentationPath));
        _shutdownCts.Cancel();

        // Then complete the work queue
        _logger.LogDebug("[Thread {CallingThread}] Completing work queue for {FileName}", callingThread, Path.GetFileName(_presentationPath));
        _workQueue.Writer.Complete();

        _logger.LogDebug("[Thread {CallingThread}] Waiting for STA thread (Id={STAThread}) to exit for {FileName}", callingThread, _staThread?.ManagedThreadId ?? -1, Path.GetFileName(_presentationPath));

        // When operation timed out, the STA thread is stuck in IDispatch.Invoke (unmanaged COM call
        // that cannot be cancelled). Kill the _pptApp process FIRST to unblock the STA thread, then wait.
        if (_operationTimedOut && _powerPointProcessId.HasValue && _staThread != null && _staThread.IsAlive)
        {
            _logger.LogWarning(
                "[Thread {CallingThread}] Operation timed out — force-killing _pptApp process {ProcessId} BEFORE waiting for STA thread to unblock IDispatch.Invoke for {FileName}",
                callingThread, _powerPointProcessId.Value, Path.GetFileName(_presentationPath));
            try
            {
                using var pptProcess = System.Diagnostics.Process.GetProcessById(_powerPointProcessId.Value);
                if (!pptProcess.HasExited)
                {
                    pptProcess.Kill();
                    pptProcess.WaitForExit(5000);
                    _logger.LogInformation(
                        "[Thread {CallingThread}] Force-killed _pptApp process {ProcessId} (pre-emptive, before STA join)",
                        callingThread, _powerPointProcessId.Value);
                }
            }
            catch (ArgumentException)
            {
                _logger.LogDebug("[Thread {CallingThread}] _pptApp process {ProcessId} already exited", callingThread, _powerPointProcessId.Value);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[Thread {CallingThread}] Failed to force-kill _pptApp process {ProcessId}", callingThread, _powerPointProcessId.Value);
            }
        }

        // Wait for STA thread to finish cleanup (with timeout)
        if (_staThread != null && _staThread.IsAlive)
        {
            // Use shorter timeout if operation timed out (_pptApp is likely hung / already killed above)
            var joinTimeout = _operationTimedOut
                ? TimeSpan.FromSeconds(10)  // Aggressive: 10 seconds when operation timed out
                : ComInteropConstants.StaThreadJoinTimeout;  // Normal: 45 seconds

            var reasonSuffix = _operationTimedOut ? " (operation timed out - aggressive cleanup)" : "";
            _logger.LogDebug(
                "[Thread {CallingThread}] Calling Join() with {Timeout} timeout on STA={STAThread}, file={FileName}{Reason}",
                callingThread, joinTimeout, _staThread.ManagedThreadId, Path.GetFileName(_presentationPath), reasonSuffix);

            // CRITICAL: StaThreadJoinTimeout >= PowerPointQuitTimeout + margin (currently 45 seconds total).
            // The join must wait at least as long as CloseAndQuit() can take, otherwise Dispose() returns
            // before _pptApp has finished closing, causing "file still open" issues in subsequent operations.
            if (!_staThread.Join(joinTimeout))
            {
                // STA thread didn't exit - _pptApp cleanup is severely stuck
                var reasonForError = _operationTimedOut ? " (operation previously timed out)" : "";
                _logger.LogError(
                    "[Thread {CallingThread}] STA thread (Id={STAThread}) did NOT exit within {Timeout} for {FileName}. " +
                    "_pptApp cleanup is severely stuck{Reason}. Attempting force-kill.",
                    callingThread, _staThread.ManagedThreadId, joinTimeout, Path.GetFileName(_presentationPath), reasonForError);

                // Force-kill the hung _pptApp process
                if (_powerPointProcessId.HasValue)
                {
                    try
                    {
                        using var pptProcess = System.Diagnostics.Process.GetProcessById(_powerPointProcessId.Value);
                        _logger.LogWarning(
                            "[Thread {CallingThread}] Force-killing _pptApp process {ProcessId} for {FileName}",
                            callingThread, _powerPointProcessId.Value, Path.GetFileName(_presentationPath));

                        pptProcess.Kill();
                        pptProcess.WaitForExit(5000); // Wait up to 5 seconds for process to die

                        _logger.LogInformation(
                            "[Thread {CallingThread}] Successfully force-killed _pptApp process {ProcessId}",
                            callingThread, _powerPointProcessId.Value);

                        // Now wait briefly for STA thread to exit after process killed
                        if (_staThread.Join(TimeSpan.FromSeconds(5)))
                        {
                            _logger.LogDebug("[Thread {CallingThread}] STA thread exited after force-kill", callingThread);
                        }
                        else
                        {
                            _logger.LogWarning(
                                "[Thread {CallingThread}] STA thread still stuck even after force-kill. Thread leak.",
                                callingThread);
                        }
                    }
                    catch (ArgumentException)
                    {
                        _logger.LogWarning(
                            "[Thread {CallingThread}] _pptApp process {ProcessId} not found (already exited?)",
                            callingThread, _powerPointProcessId.Value);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex,
                            "[Thread {CallingThread}] Failed to force-kill _pptApp process {ProcessId}",
                            callingThread, _powerPointProcessId.Value);
                    }
                }
                else
                {
                    _logger.LogError(
                        "[Thread {CallingThread}] No _pptApp process ID captured - cannot force-kill. Process will leak.",
                        callingThread);
                }
            }
        }
        else
        {
            _logger.LogDebug("[Thread {CallingThread}] STA thread was null or not alive for {FileName}", callingThread, Path.GetFileName(_presentationPath));
        }

        // Wait for _pptApp process to fully terminate to prevent CO_E_SERVER_EXEC_FAILURE
        // on subsequent Activator.CreateInstance calls. PowerPoint.Quit() + COM release doesn't
        // guarantee the PowerPoint.EXE process has exited — rapid create/destroy cycles can fail.
        if (_powerPointProcessId.HasValue)
        {
            try
            {
                using var pptProcess = System.Diagnostics.Process.GetProcessById(_powerPointProcessId.Value);
                if (!pptProcess.HasExited)
                {
                    _logger.LogDebug(
                        "[Thread {CallingThread}] Waiting for _pptApp process {ProcessId} to exit for {FileName}",
                        callingThread, _powerPointProcessId.Value, Path.GetFileName(_presentationPath));

                    if (!pptProcess.WaitForExit(5000))
                    {
                        _logger.LogWarning(
                            "[Thread {CallingThread}] _pptApp process {ProcessId} did not exit within 5s for {FileName}. Force-killing to prevent zombie accumulation.",
                            callingThread, _powerPointProcessId.Value, Path.GetFileName(_presentationPath));

                        // Force-kill: _pptApp was already told to Quit() and COM refs were released.
                        // A process still running after 5s is hung and will leak desktop resources.
                        try
                        {
                            pptProcess.Kill();
                            pptProcess.WaitForExit(3000);
                            _logger.LogInformation(
                                "[Thread {CallingThread}] Force-killed lingering _pptApp process {ProcessId} for {FileName}",
                                callingThread, _powerPointProcessId.Value, Path.GetFileName(_presentationPath));
                        }
                        catch (Exception killEx)
                        {
                            _logger.LogWarning(killEx,
                                "[Thread {CallingThread}] Failed to force-kill _pptApp process {ProcessId}",
                                callingThread, _powerPointProcessId.Value);
                        }
                    }
                }
            }
            catch (ArgumentException)
            {
                // Process already terminated — this is the expected fast path
            }
            catch (InvalidOperationException)
            {
                // Process object is not associated with a running process
            }
        }

        // Dispose cancellation token source
        _logger.LogDebug("[Thread {CallingThread}] Disposing CancellationTokenSource for {FileName}", callingThread, Path.GetFileName(_presentationPath));
        _shutdownCts.Dispose();

        _logger.LogDebug("[Thread {CallingThread}] Dispose COMPLETED for {FileName}", callingThread, Path.GetFileName(_presentationPath));
    }
}


