using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Polly;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptMcp.ComInterop.Session;

/// <summary>
/// Centralized service for PowerPoint presentation close and application quit operations.
/// Implements resilient shutdown with exponential backoff for COM busy conditions.
/// </summary>
public static class PptShutdownService
{
    private static readonly ResiliencePipeline _quitPipeline = ResiliencePipelines.CreatePowerPointQuitPipeline();

    /// <summary>
    /// Saves a PowerPoint presentation on the calling STA thread.
    /// Must be called from within <c>PptBatch.Execute()</c> so the Save() COM call
    /// runs on the correct STA thread. Timeout protection is provided by the surrounding
    /// <c>PptBatch.Execute()</c> operation timeout and the Dispose() force-kill chain.
    /// </summary>
    /// <param name="presentation">PowerPoint presentation COM object to save</param>
    /// <param name="fileName">File name for diagnostic messages (optional)</param>
    /// <param name="logger">Logger for diagnostic output (optional)</param>
    /// <param name="cancellationToken">Cancellation token checked before Save() is invoked</param>
    /// <exception cref="OperationCanceledException">Cancellation was requested before save started</exception>
    /// <exception cref="COMException">Save failed due to COM error</exception>
    /// <exception cref="InvalidOperationException">Save failed due to unexpected error</exception>
    public static void SavePresentationWithTimeout(
        PowerPoint.Presentation presentation,
        string? fileName = null,
        ILogger? logger = null,
        CancellationToken cancellationToken = default)
    {
        logger ??= NullLogger.Instance;
        fileName ??= "unknown";

        // Honour any cancellation request before we start the potentially slow COM call
        cancellationToken.ThrowIfCancellationRequested();

        logger.LogDebug("Saving presentation {FileName}", fileName);

        try
        {
            presentation.Save();

            logger.LogDebug("Presentation {FileName} saved successfully", fileName);
        }
        catch (COMException ex)
        {
            string errorMessage = ex.HResult switch
            {
                unchecked((int)0x800A03EC) =>
                    $"Cannot save '{fileName}'. " +
                    "The file may be read-only, locked by another process, or the path may not exist.",
                unchecked((int)0x800AC472) =>
                    $"Cannot save '{fileName}'. " +
                    "The file is locked for editing by another user or process.",
                _ => $"Failed to save presentation '{fileName}': {ex.Message}"
            };

            logger.LogError(ex, "Save failed for {FileName} (HResult: 0x{HResult:X8})", fileName, ex.HResult);
            throw new InvalidOperationException(errorMessage, ex);
        }
        // All other exceptions propagate; no generic catch block.
    }

    /// <summary>
    /// Closes a presentation and quits the PowerPoint application with resilient retry logic.
    /// Handles save semantics, presentation close, COM object release, and resilient Quit with backoff.
    /// </summary>
    /// <param name="presentation">PowerPoint presentation COM object (can be null)</param>
    /// <param name="powerPoint">PowerPoint application COM object (can be null)</param>
    /// <param name="save">True to save before closing, false to discard changes</param>
    /// <param name="filePath">File path for diagnostic logging (optional)</param>
    /// <param name="logger">Logger for diagnostic output (optional)</param>
    /// <remarks>
    /// <para><b>Shutdown Order:</b></para>
    /// <list type="number">
    /// <item>If save=true: Call presentation.Save()</item>
    /// <item>Close presentation with Close() - discards unsaved changes if save=false</item>
    /// <item>Release presentation COM reference</item>
    /// <item>Quit PowerPoint application with exponential backoff retry (6 attempts, 200ms base delay)</item>
    /// <item>Release PowerPoint COM reference</item>
    /// </list>
    /// <para><b>Resilience:</b> Retries Quit() on COM busy errors (RPC_E_SERVERCALL_RETRYLATER, RPC_E_CALL_REJECTED)</para>
    /// </remarks>
    public static void CloseAndQuit(
        PowerPoint.Presentation? presentation,
        PowerPoint.Application? powerPoint,
        bool save,
        string? filePath = null,
        ILogger? logger = null)
    {
        logger ??= NullLogger.Instance;
        string fileName = string.IsNullOrEmpty(filePath) ? "unknown" : Path.GetFileName(filePath);

        var stopwatch = Stopwatch.StartNew();

        try
        {
            // Step 1: Explicit save if requested (before Close call)
            if (save && presentation != null)
            {
                SavePresentationWithTimeout(presentation, fileName, logger);
            }

            // Step 2: Close presentation
            if (presentation != null)
            {
                try
                {
                    logger.LogDebug("Closing presentation {FileName} (save={Save})", fileName, save);
                    // Mark as "already saved" to suppress the save-changes dialog
                    // PowerPoint COM shows a modal dialog on Close() if there are unsaved changes,
                    // even with DisplayAlerts=ppAlertsNone. Setting Saved=true prevents this.
                    if (!save)
                    {
                        try { ((dynamic)presentation).Saved = -1; } // msoTrue
                        catch { /* best effort */ }
                    }
                    presentation.Close();
                    logger.LogDebug("Presentation {FileName} closed successfully", fileName);
                }
                catch (COMException ex)
                {
                    logger.LogWarning(ex,
                        "Failed to close presentation {FileName} (HResult: 0x{HResult:X8}) - continuing with cleanup",
                        fileName, ex.HResult);
                }
                catch (MissingMemberException ex)
                {
                    // COM proxy already disconnected (RPC_E_DISCONNECTED / 0x80010108)
                    logger.LogWarning(ex,
                        "Presentation COM proxy was disconnected while calling Close for {FileName} - continuing with cleanup",
                        fileName);
                }
                finally
                {
                    // Step 3: Release presentation COM reference
                    Marshal.ReleaseComObject(presentation);
                    presentation = null;
                }
            }

            // Step 4: Quit PowerPoint application with resilient retry + overall timeout
            if (powerPoint != null)
            {
                int attemptNumber = 0;
                Exception? lastException = null;

                // Outer timeout catches truly hung PowerPoint (modal dialogs, deadlocks)
                using var quitTimeout = new CancellationTokenSource(ComInteropConstants.PowerPointQuitTimeout);

                try
                {
                    logger.LogDebug("Attempting to quit PowerPoint for {FileName} with resilient retry ({Timeout} timeout)", fileName, ComInteropConstants.PowerPointQuitTimeout);

                    // Inner retry pipeline handles transient COM busy errors within the timeout
                    _quitPipeline.Execute(cancellationToken =>
                    {
                        attemptNumber++;
                        try
                        {
                            logger.LogDebug("Quit attempt {Attempt} for {FileName}", attemptNumber, fileName);
                            powerPoint.Quit();
                            logger.LogDebug("Quit attempt {Attempt} succeeded for {FileName}", attemptNumber, fileName);
                        }
                        catch (COMException ex)
                        {
                            lastException = ex;
                            logger.LogWarning(ex,
                                "Quit attempt {Attempt} failed for {FileName} (HResult: 0x{HResult:X8})",
                                attemptNumber, fileName, ex.HResult);
                            throw; // Let pipeline decide if retry
                        }
                    }, quitTimeout.Token);

                    logger.LogInformation("PowerPoint quit succeeded for {FileName} after {Attempts} attempt(s) in {Elapsed}ms",
                        fileName, attemptNumber, stopwatch.ElapsedMilliseconds);
                }
                catch (OperationCanceledException) when (quitTimeout.Token.IsCancellationRequested)
                {
                    // Overall timeout reached - PowerPoint is truly hung
                    logger.LogError(
                        "PowerPoint quit TIMED OUT after {Timeout} for {FileName} (Attempts: {Attempts}). " +
                        "PowerPoint is likely hung (modal dialog or deadlock). Proceeding with forced COM cleanup.",
                        ComInteropConstants.PowerPointQuitTimeout, fileName, attemptNumber);
                    lastException = new TimeoutException($"PowerPoint.Quit() timed out after {ComInteropConstants.PowerPointQuitTimeout} for {fileName}");
                }
                catch (COMException ex) when (ex.HResult == ResiliencePipelines.RPC_E_CALL_FAILED)
                {
                    // Fatal RPC connection failure - PowerPoint is unreachable
                    logger.LogError(ex,
                        "PowerPoint RPC connection FAILED (0x800706BE) for {FileName}. " +
                        "PowerPoint is unreachable - this is a FATAL error that cannot be retried. " +
                        "Proceeding with forced COM cleanup. PowerPoint process should be force-killed by caller.",
                        fileName);
                    lastException = ex;
                }
                catch (COMException ex)
                {
                    // All retry attempts exhausted or non-retriable error
                    logger.LogError(ex,
                        "PowerPoint quit failed for {FileName} after {Attempts} attempt(s) (HResult: 0x{HResult:X8}, Elapsed: {Elapsed}ms) - proceeding with COM cleanup",
                        fileName, attemptNumber, ex.HResult, stopwatch.ElapsedMilliseconds);
                    lastException = ex;
                }
                catch (MissingMemberException ex)
                {
                    logger.LogWarning(ex,
                        "PowerPoint COM proxy was disconnected while calling Quit for {FileName} - proceeding with COM cleanup",
                        fileName);
                    lastException = ex;
                }
                finally
                {
                    // Step 5: Release PowerPoint COM reference (even if Quit failed/timed out)
                    Marshal.ReleaseComObject(powerPoint);
                    powerPoint = null;
                }

                // Additional diagnostic if quit failed
                if (lastException != null)
                {
                    logger.LogWarning(
                        "PowerPoint quit unsuccessful for {FileName} (Elapsed: {Elapsed}s, Type: {ExceptionType}). " +
                        "COM cleanup completed. Process may leak if PowerPoint remains hung.",
                        fileName, stopwatch.Elapsed.TotalSeconds, lastException.GetType().Name);
                }
            }
        }
        finally
        {
            logger.LogDebug("PowerPoint shutdown sequence completed for {FileName} in {Elapsed}ms",
                fileName, stopwatch.ElapsedMilliseconds);
        }
    }
}


