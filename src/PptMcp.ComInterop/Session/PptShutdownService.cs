using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptMcp.ComInterop.Session;

/// <summary>
/// Centralized service for PowerPoint presentation close and application quit operations.
/// Shutdown terminates PowerPoint by releasing the last COM reference rather than calling Quit().
/// </summary>
public static class PptShutdownService
{
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
                    // Step 3: Release presentation COM reference.
                    // FinalReleaseComObject, not ReleaseComObject: the same COM object can be
                    // marshalled to us more than once, and each marshalling increments the RCW
                    // reference count. Decrementing by one would leave the proxy alive.
                    Marshal.FinalReleaseComObject(presentation);
                    presentation = null;
                }
            }

            // Step 4: Release the PowerPoint application reference.
            //
            // This deliberately does NOT call Application.Quit(). Quit() returns almost
            // immediately but puts PowerPoint into a shutdown state in which it stops
            // servicing COM calls, so the subsequent Release of our application proxy
            // deadlocks until COM's own call timeout expires. Measured on a plain STA
            // thread with no wrapper involved (issue #148):
            //
            //   Quit() then Release(app) : Quit 15ms, Release 60042ms, process exits after
            //   Release(app) only        :            Release    18ms, process exits after
            //
            // Dropping the last external reference is what actually terminates PowerPoint,
            // and it is what Quit() was ultimately waiting for. Releasing directly is both
            // faster and safer: if this process attached to a PowerPoint instance the user
            // already had open, Quit() would have torn down their session.
            //
            // If PowerPoint fails to exit anyway, the caller force-kills the captured
            // process id as a last resort.
            if (powerPoint != null)
            {
                Marshal.FinalReleaseComObject(powerPoint);
                powerPoint = null;

                // Release any straggler proxies still held by RCWs awaiting finalization,
                // otherwise PowerPoint keeps running because a reference remains outstanding.
                GC.Collect();
                GC.WaitForPendingFinalizers();

                logger.LogDebug("PowerPoint application reference released for {FileName} after {Elapsed}ms",
                    fileName, stopwatch.ElapsedMilliseconds);
            }
        }
        finally
        {
            logger.LogDebug("PowerPoint shutdown sequence completed for {FileName} in {Elapsed}ms",
                fileName, stopwatch.ElapsedMilliseconds);
        }
    }
}
