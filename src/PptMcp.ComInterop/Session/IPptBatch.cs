namespace PptMcp.ComInterop.Session;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

/// <summary>
/// Represents a batch of PowerPoint operations that share a single PowerPoint instance.
/// Implements IDisposable to ensure proper COM cleanup.
/// </summary>
/// <remarks>
/// Use this interface via PptSession.BeginBatch() for multi-operation workflows.
/// The batch keeps PowerPoint and the presentation open until disposed, enabling efficient
/// execution of multiple commands without repeated PowerPoint startup/shutdown overhead.
///
/// <para><b>Lifecycle:</b></para>
/// <list type="bullet">
/// <item>Created via PptSession.BeginBatch(filePath)</item>
/// <item>Operations executed via Execute()</item>
/// <item>Optional explicit save via Save()</item>
/// <item>Disposed via Dispose() or "using" pattern</item>
/// </list>
///
/// <para><b>Example:</b></para>
/// <code>
/// using var batch = PptSession.BeginBatch("presentation.pptx");
///
/// // Execute COM operations
/// batch.Execute((ctx, ct) => {
///     ctx.Presentation.Slides.Add(1, 1);
///     return 0;
/// });
///
/// // Get content from PowerPoint
/// var count = batch.Execute((ctx, ct) => {
///     return ctx.Presentation.Slides.Count;
/// });
///
/// // Explicit save
/// batch.Save();
/// </code>
/// </remarks>
public interface IPptBatch : IDisposable
{
    /// <summary>
    /// Gets the path to the PowerPoint presentation this batch operates on.
    /// For multi-presentation batches, this is the primary (first) presentation.
    /// </summary>
    string PresentationPath { get; }

    /// <summary>
    /// Gets the logger instance for diagnostic output.
    /// Returns NullLogger if no logger was provided during construction.
    /// </summary>
    Microsoft.Extensions.Logging.ILogger Logger { get; }

    /// <summary>
    /// Gets all presentations currently open in this batch, keyed by normalized file path.
    /// For single-presentation batches, contains one entry.
    /// For multi-presentation batches (cross-presentation operations), contains all open presentations.
    /// </summary>
    IReadOnlyDictionary<string, PowerPoint.Presentation> Presentations { get; }

    /// <summary>
    /// Gets the COM Presentation object for a specific file path.
    /// </summary>
    /// <param name="filePath">Path to the presentation (will be normalized)</param>
    /// <returns>PowerPoint.Presentation COM object</returns>
    /// <exception cref="KeyNotFoundException">Presentation not found in this batch</exception>
    PowerPoint.Presentation GetPresentation(string filePath);

    /// <summary>
    /// Executes a void COM operation within this batch.
    /// The operation receives a PptContext with access to the PowerPoint app and presentation.
    /// Use this overload for void operations that don't need to return values.
    /// All PowerPoint COM operations are synchronous - file I/O should be handled outside the batch.
    /// </summary>
    /// <param name="operation">COM operation to execute</param>
    /// <param name="cancellationToken">Optional cancellation token for caller-controlled cancellation</param>
    /// <exception cref="ObjectDisposedException">Batch has already been disposed</exception>
    /// <exception cref="InvalidOperationException">PowerPoint COM error occurred</exception>
    /// <exception cref="OperationCanceledException">Operation was cancelled via cancellationToken</exception>
    void Execute(
        Action<PptContext, CancellationToken> operation,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Executes a COM operation within this batch.
    /// The operation receives a PptContext with access to the PowerPoint app and presentation.
    /// All PowerPoint COM operations are synchronous - file I/O should be handled outside the batch.
    /// </summary>
    /// <typeparam name="T">Return type of the operation</typeparam>
    /// <param name="operation">COM operation to execute</param>
    /// <param name="cancellationToken">Optional cancellation token for caller-controlled cancellation</param>
    /// <returns>Result of the operation</returns>
    /// <exception cref="ObjectDisposedException">Batch has already been disposed</exception>
    /// <exception cref="InvalidOperationException">PowerPoint COM error occurred</exception>
    /// <exception cref="OperationCanceledException">Operation was cancelled via cancellationToken</exception>
    T Execute<T>(
        Func<PptContext, CancellationToken, T> operation,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Saves changes to the presentation.
    /// This is an explicit save - changes are NOT automatically saved on dispose.
    /// </summary>
    /// <param name="cancellationToken">Optional cancellation token for caller-controlled cancellation</param>
    /// <exception cref="ObjectDisposedException">Batch has already been disposed</exception>
    /// <exception cref="InvalidOperationException">Save failed (e.g., file is read-only)</exception>
    /// <exception cref="OperationCanceledException">Save operation was cancelled via cancellationToken</exception>
    void Save(CancellationToken cancellationToken = default);

    /// <summary>
    /// Checks if the underlying PowerPoint process is still alive.
    /// </summary>
    /// <returns>
    /// True if PowerPoint process exists and hasn't exited.
    /// False if process has crashed, was killed, or process ID wasn't captured.
    /// </returns>
    /// <remarks>
    /// Use this to detect dead PowerPoint processes before attempting operations.
    /// If this returns false, the session should be closed and recreated.
    /// </remarks>
    bool IsPowerPointProcessAlive();

    /// <summary>
    /// Gets the PowerPoint process ID, if captured.
    /// </summary>
    /// <returns>Process ID, or null if not captured during startup.</returns>
    int? PowerPointProcessId { get; }

    /// <summary>
    /// Gets the operation timeout for this batch.
    /// All Execute() calls will timeout after this duration.
    /// Default is 5 minutes (from ComInteropConstants.DefaultOperationTimeout).
    /// </summary>
    TimeSpan OperationTimeout { get; }

}




