namespace PptMcp.ComInterop;

/// <summary>
/// Constants for PowerPoint COM interop operations.
/// </summary>
public static class ComInteropConstants
{
    #region Timeouts

    /// <summary>
    /// Budget for releasing PowerPoint's COM references during shutdown (30 seconds).
    /// The shutdown sequence closes the presentation and releases every proxy; this timeout
    /// catches a COM server that stops responding part-way through.
    /// </summary>
    public static readonly TimeSpan PowerPointShutdownTimeout = TimeSpan.FromSeconds(30);

    /// <summary>
    /// Timeout for STA thread join after shutdown.
    /// CRITICAL: Must be >= PowerPointShutdownTimeout so Dispose() waits for the shutdown
    /// sequence to complete. Set to PowerPointShutdownTimeout + 15s margin for presentation
    /// close and COM cleanup.
    /// </summary>
    public static readonly TimeSpan StaThreadJoinTimeout = PowerPointShutdownTimeout + TimeSpan.FromSeconds(15);

    /// <summary>
    /// Timeout for save operations (5 minutes).
    /// Large presentations may take longer to save.
    /// </summary>
    public static readonly TimeSpan SaveOperationTimeout = TimeSpan.FromMinutes(5);

    /// <summary>
    /// Default timeout for individual PowerPoint operations (5 minutes).
    /// Most operations complete in under 30 seconds, but this provides buffer for slow machines.
    /// Can be overridden when creating a session via timeoutSeconds parameter.
    /// </summary>
    public static readonly TimeSpan DefaultOperationTimeout = TimeSpan.FromMinutes(5);

    /// <summary>
    /// Maximum wait time for session creation file lock acquisition (5 seconds).
    /// </summary>
    public static readonly TimeSpan SessionFileLockTimeout = TimeSpan.FromSeconds(5);

    #endregion

    #region Sleep Intervals

    /// <summary>
    /// Delay between file lock acquisition retries (100ms).
    /// </summary>
    public const int FileLockRetryDelayMs = 100;

    /// <summary>
    /// Delay between session lock acquisition retries (200ms).
    /// </summary>
    public const int SessionLockRetryDelayMs = 200;

    #endregion

    #region PowerPoint File Formats

    /// <summary>
    /// PowerPoint Open XML Presentation format code (.pptx).
    /// PpSaveAsFileType.ppSaveAsOpenXMLPresentation = 24
    /// </summary>
    public const int PpSaveAsOpenXMLPresentation = 24;

    /// <summary>
    /// PowerPoint Open XML Macro-Enabled Presentation format code (.pptm).
    /// PpSaveAsFileType.ppSaveAsOpenXMLPresentationMacroEnabled = 25
    /// </summary>
    public const int PpSaveAsOpenXMLPresentationMacroEnabled = 25;

    /// <summary>
    /// PowerPoint default format code.
    /// PpSaveAsFileType.ppSaveAsDefault = 11
    /// </summary>
    public const int PpSaveAsDefault = 11;

    #endregion
}


