using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Ppt = Microsoft.Office.Interop.PowerPoint;

namespace PptMcp.ComInterop.Session;

/// <summary>
/// Main entry point for PowerPoint COM interop operations using batch pattern.
/// All operations execute on dedicated STA threads with proper COM cleanup.
/// </summary>
public static class PptSession
{
    /// <summary>
    /// Global lock to serialize file creation operations.
    /// Prevents resource exhaustion from parallel CreateNew() calls.
    /// Each CreateNew() spawns a temporary PowerPoint instance - must be sequential.
    /// </summary>
    private static readonly SemaphoreSlim _createFileLock = new(1, 1);

    /// <summary>
    /// Begins a batch of PowerPoint operations against one or more Presentation instances.
    /// The PowerPoint instance remains open until the batch is disposed, enabling multiple operations
    /// without incurring PowerPoint startup/shutdown overhead.
    /// </summary>
    /// <param name="filePaths">Paths to PowerPoint files. First file is the primary Presentation.</param>
    /// <returns>IPptBatch for executing multiple operations</returns>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static IPptBatch BeginBatch(params string[] filePaths)
        => BeginBatch(show: false, operationTimeout: null, filePaths);

    /// <summary>
    /// Begins a batch of PowerPoint operations against one or more Presentation instances with optional UI visibility.
    /// </summary>
    /// <param name="show">Whether to show the PowerPoint window (default: false for background automation).</param>
    /// <param name="operationTimeout">Maximum time for any single operation (default: 5 minutes).</param>
    /// <param name="filePaths">Paths to PowerPoint files. First file is the primary Presentation.</param>
    /// <returns>IPptBatch for executing multiple operations</returns>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static IPptBatch BeginBatch(
        bool show,
        TimeSpan? operationTimeout,
        params string[] filePaths)
    {
        if (filePaths == null || filePaths.Length == 0)
            throw new ArgumentException("At least one file path is required", nameof(filePaths));

        string[] fullPaths = new string[filePaths.Length];
        for (int i = 0; i < filePaths.Length; i++)
        {
            string fullPath = Path.GetFullPath(filePaths[i]);

            if (!File.Exists(fullPath))
            {
                throw new FileNotFoundException($"PowerPoint file not found: {fullPath}. To create a new file, use the 'create' action instead of 'open'.", fullPath);
            }

            string extension = Path.GetExtension(fullPath).ToLowerInvariant();
            if (extension is not (".pptx" or ".pptm" or ".ppt"))
            {
                throw new ArgumentException($"Invalid file extension '{extension}'. Only PowerPoint files (.pptx, .pptm, .ppt) are supported.");
            }

            fullPaths[i] = fullPath;
        }

        return new PptBatch(fullPaths, logger: null, show: show, operationTimeout: operationTimeout);
    }

    /// <summary>
    /// Creates a new PowerPoint Presentation at the specified path with a synchronous COM operation.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static T CreateNew<T>(
        string filePath,
        bool isMacroEnabled,
        Func<PptContext, CancellationToken, T> operation,
        CancellationToken cancellationToken = default)
    {
        if (!_createFileLock.Wait(TimeSpan.FromMinutes(2), cancellationToken))
        {
            throw new TimeoutException("Timed out waiting for file creation lock. Another CreateNew operation may be stuck.");
        }
        try
        {
            string fullPath = Path.GetFullPath(filePath);

            if (fullPath.Length > 218)
            {
                throw new PathTooLongException(
                    $"File path exceeds PowerPoint's maximum length (~218 characters): {fullPath.Length} characters");
            }

            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            CreatePresentationOnStaThread(fullPath, isMacroEnabled, cancellationToken);

            using var batch = BeginBatch(fullPath);
            var result = batch.Execute(operation, cancellationToken);
            return result;
        }
        finally
        {
            _createFileLock.Release();
        }
    }

    private static void CreatePresentationOnStaThread(string fullPath, bool isMacroEnabled, CancellationToken cancellationToken)
    {
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        var thread = new Thread(() =>
        {
            Ppt.Application? pptApp = null;
            Ppt.Presentation? presentation = null;

            try
            {
                OleMessageFilter.Register();

                var pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                if (pptType == null)
                {
                    throw new InvalidOperationException("PowerPoint is not installed or not properly registered.");
                }

#pragma warning disable IL2072
                pptApp = (Ppt.Application)Activator.CreateInstance(pptType)!;
#pragma warning restore IL2072

                // PowerPoint COM does NOT allow Visible = msoFalse (0).
                // Unlike Excel, PowerPoint throws "Hiding the application window is not allowed."
                // Always set Visible = msoTrue, minimize window instead.
                ((dynamic)pptApp).Visible = -1; // msoTrue — required by PowerPoint COM
                ((dynamic)pptApp).DisplayAlerts = 0; // ppAlertsNone
                ((dynamic)pptApp).WindowState = 2; // ppWindowMinimized

                presentation = ((dynamic)pptApp).Presentations.Add();

                int fileFormat = isMacroEnabled
                    ? ComInteropConstants.PpSaveAsOpenXMLPresentationMacroEnabled
                    : ComInteropConstants.PpSaveAsOpenXMLPresentation;
                ((dynamic)presentation).SaveAs(fullPath, fileFormat);

                completion.SetResult();
            }
            catch (Exception ex)
            {
                completion.TrySetException(ex);
            }
            finally
            {
                try
                {
                    presentation?.Close();
                }
                catch { }

                if (pptApp != null)
                {
                    try { pptApp.Quit(); } catch { }
                    try { Marshal.ReleaseComObject(pptApp); } catch { }
                }
                if (presentation != null)
                {
                    try { Marshal.ReleaseComObject(presentation); } catch { }
                }

                OleMessageFilter.Revoke();
            }
        })
        {
            IsBackground = true,
            Name = $"PptCreate-{Path.GetFileName(fullPath)}"
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        if (!completion.Task.Wait(TimeSpan.FromSeconds(30), cancellationToken))
        {
            throw new TimeoutException($"File creation timed out for '{Path.GetFileName(fullPath)}'. PowerPoint may be unresponsive.");
        }

        thread.Join(TimeSpan.FromSeconds(10));
    }
}




