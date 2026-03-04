using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Window;

/// <summary>
/// PowerPoint window management: get info, minimize, restore, maximize.
/// </summary>
[ServiceCategory("window")]
[McpTool("window", Title = "Window Operations", Destructive = false, Category = "window")]
public interface IWindowCommands
{
    /// <summary>
    /// Get current window information (state, position, size).
    /// </summary>
    [ServiceAction("get-info")]
    WindowInfoResult GetInfo(IPptBatch batch);

    /// <summary>
    /// Minimize the PowerPoint window.
    /// </summary>
    [ServiceAction("minimize")]
    OperationResult Minimize(IPptBatch batch);

    /// <summary>
    /// Restore the PowerPoint window to normal size.
    /// </summary>
    [ServiceAction("restore")]
    OperationResult Restore(IPptBatch batch);

    /// <summary>
    /// Maximize the PowerPoint window.
    /// </summary>
    [ServiceAction("maximize")]
    OperationResult Maximize(IPptBatch batch);

    /// <summary>
    /// Set the zoom level of the active view (percentage).
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="zoomPercent">Zoom percentage (e.g. 100 for 100%)</param>
    [ServiceAction("set-zoom")]
    OperationResult SetZoom(IPptBatch batch, int zoomPercent);
}
