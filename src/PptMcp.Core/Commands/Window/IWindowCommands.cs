using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Window;

/// <summary>
/// PowerPoint window management: get info, minimize, restore, maximize.
/// </summary>
[ServiceCategory("window")]
[McpTool("window", Title = "Window Operations", Destructive = false, Category = "window",
    Description = "Control PowerPoint window: visibility, position, zoom, view mode. "
    + "Use for 'Agent Mode' (user watches AI work): window(get-info) to check state, then minimize/restore/maximize. "
    + "set-zoom: zoom_percent (e.g. 100 for 100%). "
    + "set-view view_type: 1=Normal, 2=Outline, 3=SlideSorter, 4=NotesPage, 5=SlideMaster.")]
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

    /// <summary>
    /// Set the view type of the active window.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="viewType">1=Normal, 2=Outline, 3=SlideSorter, 4=NotesPage, 5=SlideMaster</param>
    [ServiceAction("set-view")]
    OperationResult SetView(IPptBatch batch, int viewType);

    /// <summary>
    /// Get the current view type of the active window.
    /// </summary>
    [ServiceAction("get-view")]
    OperationResult GetView(IPptBatch batch);
}
