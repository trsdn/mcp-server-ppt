using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Slideshow;

/// <summary>
/// Slideshow presentation mode: start, stop, navigate, get status.
/// </summary>
[ServiceCategory("slideshow")]
[McpTool("slideshow", Title = "Slideshow Operations", Destructive = false, Category = "slideshow")]
public interface ISlideshowCommands
{
    /// <summary>
    /// Start the slideshow from a specific slide.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="startSlide">1-based slide to start from (0 = beginning)</param>
    [ServiceAction("start")]
    OperationResult Start(IPptBatch batch, int startSlide);

    /// <summary>
    /// Stop/end the running slideshow.
    /// </summary>
    [ServiceAction("stop")]
    OperationResult EndShow(IPptBatch batch);

    /// <summary>
    /// Navigate to a specific slide in the running slideshow.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based target slide index</param>
    [ServiceAction("goto-slide")]
    OperationResult GotoSlide(IPptBatch batch, int slideIndex);

    /// <summary>
    /// Get the current slideshow status.
    /// </summary>
    [ServiceAction("get-status")]
    SlideshowInfoResult GetStatus(IPptBatch batch);
}
