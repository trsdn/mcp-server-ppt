using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Slideshow;

/// <summary>
/// Slideshow presentation mode: start, stop, navigate, get status.
/// </summary>
[ServiceCategory("slideshow")]
[McpTool("slideshow", Title = "Slideshow Operations", Destructive = false, Category = "slideshow",
    Description = "Control presentation slideshow mode: start, stop, navigate, configure. "
    + "show_type for configure: 1=Speaker (fullscreen), 2=Browsed by individual (window), 3=Kiosk (loop). "
    + "Use 'start' with start_slide (1-based, 0=beginning). "
    + "'goto-slide' navigates during active show. 'get-status' checks if show is running.")]
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

    /// <summary>
    /// Configure slideshow settings (show type, looping, animation, narration).
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="showType">1=Speaker (full screen), 2=Browsed by individual (window), 3=Browsed at kiosk (loop)</param>
    /// <param name="loopUntilStopped">Whether to loop the slideshow continuously</param>
    /// <param name="showWithAnimation">Whether to show animations during the slideshow</param>
    /// <param name="showWithNarration">Whether to play narrations during the slideshow</param>
    [ServiceAction("configure")]
    OperationResult Configure(IPptBatch batch, int showType, bool loopUntilStopped, bool showWithAnimation, bool showWithNarration);
}
