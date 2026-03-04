using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Background;

/// <summary>
/// Slide background: get, set solid color, set image, reset to master.
/// </summary>
[ServiceCategory("background")]
[McpTool("background", Title = "Slide Background", Destructive = true, Category = "background")]
public interface IBackgroundCommands
{
    /// <summary>Get the current background info for a slide.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    [ServiceAction("get")]
    BackgroundResult GetInfo(IPptBatch batch, int slideIndex);

    /// <summary>Set a solid color background for a slide.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="colorHex">Hex color string (#RRGGBB)</param>
    [ServiceAction("set-color")]
    OperationResult SetColor(IPptBatch batch, int slideIndex, string colorHex);

    /// <summary>Reset a slide background to follow the slide master.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    [ServiceAction("reset")]
    OperationResult Reset(IPptBatch batch, int slideIndex);
}
