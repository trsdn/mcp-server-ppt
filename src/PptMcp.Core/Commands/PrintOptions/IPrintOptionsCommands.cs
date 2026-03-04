using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.PrintOptions;

/// <summary>
/// Manage print options: output type, color mode, framing, fit-to-page, hidden slides.
/// </summary>
[ServiceCategory("printoptions")]
[McpTool("printoptions", Title = "Print Options", Destructive = true, Category = "print")]
public interface IPrintOptionsCommands
{
    /// <summary>
    /// Get current print settings: output type, color type, frame slides, fit to page, print hidden slides, number of copies.
    /// </summary>
    [ServiceAction("get")]
    OperationResult GetSettings(IPptBatch batch);

    /// <summary>
    /// Set print settings. Only non-null values are changed.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="outputType">1=Slides, 2=TwoSlideHandouts, 3=ThreeSlideHandouts, 4=SixSlideHandouts, 5=NotesPages, 6=Outline</param>
    /// <param name="colorType">1=Color, 2=Grayscale, 3=BlackWhite</param>
    /// <param name="frameSlides">Whether to frame slides when printing</param>
    /// <param name="fitToPage">Whether to fit slides to page</param>
    /// <param name="printHiddenSlides">Whether to include hidden slides</param>
    [ServiceAction("set")]
    OperationResult SetSettings(IPptBatch batch, int? outputType, int? colorType, bool? frameSlides, bool? fitToPage, bool? printHiddenSlides);
}
