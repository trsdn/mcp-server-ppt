using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.ShapeAlign;

/// <summary>
/// Shape alignment and distribution operations.
/// </summary>
[ServiceCategory("shapealign")]
[McpTool("shapealign", Title = "Shape Alignment", Destructive = true, Category = "shapealign")]
public interface IShapeAlignCommands
{
    /// <summary>
    /// Align shapes on a slide.
    /// alignType: 0=AlignLeft, 1=AlignCenter, 2=AlignRight, 3=AlignTop, 4=AlignMiddle, 5=AlignBottom
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeNames">Comma-separated shape names</param>
    /// <param name="alignType">Alignment type (0-5)</param>
    [ServiceAction("align")]
    OperationResult Align(IPptBatch batch, int slideIndex, string shapeNames, int alignType);

    /// <summary>
    /// Distribute shapes evenly on a slide.
    /// distributeType: 0=Horizontally, 1=Vertically
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeNames">Comma-separated shape names</param>
    /// <param name="distributeType">0=Horizontally, 1=Vertically</param>
    [ServiceAction("distribute")]
    OperationResult Distribute(IPptBatch batch, int slideIndex, string shapeNames, int distributeType);
}
