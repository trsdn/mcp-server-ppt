using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.SlideImport;

/// <summary>
/// Import slides from another presentation file.
/// </summary>
[ServiceCategory("slideimport")]
[McpTool("slideimport", Title = "Slide Import", Destructive = true, Category = "slideimport")]
public interface ISlideImportCommands
{
    /// <summary>
    /// Import slides from another PowerPoint file.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="sourceFilePath">Path to the source .pptx file</param>
    /// <param name="slideIndices">Comma-separated 1-based slide indices to import (empty = all)</param>
    /// <param name="insertAt">Position to insert (0 = at end)</param>
    [ServiceAction("import")]
    OperationResult ImportSlides(IPptBatch batch, string sourceFilePath, string slideIndices, int insertAt);
}
