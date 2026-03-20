using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.SlideImport;

/// <summary>
/// Import slides from another presentation file.
/// </summary>
[ServiceCategory("slideimport")]
[McpTool("slideimport", Title = "Slide Import", Destructive = true, Category = "slideimport",
    Description = "Import slides from another .pptx/.pptm file into the current presentation. "
    + "slide_indices: comma-separated 1-based (e.g. '1,3,5'). Empty = import all slides. "
    + "insert_at: 0 = append at end. Source file must not be open in another session.")]
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
