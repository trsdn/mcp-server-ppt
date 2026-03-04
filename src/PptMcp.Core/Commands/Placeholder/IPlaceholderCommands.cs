using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Placeholder;

/// <summary>
/// Slide placeholder operations: list available placeholders, fill text.
/// </summary>
[ServiceCategory("placeholder")]
[McpTool("placeholder", Title = "Slide Placeholders", Destructive = true, Category = "placeholders")]
public interface IPlaceholderCommands
{
    /// <summary>List all placeholders on a slide with type and current text.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    [ServiceAction("list")]
    PlaceholderListResult List(IPptBatch batch, int slideIndex);

    /// <summary>Set text content of a placeholder by index.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="placeholderIndex">1-based placeholder index</param>
    /// <param name="text">Text to set</param>
    [ServiceAction("set-text")]
    OperationResult SetText(IPptBatch batch, int slideIndex, int placeholderIndex, string text);
}
