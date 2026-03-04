using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Tag;

/// <summary>
/// Custom tags/metadata on slides and shapes.
/// </summary>
[ServiceCategory("tag")]
[McpTool("tag", Title = "Tags & Metadata", Destructive = true, Category = "tags")]
public interface ITagCommands
{
    /// <summary>List all tags on a slide or shape.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name (null/empty = slide-level tags)</param>
    [ServiceAction("list")]
    TagListResult List(IPptBatch batch, int slideIndex, string? shapeName);

    /// <summary>Set a tag value on a slide or shape.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name (null/empty = slide-level tag)</param>
    /// <param name="tagName">Tag name (case-insensitive)</param>
    /// <param name="tagValue">Tag value</param>
    [ServiceAction("set")]
    OperationResult SetTag(IPptBatch batch, int slideIndex, string? shapeName, string tagName, string tagValue);

    /// <summary>Delete a tag from a slide or shape.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name (null/empty = slide-level tag)</param>
    /// <param name="tagName">Tag name to delete</param>
    [ServiceAction("delete")]
    OperationResult DeleteTag(IPptBatch batch, int slideIndex, string? shapeName, string tagName);
}
