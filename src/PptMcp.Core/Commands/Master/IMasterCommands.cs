using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Master;

/// <summary>
/// Slide master and layout operations: list masters, list layouts, get placeholders.
/// </summary>
[ServiceCategory("master")]
[McpTool("master", Title = "Master & Layout Operations", Destructive = false, Category = "design")]
public interface IMasterCommands
{
    /// <summary>List all slide masters and their custom layouts.</summary>
    [ServiceAction("list")]
    MasterListResult List(IPptBatch batch);

    /// <summary>List all shapes on a specific slide master.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="masterIndex">1-based slide master index</param>
    [ServiceAction("list-shapes")]
    OperationResult ListShapes(IPptBatch batch, int masterIndex);

    /// <summary>Edit the text content of a shape on a slide master.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="masterIndex">1-based slide master index</param>
    /// <param name="shapeName">Name of the shape to edit</param>
    /// <param name="text">New text content</param>
    [ServiceAction("edit-shape-text")]
    OperationResult EditShapeText(IPptBatch batch, int masterIndex, string shapeName, string text);

    /// <summary>List all custom layouts for a specific slide master.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="masterIndex">1-based slide master index</param>
    [ServiceAction("list-layouts")]
    OperationResult ListLayouts(IPptBatch batch, int masterIndex);

    /// <summary>Delete unused slide masters that have no slides referencing them. Will not delete the last remaining master.</summary>
    /// <param name="batch">Batch context</param>
    [ServiceAction("delete-unused")]
    OperationResult DeleteUnused(IPptBatch batch);
}
