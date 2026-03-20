using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.SmartArt;

/// <summary>
/// SmartArt diagram operations: create, add/remove nodes, change layout.
/// </summary>
[ServiceCategory("smartart")]
[McpTool("smartart", Title = "SmartArt Diagrams", Destructive = true, Category = "smartart",
    Description = "Create and modify SmartArt diagrams (org charts, process flows, lists). "
    + "Use 'get-info' to inspect an existing SmartArt shape. 'add-node' appends text nodes. "
    + "'set-layout' changes diagram type (layout_index: 1-based from Application.SmartArtLayouts). "
    + "'set-style' changes visual style. 'change-level' promotes/demotes nodes in hierarchy. "
    + "node_index: 1-based.")]
public interface ISmartArtCommands
{
    /// <summary>Get SmartArt info from a shape.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the SmartArt shape</param>
    [ServiceAction("get-info")]
    SmartArtInfoResult GetInfo(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>Add a text node to an existing SmartArt diagram.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the SmartArt shape</param>
    /// <param name="text">Text for the new node</param>
    [ServiceAction("add-node")]
    OperationResult AddNode(IPptBatch batch, int slideIndex, string shapeName, string text);

    /// <summary>Change the layout of a SmartArt diagram.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the SmartArt shape</param>
    /// <param name="layoutIndex">1-based index into Application.SmartArtLayouts</param>
    [ServiceAction("set-layout")]
    OperationResult SetLayout(IPptBatch batch, int slideIndex, string shapeName, int layoutIndex);

    /// <summary>Change the quick style of a SmartArt diagram.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the SmartArt shape</param>
    /// <param name="styleIndex">1-based index into Application.SmartArtQuickStyles</param>
    [ServiceAction("set-style")]
    OperationResult SetStyle(IPptBatch batch, int slideIndex, string shapeName, int styleIndex);

    /// <summary>Delete a node from a SmartArt diagram.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the SmartArt shape</param>
    /// <param name="nodeIndex">1-based index of the node to delete</param>
    [ServiceAction("delete-node")]
    OperationResult DeleteNode(IPptBatch batch, int slideIndex, string shapeName, int nodeIndex);

    /// <summary>Promote or demote a node in a SmartArt diagram to change its level.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the SmartArt shape</param>
    /// <param name="nodeIndex">1-based index of the node</param>
    /// <param name="promote">True to promote (decrease level), false to demote (increase level)</param>
    [ServiceAction("change-level")]
    OperationResult ChangeNodeLevel(IPptBatch batch, int slideIndex, string shapeName, int nodeIndex, bool promote);
}
