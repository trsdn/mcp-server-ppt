using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.SmartArt;

/// <summary>
/// SmartArt diagram operations: create, add/remove nodes, change layout.
/// </summary>
[ServiceCategory("smartart")]
[McpTool("smartart", Title = "SmartArt Diagrams", Destructive = true, Category = "smartart")]
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
}
