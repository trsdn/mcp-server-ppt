using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.CustomShow;

/// <summary>
/// Custom slide show management: list, create, delete, run.
/// </summary>
[ServiceCategory("customshow")]
[McpTool("customshow", Title = "Custom Shows", Destructive = true, Category = "customshow")]
public interface ICustomShowCommands
{
    /// <summary>List all custom shows in the presentation.</summary>
    [ServiceAction("list")]
    CustomShowListResult List(IPptBatch batch);

    /// <summary>Create a custom show from specified slide indices.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="showName">Name for the custom show</param>
    /// <param name="slideIndices">Comma-separated 1-based slide indices (e.g. "1,3,5")</param>
    [ServiceAction("create")]
    OperationResult Create(IPptBatch batch, string showName, string slideIndices);

    /// <summary>Delete a custom show by name.</summary>
    [ServiceAction("delete")]
    OperationResult Delete(IPptBatch batch, string showName);
}
