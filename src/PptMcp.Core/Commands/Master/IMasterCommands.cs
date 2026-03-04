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
}
