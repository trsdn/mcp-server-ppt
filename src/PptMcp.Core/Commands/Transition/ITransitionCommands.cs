using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Transition;

/// <summary>
/// Slide transition effects: get, set, remove.
/// </summary>
[ServiceCategory("transition")]
[McpTool("transition", Title = "Slide Transitions", Destructive = true, Category = "animation")]
public interface ITransitionCommands
{
    /// <summary>Get the transition settings for a slide.</summary>
    [ServiceAction("get")]
    TransitionResult GetTransition(IPptBatch batch, int slideIndex);

    /// <summary>Set a transition effect on a slide.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="transitionType">PpEntryEffect enum value (e.g. 3844=Fade, 3849=Push)</param>
    /// <param name="duration">Duration in seconds</param>
    /// <param name="advanceOnClick">Whether to advance on mouse click</param>
    /// <param name="advanceAfterTime">Auto-advance after N seconds (0 = disabled)</param>
    [ServiceAction("set")]
    OperationResult SetTransition(IPptBatch batch, int slideIndex, int transitionType, float duration, bool advanceOnClick, float advanceAfterTime);

    /// <summary>Remove transition from a slide.</summary>
    [ServiceAction("remove")]
    OperationResult Remove(IPptBatch batch, int slideIndex);

    /// <summary>Copy the transition from one slide to all other slides.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based index of the source slide</param>
    [ServiceAction("copy-to-all")]
    OperationResult CopyToAll(IPptBatch batch, int slideIndex);
}
