using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Animation;

/// <summary>
/// Animation effect operations: list, add, remove, reorder effects on slides.
/// </summary>
[ServiceCategory("animation")]
[McpTool("animation", Title = "Animation Operations", Destructive = true, Category = "animations")]
public interface IAnimationCommands
{
    /// <summary>
    /// List all animation effects on a slide.
    /// </summary>
    [ServiceAction("list")]
    AnimationListResult List(IPptBatch batch, int slideIndex);

    /// <summary>
    /// Add an animation effect to a shape.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the target shape</param>
    /// <param name="effectType">MsoAnimEffect integer (e.g., 1=Appear, 2=Fly, 10=Fade, 16=Wipe)</param>
    /// <param name="triggerType">1=OnClick (default), 2=WithPrevious, 3=AfterPrevious</param>
    [ServiceAction("add")]
    OperationResult Add(IPptBatch batch, int slideIndex, string shapeName, int effectType, int triggerType);

    /// <summary>
    /// Remove an animation effect by its 1-based index in the animation sequence.
    /// </summary>
    [ServiceAction("remove")]
    OperationResult Remove(IPptBatch batch, int slideIndex, int effectIndex);

    /// <summary>
    /// Remove all animation effects from a slide.
    /// </summary>
    [ServiceAction("clear")]
    OperationResult Clear(IPptBatch batch, int slideIndex);
}
