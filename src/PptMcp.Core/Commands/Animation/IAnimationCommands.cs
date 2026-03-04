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

    /// <summary>
    /// Set timing properties for an animation effect.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="effectIndex">1-based index of the effect in the animation sequence</param>
    /// <param name="duration">Duration in seconds</param>
    /// <param name="delay">Delay before start in seconds</param>
    /// <param name="triggerType">1=OnClick, 2=WithPrevious, 3=AfterPrevious</param>
    [ServiceAction("set-timing")]
    OperationResult SetTiming(IPptBatch batch, int slideIndex, int effectIndex, float duration, float delay, int triggerType);

    /// <summary>
    /// Reorder an animation effect by moving it to a new position in the sequence.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="effectIndex">1-based index of the effect to move</param>
    /// <param name="newIndex">1-based target position in the sequence</param>
    [ServiceAction("reorder")]
    OperationResult Reorder(IPptBatch batch, int slideIndex, int effectIndex, int newIndex);
}
