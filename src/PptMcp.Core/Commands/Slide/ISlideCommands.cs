using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Slide;

/// <summary>
/// Slide lifecycle commands: list, read, create, duplicate, move, delete.
/// </summary>
[ServiceCategory("slide")]
[McpTool("slide", Title = "Slide Operations", Destructive = true, Category = "slides")]
public interface ISlideCommands
{
    /// <summary>
    /// List all slides in the presentation with metadata.
    /// </summary>
    [ServiceAction("list")]
    SlideListResult List(IPptBatch batch);

    /// <summary>
    /// Get detailed information about a specific slide including all shapes.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    [ServiceAction("read")]
    SlideDetailResult Read(IPptBatch batch, int slideIndex);

    /// <summary>
    /// Add a new slide at the specified position with a layout.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="position">1-based insert position (0 = at end)</param>
    /// <param name="layoutName">Layout name from the slide master (e.g. "Title Slide", "Blank")</param>
    [ServiceAction("create")]
    OperationResult Create(IPptBatch batch, int position, string layoutName);

    /// <summary>
    /// Duplicate an existing slide.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based index of slide to duplicate</param>
    [ServiceAction("duplicate")]
    OperationResult Duplicate(IPptBatch batch, int slideIndex);

    /// <summary>
    /// Move a slide to a new position.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based index of slide to move</param>
    /// <param name="newPosition">1-based target position</param>
    [ServiceAction("move")]
    OperationResult Move(IPptBatch batch, int slideIndex, int newPosition);

    /// <summary>
    /// Delete a slide by index.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based index of slide to delete</param>
    [ServiceAction("delete")]
    OperationResult Delete(IPptBatch batch, int slideIndex);

    /// <summary>
    /// Apply a layout to an existing slide.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="layoutName">Layout name from the slide master</param>
    [ServiceAction("apply-layout")]
    OperationResult ApplyLayout(IPptBatch batch, int slideIndex, string layoutName);

    /// <summary>Set the name of a slide.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="name">New name for the slide</param>
    [ServiceAction("set-name")]
    OperationResult SetName(IPptBatch batch, int slideIndex, string name);

    /// <summary>
    /// Clone a slide multiple times and replace text placeholders in each clone.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based index of the source slide to clone</param>
    /// <param name="count">Number of clones to create</param>
    /// <param name="searchText">Text to search for in each clone</param>
    /// <param name="replaceText">Text to replace with in each clone</param>
    [ServiceAction("clone-with-replace")]
    OperationResult CloneWithReplace(IPptBatch batch, int slideIndex, int count, string searchText, string replaceText);

    /// <summary>
    /// Hide a slide so it is skipped during slideshow playback.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    [ServiceAction("hide")]
    OperationResult Hide(IPptBatch batch, int slideIndex);

    /// <summary>
    /// Unhide a slide so it is included during slideshow playback.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    [ServiceAction("unhide")]
    OperationResult Unhide(IPptBatch batch, int slideIndex);

    /// <summary>
    /// Export a slide as a PNG thumbnail to the specified file path.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="destinationPath">Full path for the output PNG file</param>
    [ServiceAction("get-thumbnail")]
    OperationResult GetThumbnail(IPptBatch batch, int slideIndex, string destinationPath);
}
