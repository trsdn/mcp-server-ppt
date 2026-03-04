using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Hyperlink;

/// <summary>
/// Hyperlink management: add, remove, and get hyperlinks on shapes and text.
/// </summary>
[ServiceCategory("hyperlink")]
[McpTool("hyperlink", Title = "Hyperlink Operations", Destructive = true, Category = "content")]
public interface IHyperlinkCommands
{
    /// <summary>
    /// Add a hyperlink to a shape (click on shape navigates to URL or slide).
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the shape to add hyperlink to</param>
    /// <param name="address">URL (https://...) or empty for slide link</param>
    /// <param name="subAddress">Slide number for internal links (e.g. '3' to jump to slide 3), or empty</param>
    /// <param name="screenTip">Optional tooltip text shown on hover</param>
    [ServiceAction("add")]
    OperationResult Add(IPptBatch batch, int slideIndex, string shapeName, string address, string? subAddress = null, string? screenTip = null);

    /// <summary>
    /// Get the hyperlink on a shape.
    /// </summary>
    [ServiceAction("get")]
    HyperlinkResult Read(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// Remove the hyperlink from a shape.
    /// </summary>
    [ServiceAction("remove")]
    OperationResult Remove(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// List all hyperlinks in the presentation.
    /// </summary>
    [ServiceAction("list")]
    HyperlinkListResult List(IPptBatch batch);

    /// <summary>
    /// Validate all hyperlinks in the presentation and report broken or empty ones.
    /// Checks every slide and shape for hyperlinks, classifying each as valid, broken, empty, external, or internal.
    /// </summary>
    /// <param name="batch">Batch context</param>
    [ServiceAction("validate")]
    OperationResult Validate(IPptBatch batch);
}
