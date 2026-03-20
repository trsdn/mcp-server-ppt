using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Accessibility;

/// <summary>
/// Accessibility audit: check alt text, title placeholders, reading order.
/// </summary>
[ServiceCategory("accessibility")]
[McpTool("accessibility", Title = "Accessibility Audit", Destructive = false, Category = "accessibility",
    Description = "Audit presentation accessibility: missing alt text, empty title placeholders, reading order issues. "
    + "Use 'audit' for full-presentation scan. Use 'get-reading-order'/'set-reading-order' to fix tab order per slide. "
    + "shape_names for set-reading-order: comma-separated names in desired order.")]
public interface IAccessibilityCommands
{
    /// <summary>
    /// Audit the entire presentation for accessibility issues: missing alt text, missing title placeholders, empty placeholders.
    /// </summary>
    [ServiceAction("audit")]
    AccessibilityAuditResult Audit(IPptBatch batch);

    /// <summary>
    /// Get the reading order (tab order) of shapes on a slide, listed by ZOrderPosition.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    [ServiceAction("get-reading-order")]
    ReadingOrderResult GetReadingOrder(IPptBatch batch, int slideIndex);

    /// <summary>
    /// Set the reading order of shapes on a slide by reordering their ZOrderPosition.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeNames">Comma-separated shape names in desired reading order</param>
    [ServiceAction("set-reading-order")]
    OperationResult SetReadingOrder(IPptBatch batch, int slideIndex, string shapeNames);
}
