using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.HeaderFooter;

/// <summary>
/// Presentation headers and footers: get settings, set date/page number/footer text.
/// </summary>
[ServiceCategory("headerfooter")]
[McpTool("headerfooter", Title = "Headers & Footers", Destructive = true, Category = "headerfooter",
    Description = "Get and set presentation-wide footer text, slide numbers, and date display. "
    + "Use 'get' to see current settings. Use 'set' with show_footer/show_slide_number/show_date (bool) "
    + "and footer_text (string). Pass null to leave a setting unchanged.")]
public interface IHeaderFooterCommands
{
    /// <summary>Get header/footer settings for the presentation.</summary>
    /// <param name="batch">Batch context</param>
    [ServiceAction("get")]
    HeaderFooterResult GetInfo(IPptBatch batch);

    /// <summary>
    /// Set header/footer options. Pass null to leave unchanged.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="footerText">Footer text (null = don't change)</param>
    /// <param name="showFooter">Show footer on slides</param>
    /// <param name="showSlideNumber">Show slide numbers</param>
    /// <param name="showDate">Show date/time</param>
    [ServiceAction("set")]
    OperationResult Update(IPptBatch batch, string? footerText, bool? showFooter, bool? showSlideNumber, bool? showDate);
}
