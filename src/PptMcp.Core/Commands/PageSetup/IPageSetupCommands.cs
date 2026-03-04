using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.PageSetup;

/// <summary>
/// Slide size and page setup operations.
/// </summary>
[ServiceCategory("pagesetup")]
[McpTool("pagesetup", Title = "Page Setup", Destructive = true, Category = "pagesetup")]
public interface IPageSetupCommands
{
    /// <summary>Get the current slide size and orientation.</summary>
    [ServiceAction("get")]
    PageSetupResult GetInfo(IPptBatch batch);

    /// <summary>
    /// Set the slide size. Common sizes: 0=OnScreen (4:3), 1=LetterPaper, 2=Overhead,
    /// 3=A3, 4=A4, 5=B4ISO, 6=B5ISO, 7=35mm, 8=Custom, 9=OnScreen16x9, 10=OnScreen16x10.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideWidth">Slide width in points (1 inch = 72 points). 0 = don't change.</param>
    /// <param name="slideHeight">Slide height in points. 0 = don't change.</param>
    [ServiceAction("set-size")]
    OperationResult SetSize(IPptBatch batch, float slideWidth, float slideHeight);
}
