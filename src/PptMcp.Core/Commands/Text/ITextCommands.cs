using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Text;

/// <summary>
/// Text operations within shapes: get, set, format, find, replace.
/// </summary>
[ServiceCategory("text")]
[McpTool("text", Title = "Text Operations", Destructive = true, Category = "text")]
public interface ITextCommands
{
    /// <summary>
    /// Get text content from a shape including paragraph and run details.
    /// </summary>
    [ServiceAction("get")]
    TextResult GetText(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// Set the text content of a shape (replaces all existing text).
    /// </summary>
    [ServiceAction("set")]
    OperationResult SetText(IPptBatch batch, int slideIndex, string shapeName, string text);

    /// <summary>
    /// Find text across all shapes in a slide or entire presentation.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="searchText">Text to find</param>
    /// <param name="slideIndex">0 for all slides, or specific 1-based index</param>
    [ServiceAction("find")]
    OperationResult Find(IPptBatch batch, string searchText, int slideIndex);

    /// <summary>
    /// Replace text across all shapes in a slide or entire presentation.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="searchText">Text to find</param>
    /// <param name="replaceText">Replacement text</param>
    /// <param name="slideIndex">0 for all slides, or specific 1-based index</param>
    [ServiceAction("replace")]
    OperationResult Replace(IPptBatch batch, string searchText, string replaceText, int slideIndex);

    /// <summary>
    /// Format text in a shape (font, size, bold, italic, color, alignment).
    /// Horizontal alignment: left, center, right, justify.
    /// Vertical alignment: top, middle, bottom.
    /// </summary>
    [ServiceAction("format")]
    OperationResult Format(IPptBatch batch, int slideIndex, string shapeName, string? fontName, float? fontSize, bool? bold, bool? italic, string? color, string? alignment, string? verticalAlignment);

    /// <summary>
    /// Set advanced text formatting: underline, strikethrough, subscript, superscript.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name</param>
    /// <param name="underline">Set underline (null = don't change)</param>
    /// <param name="strikethrough">Set strikethrough (null = don't change)</param>
    /// <param name="subscript">Set subscript (null = don't change)</param>
    /// <param name="superscript">Set superscript (null = don't change)</param>
    [ServiceAction("format-advanced")]
    OperationResult FormatAdvanced(IPptBatch batch, int slideIndex, string shapeName, bool? underline, bool? strikethrough, bool? subscript, bool? superscript);
}
