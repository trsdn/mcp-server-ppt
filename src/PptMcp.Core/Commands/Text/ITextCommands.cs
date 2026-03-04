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

    /// <summary>
    /// Count words across all slides or a specific slide.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">0 for all slides, or specific 1-based index</param>
    [ServiceAction("word-count")]
    OperationResult WordCount(IPptBatch batch, int slideIndex);

    /// <summary>
    /// Report shapes missing alt text (AlternativeText).
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">0 for all slides, or specific 1-based index</param>
    [ServiceAction("alt-text-audit")]
    OperationResult AltTextAudit(IPptBatch batch, int slideIndex);

    /// <summary>
    /// Find unfilled placeholders with empty text.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">0 for all slides, or specific 1-based index</param>
    [ServiceAction("empty-placeholder-audit")]
    OperationResult EmptyPlaceholderAudit(IPptBatch batch, int slideIndex);

    /// <summary>
    /// Set paragraph and character spacing for text in a shape.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name</param>
    /// <param name="lineSpacing">Line spacing in points (null = don't change)</param>
    /// <param name="spaceBefore">Space before paragraph in points (null = don't change)</param>
    /// <param name="spaceAfter">Space after paragraph in points (null = don't change)</param>
    /// <param name="characterSpacing">Character spacing in points (null = don't change)</param>
    [ServiceAction("set-spacing")]
    OperationResult SetSpacing(IPptBatch batch, int slideIndex, string shapeName, float? lineSpacing, float? spaceBefore, float? spaceAfter, float? characterSpacing);

    /// <summary>
    /// Set bullet point style for text in a shape.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name</param>
    /// <param name="bulletType">0=None, 1=Unnumbered (bullets), 2=Numbered</param>
    /// <param name="bulletCharacter">Custom bullet character (e.g. "•", "→") - only used when bulletType is 1</param>
    /// <param name="indentLevel">Indent level 0-4</param>
    [ServiceAction("set-bullets")]
    OperationResult SetBullets(IPptBatch batch, int slideIndex, string shapeName, int bulletType, string? bulletCharacter, int indentLevel);

    /// <summary>
    /// Insert a hyperlink on existing text within a shape.
    /// Finds linkText within the shape's text and adds a hyperlink to it.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name</param>
    /// <param name="linkText">Text to find and make into a hyperlink</param>
    /// <param name="url">URL for the hyperlink</param>
    [ServiceAction("insert-link")]
    OperationResult InsertLink(IPptBatch batch, int slideIndex, string shapeName, string linkText, string url);

    /// <summary>
    /// Change the case of text in a shape.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name</param>
    /// <param name="caseType">1=Sentence, 2=Lower, 3=Upper, 4=Title, 5=Toggle</param>
    [ServiceAction("change-case")]
    OperationResult ChangeCase(IPptBatch batch, int slideIndex, string shapeName, int caseType);

    /// <summary>
    /// Read paragraph and character spacing from a shape's text.
    /// Returns SpaceWithin, SpaceBefore, SpaceAfter, and character Spacing.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name</param>
    [ServiceAction("read-spacing")]
    OperationResult ReadSpacing(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// Read bullet settings from a shape's text.
    /// Returns Bullet.Type, Bullet.Character, and IndentLevel for each paragraph.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name</param>
    [ServiceAction("read-bullets")]
    OperationResult ReadBullets(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// Insert a symbol character from a specified font into a shape's text.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name</param>
    /// <param name="fontName">Font name containing the symbol (e.g. "Wingdings")</param>
    /// <param name="charNumber">Unicode/character code of the symbol</param>
    [ServiceAction("insert-symbol")]
    OperationResult InsertSymbol(IPptBatch batch, int slideIndex, string shapeName, string fontName, int charNumber);

    /// <summary>
    /// Insert a date/time field into a shape's text.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name</param>
    /// <param name="dateTimeFormat">PpDateTimeFormat value (1-13)</param>
    [ServiceAction("insert-datetime")]
    OperationResult InsertDateTime(IPptBatch batch, int slideIndex, string shapeName, int dateTimeFormat);

    /// <summary>
    /// Insert a slide number field into a shape's text.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name</param>
    [ServiceAction("insert-slide-number")]
    OperationResult InsertSlideNumber(IPptBatch batch, int slideIndex, string shapeName);
}
