using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Proofing;

/// <summary>
/// Proofing and language operations: check spelling, get/set language for text.
/// </summary>
[ServiceCategory("proofing")]
[McpTool("proofing", Title = "Proofing & Language", Destructive = true, Category = "proofing",
    Description = "Spelling check and language settings for presentation text. "
    + "'check-spelling' extracts all unique words for review. "
    + "'set-language' sets proofing language: slide_index=0 for all slides, shape_name='' for all shapes. "
    + "language_id (MsoLanguageID): 1033=English US, 2057=English UK, 1031=German, 1036=French, 1034=Spanish, "
    + "1040=Italian, 1041=Japanese, 2052=Chinese Simplified.")]
public interface IProofingCommands
{
    /// <summary>
    /// Collect all unique words from the presentation text for spelling review.
    /// Returns deduplicated words from all slides and shapes.
    /// </summary>
    /// <param name="batch">Batch context</param>
    [ServiceAction("check-spelling")]
    OperationResult CheckSpelling(IPptBatch batch);

    /// <summary>
    /// Set the proofing language (LanguageID) for text in shapes.
    /// Common MsoLanguageID values: 1033=English US, 2057=English UK, 1031=German, 1036=French, 1034=Spanish, 1040=Italian, 1041=Japanese, 2052=Chinese Simplified.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">0 for all slides, or specific 1-based slide index</param>
    /// <param name="shapeName">Empty string for all shapes on slide, or specific shape name</param>
    /// <param name="languageId">MsoLanguageID value (e.g. 1033 for English US)</param>
    [ServiceAction("set-language")]
    OperationResult SetLanguage(IPptBatch batch, int slideIndex, string shapeName, int languageId);

    /// <summary>
    /// Get the proofing language (LanguageID) of text in a shape.
    /// Returns the MsoLanguageID value and language name.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Shape name</param>
    [ServiceAction("get-language")]
    OperationResult GetLanguage(IPptBatch batch, int slideIndex, string shapeName);
}
