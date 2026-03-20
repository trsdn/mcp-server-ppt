using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Notes;

/// <summary>
/// Speaker notes: get, set, clear.
/// </summary>
[ServiceCategory("notes")]
[McpTool("notes", Title = "Speaker Notes", Destructive = true, Category = "notes",
    Description = "Get, set, clear, or append speaker notes per slide. "
    + "Use 'read-all' to get notes from every slide at once. "
    + "'append' adds text with a newline separator to existing notes. "
    + "Useful for building presenter scripts alongside slide creation.")]
public interface INotesCommands
{
    /// <summary>Get speaker notes for a slide.</summary>
    [ServiceAction("get")]
    NotesResult GetNotes(IPptBatch batch, int slideIndex);

    /// <summary>Set speaker notes for a slide.</summary>
    [ServiceAction("set")]
    OperationResult SetNotes(IPptBatch batch, int slideIndex, string text);

    /// <summary>Clear speaker notes for a slide.</summary>
    [ServiceAction("clear")]
    OperationResult Clear(IPptBatch batch, int slideIndex);

    /// <summary>Append text to existing speaker notes (adds newline separator).</summary>
    [ServiceAction("append")]
    OperationResult Append(IPptBatch batch, int slideIndex, string text);

    /// <summary>Read speaker notes from all slides in the presentation.</summary>
    [ServiceAction("read-all")]
    OperationResult ReadAll(IPptBatch batch);
}
