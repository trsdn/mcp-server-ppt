using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Section;

/// <summary>
/// Presentation section management: list, add, rename, delete, and move sections.
/// Sections group slides for easier navigation and organization.
/// </summary>
[ServiceCategory("section")]
[McpTool("section", Title = "Section Operations", Destructive = true, Category = "structure",
    Description = "Organize slides into named sections for navigation and structure. "
    + "Use 'list' to see all sections with slide ranges. 'add' creates a section starting at a slide. "
    + "'delete' removes the section marker (slides are kept). section_index: 1-based.")]
public interface ISectionCommands
{
    /// <summary>
    /// List all sections in the presentation with their slide ranges.
    /// </summary>
    [ServiceAction("list")]
    SectionListResult List(IPptBatch batch);

    /// <summary>
    /// Add a new section starting at a specific slide.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="sectionName">Name for the new section</param>
    /// <param name="slideIndex">1-based slide index where the section starts</param>
    [ServiceAction("add")]
    OperationResult Add(IPptBatch batch, string sectionName, int slideIndex);

    /// <summary>
    /// Rename an existing section.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="sectionIndex">1-based section index</param>
    /// <param name="newName">New section name</param>
    [ServiceAction("rename")]
    OperationResult Rename(IPptBatch batch, int sectionIndex, string newName);

    /// <summary>
    /// Delete a section (slides are kept, only section marker is removed).
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="sectionIndex">1-based section index</param>
    [ServiceAction("delete")]
    OperationResult Delete(IPptBatch batch, int sectionIndex);
}
