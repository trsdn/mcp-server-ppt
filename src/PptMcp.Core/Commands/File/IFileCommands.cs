using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.File;

/// <summary>
/// File management commands for PowerPoint presentations.
/// Handles file validation and metadata retrieval.
/// </summary>
[ServiceCategory("file")]
[NoSession]
public interface IFileCommands
{
    /// <summary>
    /// Validate a PowerPoint file and return metadata (size, slide count, macro status).
    /// </summary>
    /// <param name="filePath">Path to the .pptx or .pptm file</param>
    [ServiceAction("test")]
    FileValidationInfo Test(string filePath);
}
