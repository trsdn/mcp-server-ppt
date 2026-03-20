using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.DocumentProperty;

/// <summary>
/// Document property management: read and write presentation metadata like title, author, subject, keywords.
/// </summary>
[ServiceCategory("docproperty")]
[McpTool("docproperty", Title = "Document Properties", Destructive = false, Category = "metadata",
    Description = "Read and write presentation metadata: title, author, subject, keywords, comments, company, category. "
    + "Use 'get' for all built-in properties. Use 'set' (pass null to leave unchanged). "
    + "'get-custom'/'set-custom' for arbitrary key-value metadata via property_name/property_value.")]
public interface IDocumentPropertyCommands
{
    /// <summary>
    /// Get all built-in document properties (title, author, subject, keywords, comments, company, category).
    /// </summary>
    [ServiceAction("get")]
    DocumentPropertyResult GetAll(IPptBatch batch);

    /// <summary>
    /// Set built-in document properties. Pass null or empty to leave a property unchanged.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="title">Presentation title</param>
    /// <param name="subject">Subject or topic</param>
    /// <param name="author">Author name</param>
    /// <param name="keywords">Keywords for search (comma-separated)</param>
    /// <param name="comments">Description or comments</param>
    /// <param name="company">Company or organization name</param>
    /// <param name="category">Category</param>
    [ServiceAction("set")]
    OperationResult SetAll(IPptBatch batch, string title, string subject, string author, string keywords, string comments, string company, string category);

    /// <summary>
    /// Get a custom document property by name.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="propertyName">Custom property name</param>
    [ServiceAction("get-custom")]
    OperationResult GetCustom(IPptBatch batch, string propertyName);

    /// <summary>
    /// Set a custom document property (creates if not exists).
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="propertyName">Custom property name</param>
    /// <param name="propertyValue">Property value (string)</param>
    [ServiceAction("set-custom")]
    OperationResult SetCustom(IPptBatch batch, string propertyName, string propertyValue);
}
