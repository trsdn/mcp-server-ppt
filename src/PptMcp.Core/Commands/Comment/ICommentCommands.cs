using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Comment;

/// <summary>
/// Slide comments: list, add, delete.
/// </summary>
[ServiceCategory("comment")]
[McpTool("comment", Title = "Slide Comments", Destructive = true, Category = "comments")]
public interface ICommentCommands
{
    /// <summary>List all comments on a slide (0 = all slides).</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index, or 0 for all slides</param>
    [ServiceAction("list")]
    CommentListResult List(IPptBatch batch, int slideIndex);

    /// <summary>Add a comment to a slide.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="text">Comment text</param>
    /// <param name="author">Author name</param>
    /// <param name="left">Horizontal position in points (0 = top-left)</param>
    /// <param name="top">Vertical position in points (0 = top-left)</param>
    [ServiceAction("add")]
    OperationResult Add(IPptBatch batch, int slideIndex, string text, string author, float left, float top);

    /// <summary>Delete a comment by index on a slide.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="commentIndex">1-based comment index</param>
    [ServiceAction("delete")]
    OperationResult Delete(IPptBatch batch, int slideIndex, int commentIndex);

    /// <summary>Delete all comments on a slide (0 = all slides).</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index, or 0 for all slides</param>
    [ServiceAction("clear")]
    OperationResult Clear(IPptBatch batch, int slideIndex);
}
