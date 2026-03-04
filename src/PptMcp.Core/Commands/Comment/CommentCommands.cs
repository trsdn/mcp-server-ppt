using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Comment;

public class CommentCommands : ICommentCommands
{
    public CommentListResult List(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            var result = new CommentListResult { Success = true, FilePath = ctx.PresentationPath };
            dynamic pres = ctx.Presentation;

            if (slideIndex > 0)
            {
                dynamic slide = pres.Slides.Item(slideIndex);
                try
                {
                    ReadCommentsFromSlide(slide, slideIndex, result);
                }
                finally
                {
                    ComUtilities.Release(ref slide!);
                }
            }
            else
            {
                dynamic slides = pres.Slides;
                try
                {
                    int count = (int)slides.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        dynamic slide = slides.Item(i);
                        try
                        {
                            ReadCommentsFromSlide(slide, i, result);
                        }
                        finally
                        {
                            ComUtilities.Release(ref slide!);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref slides!);
                }
            }

            return result;
        });
    }

    public OperationResult Add(IPptBatch batch, int slideIndex, string text, string author, float left, float top)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        ArgumentException.ThrowIfNullOrWhiteSpace(author);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                // Comments.Add2(Left, Top, Author, AuthorInitials, Text)
                string initials = author.Length >= 2
                    ? string.Concat(author.AsSpan(0, 1).ToString().ToUpperInvariant(), author.AsSpan(1, 1))
                    : author.ToUpperInvariant();
                slide.Comments.Add2(left, top, author, initials, text);

                return new OperationResult
                {
                    Success = true,
                    Action = "add",
                    Message = $"Added comment by '{author}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Delete(IPptBatch batch, int slideIndex, int commentIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic? comments = null;
            dynamic? comment = null;
            try
            {
                comments = slide.Comments;
                comment = comments.Item(commentIndex);
                comment.Delete();

                return new OperationResult
                {
                    Success = true,
                    Action = "delete",
                    Message = $"Deleted comment {commentIndex} on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (comment != null) ComUtilities.Release(ref comment!);
                if (comments != null) ComUtilities.Release(ref comments!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Clear(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            int cleared = 0;

            void ClearSlide(dynamic s)
            {
                dynamic comments = s.Comments;
                try
                {
                    // Delete from last to first to avoid index shift
                    for (int i = (int)comments.Count; i >= 1; i--)
                    {
                        dynamic c = comments.Item(i);
                        try { c.Delete(); cleared++; }
                        finally { ComUtilities.Release(ref c!); }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref comments!);
                }
            }

            if (slideIndex > 0)
            {
                dynamic slide = pres.Slides.Item(slideIndex);
                try { ClearSlide(slide); }
                finally { ComUtilities.Release(ref slide!); }
            }
            else
            {
                dynamic slides = pres.Slides;
                try
                {
                    int count = (int)slides.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        dynamic slide = slides.Item(i);
                        try { ClearSlide(slide); }
                        finally { ComUtilities.Release(ref slide!); }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref slides!);
                }
            }

            return new OperationResult
            {
                Success = true,
                Action = "clear",
                Message = slideIndex > 0
                    ? $"Cleared {cleared} comment(s) from slide {slideIndex}"
                    : $"Cleared {cleared} comment(s) from all slides",
                FilePath = ctx.PresentationPath
            };
        });
    }

    private static void ReadCommentsFromSlide(dynamic slide, int slideIdx, CommentListResult result)
    {
        dynamic comments = slide.Comments;
        try
        {
            int count = (int)comments.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic c = comments.Item(i);
                try
                {
                    var info = new CommentInfo
                    {
                        SlideIndex = slideIdx,
                        CommentIndex = i,
                        Text = c.Text?.ToString() ?? "",
                        Author = c.Author?.ToString() ?? "",
                        Left = Convert.ToSingle(c.Left),
                        Top = Convert.ToSingle(c.Top),
                    };
                    try { info.DateTime = c.DateTime?.ToString(); } catch { }
                    result.Comments.Add(info);
                }
                finally
                {
                    ComUtilities.Release(ref c!);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref comments!);
        }
    }
}
