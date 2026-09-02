using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Notes;

public class NotesCommands : INotesCommands
{
    public NotesResult GetNotes(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ComUtilities.GetSlide(ctx.Presentation, slideIndex);
            try
            {
                string text = "";
                // Notes page has placeholders; placeholder 2 is the text body.
                // A slide without a notes placeholder throws, and that is not an error.
                try { text = ComUtilities.GetNotesText(slide); } catch { }

                return new NotesResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    SlideIndex = slideIndex,
                    Text = text
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetNotes(IPptBatch batch, int slideIndex, string text)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ComUtilities.GetSlide(ctx.Presentation, slideIndex);
            try
            {
                ComUtilities.SetNotesText(slide, text);
                return new OperationResult
                {
                    Success = true,
                    Action = "set",
                    Message = $"Set notes on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Clear(IPptBatch batch, int slideIndex)
    {
        return SetNotes(batch, slideIndex, "");
    }

    public OperationResult Append(IPptBatch batch, int slideIndex, string text)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ComUtilities.GetSlide(ctx.Presentation, slideIndex);
            try
            {
                string existing = "";
                try { existing = ComUtilities.GetNotesText(slide); } catch { }

                string newText = string.IsNullOrEmpty(existing) ? text : existing + "\n" + text;
                ComUtilities.SetNotesText(slide, newText);

                return new OperationResult
                {
                    Success = true,
                    Action = "append",
                    Message = $"Appended notes on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult ReadAll(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic slides = pres.Slides;
            try
            {
                int count = (int)slides.Count;
                var lines = new List<string>();

                for (int i = 1; i <= count; i++)
                {
                    dynamic slide = slides.Item(i);
                    try
                    {
                        string text = "";
                        try { text = ComUtilities.GetNotesText(slide); } catch { }

                        lines.Add($"Slide {i}: {text}");
                    }
                    finally
                    {
                        ComUtilities.Release(ref slide!);
                    }
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "read-all",
                    Message = string.Join("\n", lines),
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slides!);
            }
        });
    }
}
