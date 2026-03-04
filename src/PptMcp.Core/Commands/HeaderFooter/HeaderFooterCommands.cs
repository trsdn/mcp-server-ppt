using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.HeaderFooter;

public class HeaderFooterCommands : IHeaderFooterCommands
{
    public HeaderFooterResult GetInfo(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            // Slide-level headers/footers are in Designs(1).SlideMaster.HeadersFooters
            // Simpler: use first slide's HeadersFooters as representative
            dynamic slides = pres.Slides;
            try
            {
                if ((int)slides.Count == 0)
                {
                    return new HeaderFooterResult
                    {
                        Success = true,
                        FilePath = ctx.PresentationPath,
                    };
                }

                dynamic slide = slides.Item(1);
                dynamic? hf = null;
                try
                {
                    hf = slide.HeadersFooters;
                    var result = new HeaderFooterResult
                    {
                        Success = true,
                        FilePath = ctx.PresentationPath,
                    };

                    try { result.ShowFooter = Convert.ToInt32(hf.Footer.Visible) != 0; } catch { }
                    try { result.FooterText = hf.Footer.Text?.ToString(); } catch { }
                    try { result.ShowSlideNumber = Convert.ToInt32(hf.SlideNumber.Visible) != 0; } catch { }
                    try { result.ShowDate = Convert.ToInt32(hf.DateAndTime.Visible) != 0; } catch { }

                    return result;
                }
                finally
                {
                    if (hf != null) ComUtilities.Release(ref hf!);
                    ComUtilities.Release(ref slide!);
                }
            }
            finally
            {
                ComUtilities.Release(ref slides!);
            }
        });
    }

    public OperationResult Update(IPptBatch batch, string? footerText, bool? showFooter, bool? showSlideNumber, bool? showDate)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic slides = pres.Slides;
            try
            {
                int count = (int)slides.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic slide = slides.Item(i);
                    dynamic? hf = null;
                    try
                    {
                        hf = slide.HeadersFooters;

                        if (showFooter.HasValue)
                            hf.Footer.Visible = showFooter.Value ? -1 : 0;
                        if (footerText != null)
                            hf.Footer.Text = footerText;
                        if (showSlideNumber.HasValue)
                            hf.SlideNumber.Visible = showSlideNumber.Value ? -1 : 0;
                        if (showDate.HasValue)
                            hf.DateAndTime.Visible = showDate.Value ? -1 : 0;
                    }
                    finally
                    {
                        if (hf != null) ComUtilities.Release(ref hf!);
                        ComUtilities.Release(ref slide!);
                    }
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "set",
                    Message = $"Updated header/footer settings on {count} slide(s)",
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
