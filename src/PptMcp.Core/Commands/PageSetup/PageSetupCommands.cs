using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.PageSetup;

public class PageSetupCommands : IPageSetupCommands
{
    public PageSetupResult GetInfo(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic pageSetup = pres.PageSetup;
            try
            {
                return new PageSetupResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    SlideWidth = Convert.ToSingle(pageSetup.SlideWidth),
                    SlideHeight = Convert.ToSingle(pageSetup.SlideHeight),
                    SlideOrientation = Convert.ToInt32(pageSetup.SlideOrientation),
                    NotesOrientation = Convert.ToInt32(pageSetup.NotesOrientation)
                };
            }
            finally
            {
                ComUtilities.Release(ref pageSetup!);
            }
        });
    }

    public OperationResult SetSize(IPptBatch batch, float slideWidth, float slideHeight)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic pageSetup = pres.PageSetup;
            try
            {
                if (slideWidth > 0) pageSetup.SlideWidth = slideWidth;
                if (slideHeight > 0) pageSetup.SlideHeight = slideHeight;

                return new OperationResult
                {
                    Success = true,
                    Action = "set-size",
                    Message = $"Set slide size to {slideWidth}x{slideHeight} points",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref pageSetup!);
            }
        });
    }

    public OperationResult SetFirstNumber(IPptBatch batch, int firstSlideNumber)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic pageSetup = pres.PageSetup;
            try
            {
                pageSetup.FirstSlideNumber = firstSlideNumber;

                return new OperationResult
                {
                    Success = true,
                    Action = "set-first-number",
                    Message = $"Set first slide number to {firstSlideNumber}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref pageSetup!);
            }
        });
    }
}
