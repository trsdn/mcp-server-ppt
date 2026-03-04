using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Slideshow;

public class SlideshowCommands : ISlideshowCommands
{
    public OperationResult Start(IPptBatch batch, int startSlide)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic settings = pres.SlideShowSettings;
            try
            {
                if (startSlide > 0)
                {
                    settings.StartingSlide = startSlide;
                    settings.EndingSlide = (int)pres.Slides.Count;
                }

                // ppShowTypeSpeaker = 1 (full screen)
                settings.ShowType = 1;
                dynamic window = settings.Run();
                ComUtilities.Release(ref window!);

                return new OperationResult
                {
                    Success = true,
                    Action = "start",
                    Message = startSlide > 0
                        ? $"Started slideshow from slide {startSlide}"
                        : "Started slideshow from beginning",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref settings!);
            }
        });
    }

    public OperationResult EndShow(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic? window = null;
            try
            {
                window = pres.SlideShowWindow;
                dynamic? view = null;
                try
                {
                    view = window.View;
                    view.Exit();
                }
                finally
                {
                    if (view != null) ComUtilities.Release(ref view!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "stop",
                    Message = "Stopped slideshow",
                    FilePath = ctx.PresentationPath
                };
            }
            catch
            {
                // No slideshow running
                return new OperationResult
                {
                    Success = true,
                    Action = "stop",
                    Message = "No slideshow was running",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (window != null) ComUtilities.Release(ref window!);
            }
        });
    }

    public OperationResult GotoSlide(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic window = pres.SlideShowWindow;
            dynamic view = window.View;
            try
            {
                view.GotoSlide(slideIndex);
                return new OperationResult
                {
                    Success = true,
                    Action = "goto-slide",
                    Message = $"Navigated to slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref view!);
                ComUtilities.Release(ref window!);
            }
        });
    }

    public SlideshowInfoResult GetStatus(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            int totalSlides = (int)pres.Slides.Count;

            bool isRunning = false;
            int currentSlide = 0;
            try
            {
                dynamic window = pres.SlideShowWindow;
                dynamic? view = null;
                try
                {
                    view = window.View;
                    isRunning = true;
                    currentSlide = (int)view.CurrentShowPosition;
                }
                finally
                {
                    if (view != null) ComUtilities.Release(ref view!);
                    ComUtilities.Release(ref window!);
                }
            }
            catch
            {
                // No slideshow running
            }

            return new SlideshowInfoResult
            {
                Success = true,
                FilePath = ctx.PresentationPath,
                IsRunning = isRunning,
                CurrentSlide = currentSlide,
                TotalSlides = totalSlides
            };
        });
    }
}
