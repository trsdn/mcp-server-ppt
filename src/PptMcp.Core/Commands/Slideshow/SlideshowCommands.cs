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
                    settings.EndingSlide = ComUtilities.GetSlideCount(pres);
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
                // COM offers no IsSlideShowRunning, so acquiring SlideShowWindow and
                // letting it throw *is* the query. The probe is deliberately narrowed to
                // that one statement: the old catch spanned View.Exit() as well, so a
                // genuine failure while stopping a running show was reported back as
                // "No slideshow was running" (#126).
                try
                {
                    window = pres.SlideShowWindow;
                }
                catch
                {
                    return new OperationResult
                    {
                        Success = true,
                        Action = "stop",
                        Message = "No slideshow was running",
                        FilePath = ctx.PresentationPath
                    };
                }

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
            int totalSlides = ComUtilities.GetSlideCount(pres);

            bool isRunning = false;
            int currentSlide = 0;

            // Same narrowed existence probe as EndShow: only the acquisition may throw
            // to mean "not running". The old catch also spanned the CurrentShowPosition
            // read, so a failure there reported a stopped show at slide 0 (#126).
            dynamic? window = null;
            try
            {
                window = pres.SlideShowWindow;
            }
            catch
            {
                window = null;
            }

            if (window != null)
            {
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

    public OperationResult Configure(IPptBatch batch, int showType, bool loopUntilStopped, bool showWithAnimation, bool showWithNarration)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic settings = pres.SlideShowSettings;
            try
            {
                int type = showType >= 1 && showType <= 3 ? showType : 1;
                settings.ShowType = type;
                settings.LoopUntilStopped = loopUntilStopped ? -1 : 0;
                settings.ShowWithAnimation = showWithAnimation ? -1 : 0;
                settings.ShowWithNarration = showWithNarration ? -1 : 0;

                string typeName = type switch
                {
                    1 => "Speaker (full screen)",
                    2 => "Browsed by individual (window)",
                    3 => "Browsed at kiosk (loop)",
                    _ => "Unknown"
                };

                return new OperationResult
                {
                    Success = true,
                    Action = "configure",
                    Message = $"Configured slideshow: type={typeName}, loop={loopUntilStopped}, animation={showWithAnimation}, narration={showWithNarration}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref settings!);
            }
        });
    }
}
