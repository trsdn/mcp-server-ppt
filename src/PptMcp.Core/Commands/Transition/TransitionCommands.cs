using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Transition;

public class TransitionCommands : ITransitionCommands
{
    public TransitionResult GetTransition(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                dynamic trans = slide.SlideShowTransition;
                return new TransitionResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    SlideIndex = slideIndex,
                    TransitionType = Convert.ToInt32(trans.EntryEffect).ToString(),
                    Duration = Convert.ToSingle(trans.Duration),
                    AdvanceOnClick = Convert.ToInt32(trans.AdvanceOnClick) != 0,
                    AdvanceAfterTime = Convert.ToSingle(trans.AdvanceTime)
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetTransition(IPptBatch batch, int slideIndex, int transitionType, float duration, bool advanceOnClick, float advanceAfterTime)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                dynamic trans = slide.SlideShowTransition;
                trans.EntryEffect = transitionType;
                trans.Duration = duration;
                trans.AdvanceOnClick = advanceOnClick ? -1 : 0;

                if (advanceAfterTime > 0)
                {
                    trans.AdvanceOnTime = -1; // msoTrue
                    trans.AdvanceTime = advanceAfterTime;
                }
                else
                {
                    trans.AdvanceOnTime = 0; // msoFalse
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "set",
                    Message = $"Set transition on slide {slideIndex} (effect {transitionType}, {duration}s)",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Remove(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                // ppEffectNone = 0
                slide.SlideShowTransition.EntryEffect = 0;
                return new OperationResult
                {
                    Success = true,
                    Action = "remove",
                    Message = $"Removed transition from slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult CopyToAll(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic srcSlide = pres.Slides.Item(slideIndex);
            dynamic slides = pres.Slides;
            try
            {
                dynamic srcTrans = srcSlide.SlideShowTransition;
                int effect = Convert.ToInt32(srcTrans.EntryEffect);
                float duration = Convert.ToSingle(srcTrans.Duration);
                int advClick = Convert.ToInt32(srcTrans.AdvanceOnClick);
                int advTime = Convert.ToInt32(srcTrans.AdvanceOnTime);
                float advSeconds = Convert.ToSingle(srcTrans.AdvanceTime);

                int count = (int)slides.Count;
                int applied = 0;
                for (int i = 1; i <= count; i++)
                {
                    if (i == slideIndex) continue;
                    dynamic slide = slides.Item(i);
                    try
                    {
                        dynamic trans = slide.SlideShowTransition;
                        trans.EntryEffect = effect;
                        trans.Duration = duration;
                        trans.AdvanceOnClick = advClick;
                        trans.AdvanceOnTime = advTime;
                        if (advTime != 0) trans.AdvanceTime = advSeconds;
                        applied++;
                    }
                    finally
                    {
                        ComUtilities.Release(ref slide!);
                    }
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "copy-to-all",
                    Message = $"Copied transition from slide {slideIndex} to {applied} other slide(s)",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slides!);
                ComUtilities.Release(ref srcSlide!);
            }
        });
    }
}
