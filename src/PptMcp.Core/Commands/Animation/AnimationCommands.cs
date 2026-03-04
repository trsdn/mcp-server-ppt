using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Animation;

public class AnimationCommands : IAnimationCommands
{
    public AnimationListResult List(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic? timeline = null;
            dynamic? mainSequence = null;
            try
            {
                timeline = slide.TimeLine;
                mainSequence = timeline.MainSequence;
                int count = (int)mainSequence.Count;

                var result = new AnimationListResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    SlideIndex = slideIndex
                };

                for (int i = 1; i <= count; i++)
                {
                    dynamic effect = mainSequence.Item(i);
                    try
                    {
                        dynamic effectShape = effect.Shape;
                        int effectType = Convert.ToInt32(effect.EffectType);
                        int triggerType = Convert.ToInt32(effect.Timing.TriggerType);
                        float duration = 0;
                        float delay = 0;
                        try { duration = (float)effect.Timing.Duration; } catch { }
                        try { delay = (float)effect.Timing.TriggerDelayTime; } catch { }

                        result.Animations.Add(new AnimationInfo
                        {
                            Index = i,
                            ShapeId = (int)effectShape.Id,
                            ShapeName = effectShape.Name?.ToString() ?? "",
                            EffectType = GetEffectTypeName(effectType),
                            Timing = GetTriggerTypeName(triggerType),
                            Duration = duration,
                            Delay = delay
                        });
                        ComUtilities.Release(ref effectShape!);
                    }
                    finally
                    {
                        ComUtilities.Release(ref effect!);
                    }
                }

                return result;
            }
            finally
            {
                if (mainSequence != null) ComUtilities.Release(ref mainSequence!);
                if (timeline != null) ComUtilities.Release(ref timeline!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Add(IPptBatch batch, int slideIndex, string shapeName, int effectType, int triggerType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? timeline = null;
            dynamic? mainSequence = null;
            dynamic? effect = null;
            try
            {
                timeline = slide.TimeLine;
                mainSequence = timeline.MainSequence;
                // AddEffect(Shape, effectId, level=0, trigger, index=-1)
                effect = mainSequence.AddEffect(shape, effectType, 0, triggerType, -1);

                return new OperationResult
                {
                    Success = true,
                    Action = "add",
                    Message = $"Added animation effect {GetEffectTypeName(effectType)} to shape '{shapeName}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (effect != null) ComUtilities.Release(ref effect!);
                if (mainSequence != null) ComUtilities.Release(ref mainSequence!);
                if (timeline != null) ComUtilities.Release(ref timeline!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Remove(IPptBatch batch, int slideIndex, int effectIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic? timeline = null;
            dynamic? mainSequence = null;
            dynamic? effect = null;
            try
            {
                timeline = slide.TimeLine;
                mainSequence = timeline.MainSequence;
                effect = mainSequence.Item(effectIndex);
                effect.Delete();

                return new OperationResult
                {
                    Success = true,
                    Action = "remove",
                    Message = $"Removed animation effect {effectIndex} from slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (effect != null) ComUtilities.Release(ref effect!);
                if (mainSequence != null) ComUtilities.Release(ref mainSequence!);
                if (timeline != null) ComUtilities.Release(ref timeline!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Clear(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic? timeline = null;
            dynamic? mainSequence = null;
            try
            {
                timeline = slide.TimeLine;
                mainSequence = timeline.MainSequence;
                int count = (int)mainSequence.Count;
                // Delete in reverse order to avoid index shifting
                for (int i = count; i >= 1; i--)
                {
                    dynamic effect = mainSequence.Item(i);
                    effect.Delete();
                    ComUtilities.Release(ref effect!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "clear",
                    Message = $"Cleared {count} animation effects from slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (mainSequence != null) ComUtilities.Release(ref mainSequence!);
                if (timeline != null) ComUtilities.Release(ref timeline!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    private static string GetEffectTypeName(int effectType) => effectType switch
    {
        1 => "Appear",
        2 => "Fly",
        3 => "Blinds",
        10 => "Fade",
        16 => "Wipe",
        22 => "RandomBars",
        26 => "Dissolve",
        53 => "GrowShrink",
        55 => "Spin",
        _ => $"Effect({effectType})"
    };

    private static string GetTriggerTypeName(int triggerType) => triggerType switch
    {
        1 => "OnClick",
        2 => "WithPrevious",
        3 => "AfterPrevious",
        _ => $"Trigger({triggerType})"
    };
}
