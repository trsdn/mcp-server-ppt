using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Tag;

public class TagCommands : ITagCommands
{
    public TagListResult List(IPptBatch batch, int slideIndex, string? shapeName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                dynamic tags = GetTags(slide, shapeName);
                try
                {
                    var result = new TagListResult
                    {
                        Success = true,
                        FilePath = ctx.PresentationPath,
                        SlideIndex = slideIndex,
                        ShapeName = shapeName
                    };

                    int count = (int)tags.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        result.Tags.Add(new TagInfo
                        {
                            Name = tags.Name(i)?.ToString() ?? "",
                            Value = tags.Value(i)?.ToString() ?? ""
                        });
                    }

                    return result;
                }
                finally
                {
                    ComUtilities.Release(ref tags!);
                }
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetTag(IPptBatch batch, int slideIndex, string? shapeName, string tagName, string tagValue)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(tagName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                dynamic tags = GetTags(slide, shapeName);
                try
                {
                    tags.Add(tagName, tagValue);
                }
                finally
                {
                    ComUtilities.Release(ref tags!);
                }

                string target = string.IsNullOrEmpty(shapeName)
                    ? $"slide {slideIndex}"
                    : $"shape '{shapeName}' on slide {slideIndex}";

                return new OperationResult
                {
                    Success = true,
                    Action = "set",
                    Message = $"Set tag '{tagName}' = '{tagValue}' on {target}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult DeleteTag(IPptBatch batch, int slideIndex, string? shapeName, string tagName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(tagName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                dynamic tags = GetTags(slide, shapeName);
                try
                {
                    tags.Delete(tagName);
                }
                finally
                {
                    ComUtilities.Release(ref tags!);
                }

                string target = string.IsNullOrEmpty(shapeName)
                    ? $"slide {slideIndex}"
                    : $"shape '{shapeName}' on slide {slideIndex}";

                return new OperationResult
                {
                    Success = true,
                    Action = "delete",
                    Message = $"Deleted tag '{tagName}' from {target}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    private static dynamic GetTags(dynamic slide, string? shapeName)
    {
        if (string.IsNullOrWhiteSpace(shapeName))
            return slide.Tags;

        dynamic shape = slide.Shapes.Item(shapeName);
        try
        {
            dynamic tags = shape.Tags;
            ComUtilities.Release(ref shape!);
            return tags;
        }
        catch
        {
            ComUtilities.Release(ref shape!);
            throw;
        }
    }
}
