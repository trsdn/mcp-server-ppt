using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Placeholder;

public class PlaceholderCommands : IPlaceholderCommands
{
    public PlaceholderListResult List(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                var result = new PlaceholderListResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    SlideIndex = slideIndex
                };

                dynamic placeholders = slide.Shapes.Placeholders;
                try
                {
                    int count = (int)placeholders.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        dynamic ph = placeholders.Item(i);
                        try
                        {
                            int phType = Convert.ToInt32(ph.PlaceholderFormat.Type);
                            var info = new PlaceholderInfo
                            {
                                Index = i,
                                Name = ph.Name?.ToString() ?? "",
                                PlaceholderType = phType,
                                PlaceholderTypeName = GetPlaceholderTypeName(phType),
                            };

                            try
                            {
                                info.HasTextFrame = Convert.ToInt32(ph.HasTextFrame) != 0;
                                if (info.HasTextFrame)
                                {
                                    info.Text = ph.TextFrame.TextRange.Text?.ToString();
                                }
                            }
                            catch { info.HasTextFrame = false; }

                            result.Placeholders.Add(info);
                        }
                        finally
                        {
                            ComUtilities.Release(ref ph!);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref placeholders!);
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetText(IPptBatch batch, int slideIndex, int placeholderIndex, string text)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic? ph = null;
            try
            {
                ph = slide.Shapes.Placeholders.Item(placeholderIndex);
                if (Convert.ToInt32(ph.HasTextFrame) == 0)
                    throw new InvalidOperationException($"Placeholder {placeholderIndex} on slide {slideIndex} does not have a text frame.");

                ph.TextFrame.TextRange.Text = text;
                return new OperationResult
                {
                    Success = true,
                    Action = "set-text",
                    Message = $"Set text on placeholder {placeholderIndex} (slide {slideIndex})",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (ph != null) ComUtilities.Release(ref ph!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetImage(IPptBatch batch, int slideIndex, int placeholderIndex, string imagePath)
    {
        if (!System.IO.File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic? ph = null;
            try
            {
                ph = slide.Shapes.Placeholders.Item(placeholderIndex);

                // Capture placeholder position and size
                float left = (float)ph.Left;
                float top = (float)ph.Top;
                float width = (float)ph.Width;
                float height = (float)ph.Height;

                // Delete the placeholder
                ph.Delete();
                ComUtilities.Release(ref ph!);
                ph = null;

                // Insert picture at the same position
                // AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
                // msoFalse = 0, msoTrue = -1
                dynamic pic = slide.Shapes.AddPicture(imagePath, 0, -1, left, top, width, height);
                ComUtilities.Release(ref pic!);

                return new OperationResult
                {
                    Success = true,
                    Action = "set-image",
                    Message = $"Replaced placeholder {placeholderIndex} on slide {slideIndex} with image",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (ph != null) ComUtilities.Release(ref ph!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    private static string GetPlaceholderTypeName(int ppPlaceholderType) => ppPlaceholderType switch
    {
        1 => "Title",
        2 => "Body",
        3 => "CenterTitle",
        4 => "Subtitle",
        5 => "DateAndTime",
        6 => "SlideNumber",
        7 => "Footer",
        8 => "Header",
        9 => "Object",
        10 => "Chart",
        11 => "OrgChart",
        12 => "Table",
        13 => "MediaClip",
        14 => "Bitmap",
        15 => "VerticalTitle",
        16 => "VerticalBody",
        _ => $"Unknown({ppPlaceholderType})"
    };
}
