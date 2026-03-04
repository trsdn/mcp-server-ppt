using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Background;

public class BackgroundCommands : IBackgroundCommands
{
    public BackgroundResult GetInfo(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                bool followMaster = Convert.ToInt32(slide.FollowMasterBackground) != 0;
                var result = new BackgroundResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    SlideIndex = slideIndex,
                    FollowMasterBackground = followMaster,
                };

                if (!followMaster)
                {
                    try
                    {
                        int fillType = Convert.ToInt32(slide.Background.Fill.Type);
                        result.FillType = GetFillTypeName(fillType);

                        if (fillType == 1) // msoFillSolid
                        {
                            int rgb = Convert.ToInt32(slide.Background.Fill.ForeColor.RGB);
                            result.Color = $"#{rgb:X6}";
                        }
                    }
                    catch { result.FillType = "Unknown"; }
                }
                else
                {
                    result.FillType = "Master";
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetColor(IPptBatch batch, int slideIndex, string colorHex)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(colorHex);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                slide.FollowMasterBackground = 0; // msoFalse
                slide.Background.Fill.Solid();
                slide.Background.Fill.ForeColor.RGB = HexToOleColor(colorHex);

                return new OperationResult
                {
                    Success = true,
                    Action = "set-color",
                    Message = $"Set background color of slide {slideIndex} to '{colorHex}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Reset(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                slide.FollowMasterBackground = -1; // msoTrue

                return new OperationResult
                {
                    Success = true,
                    Action = "reset",
                    Message = $"Reset background of slide {slideIndex} to master",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    private static string GetFillTypeName(int msoFillType) => msoFillType switch
    {
        1 => "Solid",
        2 => "Patterned",
        3 => "Gradient",
        4 => "Textured",
        5 => "Background",
        6 => "Picture",
        _ => $"Unknown({msoFillType})"
    };

    private static int HexToOleColor(string hex)
    {
        hex = hex.TrimStart('#');
        if (hex.Length == 3)
            hex = string.Concat(hex[0], hex[0], hex[1], hex[1], hex[2], hex[2]);
        int r = Convert.ToInt32(hex[..2], 16);
        int g = Convert.ToInt32(hex[2..4], 16);
        int b = Convert.ToInt32(hex[4..6], 16);
        return r | (g << 8) | (b << 16);
    }
}
