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

    public OperationResult SetImage(IPptBatch batch, int slideIndex, string imagePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(imagePath);

        return batch.Execute((ctx, ct) =>
        {
            string fullPath = Path.GetFullPath(imagePath);
            if (!System.IO.File.Exists(fullPath))
                throw new FileNotFoundException($"Image file not found: '{fullPath}'");

            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                slide.FollowMasterBackground = 0; // msoFalse
                slide.Background.Fill.UserPicture(fullPath);

                return new OperationResult
                {
                    Success = true,
                    Action = "set-image",
                    Message = $"Set background image of slide {slideIndex} to '{Path.GetFileName(fullPath)}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetGradient(IPptBatch batch, int slideIndex, string color1, string color2, int gradientStyle)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(color1);
        ArgumentException.ThrowIfNullOrWhiteSpace(color2);

        if (gradientStyle < 1 || gradientStyle > 6)
            throw new ArgumentOutOfRangeException(nameof(gradientStyle), "gradientStyle must be 1-6 (1=Horizontal, 2=Vertical, 3=DiagonalUp, 4=DiagonalDown, 5=FromCorner, 6=FromCenter)");

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                slide.FollowMasterBackground = 0; // msoFalse
                slide.Background.Fill.TwoColorGradient(gradientStyle, 1);
                slide.Background.Fill.ForeColor.RGB = HexToOleColor(color1);
                slide.Background.Fill.BackColor.RGB = HexToOleColor(color2);

                return new OperationResult
                {
                    Success = true,
                    Action = "set-gradient",
                    Message = $"Set gradient background on slide {slideIndex} from '{color1}' to '{color2}' (style {gradientStyle})",
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
