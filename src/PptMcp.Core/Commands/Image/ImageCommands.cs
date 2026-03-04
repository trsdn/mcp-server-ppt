using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Image;

public class ImageCommands : IImageCommands
{
    public OperationResult Insert(IPptBatch batch, int slideIndex, string imagePath, float left, float top, float width, float height)
    {
        return batch.Execute((ctx, ct) =>
        {
            string fullImagePath = Path.GetFullPath(imagePath);
            if (!System.IO.File.Exists(fullImagePath))
                throw new FileNotFoundException($"Image file not found: '{fullImagePath}'");

            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                // AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
                // msoFalse=0, msoTrue=-1
                dynamic shape = (width > 0 && height > 0)
                    ? slide.Shapes.AddPicture(fullImagePath, 0, -1, left, top, width, height)
                    : slide.Shapes.AddPicture(fullImagePath, 0, -1, left, top);

                string name = shape.Name?.ToString() ?? "";
                ComUtilities.Release(ref shape!);

                return new OperationResult
                {
                    Success = true,
                    Action = "insert",
                    Message = $"Inserted image '{Path.GetFileName(fullImagePath)}' as '{name}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Crop(IPptBatch batch, int slideIndex, string shapeName, float cropLeft, float cropRight, float cropTop, float cropBottom)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                dynamic picFormat = shape.PictureFormat;
                try
                {
                    picFormat.CropLeft = cropLeft;
                    picFormat.CropRight = cropRight;
                    picFormat.CropTop = cropTop;
                    picFormat.CropBottom = cropBottom;
                }
                finally
                {
                    ComUtilities.Release(ref picFormat!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "crop",
                    Message = $"Cropped image '{shapeName}' on slide {slideIndex} (L:{cropLeft}, R:{cropRight}, T:{cropTop}, B:{cropBottom})",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetBrightnessContrast(IPptBatch batch, int slideIndex, string shapeName, float brightness, float contrast)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                dynamic picFormat = shape.PictureFormat;
                try
                {
                    picFormat.Brightness = brightness;
                    picFormat.Contrast = contrast;
                }
                finally
                {
                    ComUtilities.Release(ref picFormat!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "set-brightness-contrast",
                    Message = $"Set brightness={brightness:F2}, contrast={contrast:F2} on image '{shapeName}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetTransparentColor(IPptBatch batch, int slideIndex, string shapeName, string colorHex)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);
        ArgumentException.ThrowIfNullOrWhiteSpace(colorHex);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                dynamic picFormat = shape.PictureFormat;
                try
                {
                    picFormat.TransparencyColor = HexToOleColor(colorHex);
                    // msoTrue = -1 enables transparent background
                    picFormat.TransparentBackground = -1;
                }
                finally
                {
                    ComUtilities.Release(ref picFormat!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "set-transparent-color",
                    Message = $"Set transparent color '{colorHex}' on image '{shapeName}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

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
