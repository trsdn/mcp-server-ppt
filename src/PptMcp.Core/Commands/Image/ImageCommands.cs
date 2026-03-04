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
}
