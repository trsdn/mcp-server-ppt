using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Media;

public class MediaCommands : IMediaCommands
{
    public OperationResult InsertAudio(IPptBatch batch, int slideIndex, string filePath, float left, float top, bool linkToFile, bool saveWithDocument)
    {
        return batch.Execute((ctx, ct) =>
        {
            if (!System.IO.File.Exists(filePath))
                throw new FileNotFoundException($"Audio file not found: {filePath}");

            dynamic slide = ComUtilities.GetSlide(ctx.Presentation, slideIndex);
            dynamic? shapes = null;
            dynamic? shape = null;
            try
            {
                // AddMediaObject2(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
                // -1 = msoTrue, 0 = msoFalse
                int link = linkToFile ? -1 : 0;
                int saveWith = saveWithDocument ? -1 : 0;
                shapes = slide.Shapes;
                shape = shapes.AddMediaObject2(filePath, link, saveWith, left, top);
                string name = shape.Name?.ToString() ?? "";

                return new OperationResult
                {
                    Success = true,
                    Action = "insert-audio",
                    Message = $"Inserted audio '{System.IO.Path.GetFileName(filePath)}' as '{name}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref shapes!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult InsertVideo(IPptBatch batch, int slideIndex, string filePath, float left, float top, float width, float height, bool linkToFile)
    {
        return batch.Execute((ctx, ct) =>
        {
            if (!System.IO.File.Exists(filePath))
                throw new FileNotFoundException($"Video file not found: {filePath}");

            dynamic slide = ComUtilities.GetSlide(ctx.Presentation, slideIndex);
            dynamic? shapes = null;
            dynamic? shape = null;
            try
            {
                int link = linkToFile ? -1 : 0;
                // saveWithDocument = msoTrue (-1) by default for videos
                shapes = slide.Shapes;
                shape = shapes.AddMediaObject2(filePath, link, -1, left, top, width > 0 ? width : -1, height > 0 ? height : -1);
                string name = shape.Name?.ToString() ?? "";

                return new OperationResult
                {
                    Success = true,
                    Action = "insert-video",
                    Message = $"Inserted video '{System.IO.Path.GetFileName(filePath)}' as '{name}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref shapes!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public MediaInfoResult GetInfo(IPptBatch batch, int slideIndex, string shapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ComUtilities.GetSlide(ctx.Presentation, slideIndex);
            dynamic shape = ComUtilities.GetShape(slide, shapeName);
            try
            {
                // ppMediaTypeMovie = 3, ppMediaTypeSound = 1
                int mediaType = 0;
                string mediaTypeName = "Unknown";
                string sourceFile = "";

                try
                {
                    mediaType = Convert.ToInt32(shape.MediaType);
                    mediaTypeName = mediaType switch
                    {
                        1 => "Audio",
                        2 => "Other",
                        3 => "Video",
                        _ => $"Unknown({mediaType})"
                    };
                }
                catch (Exception ex)
                {
                    // shape.MediaType is only valid on a media shape. Absorbing the
                    // failure produced MediaType = "Unknown" under Success = true, so
                    // asking for media details about a rectangle returned a confident
                    // wrong answer instead of an error (#126).
                    throw new InvalidOperationException(
                        $"Shape '{shapeName}' on slide {slideIndex} is not a media shape.", ex);
                }

                try { sourceFile = shape.LinkFormat?.SourceFullName?.ToString() ?? ""; } catch { }

                return new MediaInfoResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    SlideIndex = slideIndex,
                    ShapeName = shapeName,
                    MediaType = mediaTypeName,
                    SourceFile = sourceFile,
                    Left = Convert.ToSingle(shape.Left),
                    Top = Convert.ToSingle(shape.Top),
                    Width = Convert.ToSingle(shape.Width),
                    Height = Convert.ToSingle(shape.Height)
                };
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetPlayback(IPptBatch batch, int slideIndex, string shapeName, float? volume, bool? muted, float? fadeInSeconds, float? fadeOutSeconds)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ComUtilities.GetSlide(ctx.Presentation, slideIndex);
            dynamic shape = ComUtilities.GetShape(slide, shapeName);
            try
            {
                dynamic mediaFormat = shape.MediaFormat;
                try
                {
                    if (volume.HasValue) mediaFormat.Volume = volume.Value;
                    if (muted.HasValue) mediaFormat.Muted = muted.Value ? -1 : 0;
                    if (fadeInSeconds.HasValue) mediaFormat.FadeInDuration = (int)(fadeInSeconds.Value * 1000);
                    if (fadeOutSeconds.HasValue) mediaFormat.FadeOutDuration = (int)(fadeOutSeconds.Value * 1000);

                    return new OperationResult
                    {
                        Success = true,
                        Action = "set-playback",
                        Message = $"Set playback properties on media shape '{shapeName}' on slide {slideIndex}",
                        FilePath = ctx.PresentationPath
                    };
                }
                finally
                {
                    ComUtilities.Release(ref mediaFormat!);
                }
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }
}
