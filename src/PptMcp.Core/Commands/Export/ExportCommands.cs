using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Export;

public class ExportCommands : IExportCommands
{
    public ExportResult ToPdf(IPptBatch batch, string destinationPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);

        return batch.Execute((ctx, ct) =>
        {
            string fullPath = Path.GetFullPath(destinationPath);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            dynamic pres = ctx.Presentation;
            // ppSaveAsPDF = 32
            pres.SaveAs(fullPath, 32);

            return new ExportResult
            {
                Success = true,
                FilePath = ctx.PresentationPath,
                OutputPath = fullPath,
                Format = "PDF"
            };
        });
    }

    public ExportResult SlideToImage(IPptBatch batch, int slideIndex, string destinationPath, int width, int height)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);

        return batch.Execute((ctx, ct) =>
        {
            string fullPath = Path.GetFullPath(destinationPath);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                slide.Export(fullPath, "PNG", width > 0 ? width : 1920, height > 0 ? height : 1080);
                return new ExportResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    OutputPath = fullPath,
                    Format = "PNG"
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public ExportResult ToVideo(IPptBatch batch, string destinationPath, int defaultSlideSeconds, int resolution)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);

        return batch.Execute((ctx, ct) =>
        {
            string fullPath = Path.GetFullPath(destinationPath);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            dynamic pres = ctx.Presentation;
            int seconds = defaultSlideSeconds > 0 ? defaultSlideSeconds : 5;
            // Resolution: 1=1080p, 2=720p, 3=480p (maps to ppResolution enum)
            int res = resolution >= 1 && resolution <= 3 ? resolution : 1;

            // CreateVideo(FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertResolution, FramesPerSecond, Quality)
            pres.CreateVideo(fullPath, false, seconds, res == 1 ? 1080 : res == 2 ? 720 : 480, 30, 85);

            // Wait for video creation to complete
            int timeout = 300; // 5 minutes max
            while (Convert.ToInt32(pres.CreateVideoStatus) == 1 && timeout > 0) // ppMediaTaskStatusInProgress = 1
            {
                System.Threading.Thread.Sleep(1000);
                timeout--;
            }

            int status = Convert.ToInt32(pres.CreateVideoStatus);
            if (status == 3) // ppMediaTaskStatusFailed
                throw new InvalidOperationException("Video creation failed.");

            return new ExportResult
            {
                Success = true,
                FilePath = ctx.PresentationPath,
                OutputPath = fullPath,
                Format = "MP4"
            };
        });
    }

    public OperationResult Print(IPptBatch batch, int copies, int fromSlide, int toSlide)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            int numCopies = copies > 0 ? copies : 1;
            int from = fromSlide > 0 ? fromSlide : -1;
            int to = toSlide > 0 ? toSlide : -1;

            if (from > 0 && to > 0)
                pres.PrintOut(from, to, "", numCopies);
            else
                pres.PrintOut(1, (int)pres.Slides.Count, "", numCopies);

            return new OperationResult
            {
                Success = true,
                Action = "print",
                Message = $"Printed {numCopies} copy(ies)" +
                    (from > 0 ? $" (slides {from}-{to})" : " (all slides)"),
                FilePath = ctx.PresentationPath
            };
        });
    }
}
