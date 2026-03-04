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
}
