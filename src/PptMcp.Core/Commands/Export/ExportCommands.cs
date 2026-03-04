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

    public ExportResult SaveAs(IPptBatch batch, string destinationPath, int format)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);

        return batch.Execute((ctx, ct) =>
        {
            string fullPath = Path.GetFullPath(destinationPath);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            // ppSaveAsDefault=11, ppSaveAsOpenXMLPresentation=24, ppSaveAsOpenXMLPresentationMacroEnabled=25,
            // ppSaveAsTemplate=5, ppSaveAsOpenXMLShow=28, ppSaveAsPDF=32, ppSaveAsXPS=33, ppSaveAsODP=37
            int ppFormat = format switch
            {
                1 => 24, // pptx
                2 => 25, // pptm
                3 => 5,  // potx (template)
                4 => 28, // ppsx (show)
                5 => 32, // pdf
                6 => 33, // xps
                7 => 37, // odp
                _ => 24  // default to pptx
            };

            string formatName = format switch
            {
                1 => "PPTX",
                2 => "PPTM",
                3 => "POTX",
                4 => "PPSX",
                5 => "PDF",
                6 => "XPS",
                7 => "ODP",
                _ => "PPTX"
            };

            dynamic pres = ctx.Presentation;
            pres.SaveAs(fullPath, ppFormat);

            return new ExportResult
            {
                Success = true,
                FilePath = ctx.PresentationPath,
                OutputPath = fullPath,
                Format = formatName
            };
        });
    }

    public ExportResult AllSlidesToImages(IPptBatch batch, string destinationDirectory, int width, int height)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationDirectory);

        return batch.Execute((ctx, ct) =>
        {
            string fullDir = Path.GetFullPath(destinationDirectory);
            if (!Directory.Exists(fullDir))
                Directory.CreateDirectory(fullDir);

            dynamic slides = ((dynamic)ctx.Presentation).Slides;
            try
            {
                int count = (int)slides.Count;
                int w = width > 0 ? width : 1920;
                int h = height > 0 ? height : 1080;

                for (int i = 1; i <= count; i++)
                {
                    dynamic slide = slides.Item(i);
                    try
                    {
                        string fileName = $"slide_{i:D3}.png";
                        string filePath = Path.Combine(fullDir, fileName);
                        slide.Export(filePath, "PNG", w, h);
                    }
                    finally
                    {
                        ComUtilities.Release(ref slide!);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref slides!);
            }

            return new ExportResult
            {
                Success = true,
                FilePath = ctx.PresentationPath,
                OutputPath = fullDir,
                Format = "PNG"
            };
        });
    }

    public OperationResult ExtractText(IPptBatch batch, string destinationPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);

        return batch.Execute((ctx, ct) =>
        {
            string fullPath = Path.GetFullPath(destinationPath);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            dynamic slides = ((dynamic)ctx.Presentation).Slides;
            try
            {
                int slideCount = (int)slides.Count;
                using var writer = new StreamWriter(fullPath, false, System.Text.Encoding.UTF8);

                for (int i = 1; i <= slideCount; i++)
                {
                    dynamic slide = slides.Item(i);
                    dynamic shapes = slide.Shapes;
                    try
                    {
                        writer.WriteLine($"=== Slide {i} ===");
                        int shapeCount = (int)shapes.Count;

                        for (int j = 1; j <= shapeCount; j++)
                        {
                            dynamic shape = shapes.Item(j);
                            try
                            {
                                if ((bool)shape.HasTextFrame)
                                {
                                    dynamic textFrame = shape.TextFrame;
                                    dynamic textRange = textFrame.TextRange;
                                    try
                                    {
                                        string text = textRange.Text?.ToString() ?? "";
                                        if (!string.IsNullOrWhiteSpace(text))
                                            writer.WriteLine(text);
                                    }
                                    finally
                                    {
                                        ComUtilities.Release(ref textRange!);
                                        ComUtilities.Release(ref textFrame!);
                                    }
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref shape!);
                            }
                        }

                        writer.WriteLine();
                    }
                    finally
                    {
                        ComUtilities.Release(ref shapes!);
                        ComUtilities.Release(ref slide!);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref slides!);
            }

            return new OperationResult
            {
                Success = true,
                Action = "extract-text",
                Message = $"Extracted text to '{fullPath}'",
                FilePath = ctx.PresentationPath
            };
        });
    }

    public OperationResult ExtractImages(IPptBatch batch, string destinationDirectory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationDirectory);

        return batch.Execute((ctx, ct) =>
        {
            string fullDir = Path.GetFullPath(destinationDirectory);
            if (!Directory.Exists(fullDir))
                Directory.CreateDirectory(fullDir);

            dynamic slides = ((dynamic)ctx.Presentation).Slides;
            try
            {
                int slideCount = (int)slides.Count;
                int imageCount = 0;

                for (int i = 1; i <= slideCount; i++)
                {
                    dynamic slide = slides.Item(i);
                    dynamic shapes = slide.Shapes;
                    try
                    {
                        int shapeCount = (int)shapes.Count;
                        for (int j = 1; j <= shapeCount; j++)
                        {
                            dynamic shape = shapes.Item(j);
                            try
                            {
                                int shapeType = Convert.ToInt32(shape.Type);
                                // msoPicture=13, msoLinkedPicture=11
                                if (shapeType == 13 || shapeType == 11)
                                {
                                    imageCount++;
                                    string fileName = $"slide{i:D3}_image{imageCount:D3}.png";
                                    string filePath = Path.Combine(fullDir, fileName);
                                    // ppShapeFormatPNG = 2
                                    shape.Export(filePath, 2);
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref shape!);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref shapes!);
                        ComUtilities.Release(ref slide!);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref slides!);
            }

            return new OperationResult
            {
                Success = true,
                Action = "extract-images",
                Message = $"Extracted images to '{fullDir}'",
                FilePath = ctx.PresentationPath
            };
        });
    }
}
