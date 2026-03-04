using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.SlideImport;

public class SlideImportCommands : ISlideImportCommands
{
    public OperationResult ImportSlides(IPptBatch batch, string sourceFilePath, string slideIndices, int insertAt)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceFilePath);

        return batch.Execute((ctx, ct) =>
        {
            string fullPath = Path.GetFullPath(sourceFilePath);
            if (!System.IO.File.Exists(fullPath))
                throw new FileNotFoundException($"Source file not found: '{fullPath}'");

            dynamic pres = ctx.Presentation;
            dynamic slides = pres.Slides;
            try
            {
                int currentCount = (int)slides.Count;
                int position = insertAt <= 0 ? currentCount : insertAt - 1;

                // InsertFromFile(FileName, Index, SlideStart, SlideEnd)
                if (string.IsNullOrWhiteSpace(slideIndices))
                {
                    // Import all slides
                    slides.InsertFromFile(fullPath, position);
                }
                else
                {
                    int[] indices = slideIndices.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        .Select(s => int.Parse(s, System.Globalization.CultureInfo.InvariantCulture))
                        .OrderBy(i => i)
                        .ToArray();

                    // Import slide by slide (InsertFromFile with start/end)
                    int offset = 0;
                    foreach (int idx in indices)
                    {
                        slides.InsertFromFile(fullPath, position + offset, idx, idx);
                        offset++;
                    }
                }

                int newCount = (int)slides.Count;
                int imported = newCount - currentCount;

                return new OperationResult
                {
                    Success = true,
                    Action = "import",
                    Message = $"Imported {imported} slide(s) from '{Path.GetFileName(fullPath)}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slides!);
            }
        });
    }
}
