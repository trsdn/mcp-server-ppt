using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.PrintOptions;

public class PrintOptionsCommands : IPrintOptionsCommands
{
    public OperationResult GetSettings(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic printOptions = pres.PrintOptions;
            try
            {
                int outputType = Convert.ToInt32(printOptions.OutputType);
                int colorType = Convert.ToInt32(printOptions.PrintColorType);
                bool frameSlides = Convert.ToBoolean(printOptions.FrameSlides);
                bool fitToPage = Convert.ToBoolean(printOptions.FitToPage);
                bool printHiddenSlides = Convert.ToBoolean(printOptions.PrintHiddenSlides);
                int numberOfCopies = Convert.ToInt32(printOptions.NumberOfCopies);

                return new OperationResult
                {
                    Success = true,
                    Action = "get",
                    Message = $"OutputType={outputType}, ColorType={colorType}, FrameSlides={frameSlides}, FitToPage={fitToPage}, PrintHiddenSlides={printHiddenSlides}, NumberOfCopies={numberOfCopies}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref printOptions!);
            }
        });
    }

    public OperationResult SetSettings(IPptBatch batch, int? outputType, int? colorType, bool? frameSlides, bool? fitToPage, bool? printHiddenSlides)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic printOptions = pres.PrintOptions;
            try
            {
                var changes = new List<string>();

                if (outputType.HasValue)
                {
                    printOptions.OutputType = outputType.Value;
                    changes.Add($"OutputType={outputType.Value}");
                }
                if (colorType.HasValue)
                {
                    printOptions.PrintColorType = colorType.Value;
                    changes.Add($"ColorType={colorType.Value}");
                }
                if (frameSlides.HasValue)
                {
                    printOptions.FrameSlides = frameSlides.Value ? -1 : 0;
                    changes.Add($"FrameSlides={frameSlides.Value}");
                }
                if (fitToPage.HasValue)
                {
                    printOptions.FitToPage = fitToPage.Value ? -1 : 0;
                    changes.Add($"FitToPage={fitToPage.Value}");
                }
                if (printHiddenSlides.HasValue)
                {
                    printOptions.PrintHiddenSlides = printHiddenSlides.Value ? -1 : 0;
                    changes.Add($"PrintHiddenSlides={printHiddenSlides.Value}");
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "set",
                    Message = changes.Count > 0
                        ? $"Updated print settings: {string.Join(", ", changes)}"
                        : "No print settings changed (all parameters were null)",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref printOptions!);
            }
        });
    }
}
