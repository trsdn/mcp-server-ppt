using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Proofing;

public class ProofingCommands : IProofingCommands
{
    public OperationResult CheckSpelling(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            var words = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            dynamic slides = pres.Slides;
            try
            {
                int slideCount = (int)slides.Count;
                for (int i = 1; i <= slideCount; i++)
                {
                    dynamic slide = slides.Item(i);
                    try
                    {
                        CollectWordsFromSlide(slide, words);
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

            var sorted = words.OrderBy(w => w, StringComparer.OrdinalIgnoreCase).ToList();
            return new OperationResult
            {
                Success = true,
                Action = "check-spelling",
                Message = $"Found {sorted.Count} unique words across all slides:\n{string.Join(", ", sorted)}",
                FilePath = ctx.PresentationPath
            };
        });
    }

    public OperationResult SetLanguage(IPptBatch batch, int slideIndex, string shapeName, int languageId)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            int affected = 0;

            if (slideIndex > 0)
            {
                dynamic slide = pres.Slides.Item(slideIndex);
                try
                {
                    affected += SetLanguageOnSlide(slide, shapeName, languageId);
                }
                finally
                {
                    ComUtilities.Release(ref slide!);
                }
            }
            else
            {
                dynamic slides = pres.Slides;
                try
                {
                    int slideCount = (int)slides.Count;
                    for (int i = 1; i <= slideCount; i++)
                    {
                        dynamic slide = slides.Item(i);
                        try
                        {
                            affected += SetLanguageOnSlide(slide, shapeName, languageId);
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
            }

            string scope = slideIndex > 0 ? $"slide {slideIndex}" : "all slides";
            string shapeScope = string.IsNullOrEmpty(shapeName) ? "all shapes" : $"shape '{shapeName}'";
            return new OperationResult
            {
                Success = true,
                Action = "set-language",
                Message = $"Set language to {languageId} ({GetLanguageName(languageId)}) on {shapeScope} in {scope}. {affected} shape(s) updated.",
                FilePath = ctx.PresentationPath
            };
        });
    }

    public OperationResult GetLanguage(IPptBatch batch, int slideIndex, string shapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                if (Convert.ToInt32(shape.HasTextFrame) == 0)
                    throw new InvalidOperationException($"Shape '{shapeName}' does not have a text frame.");

                dynamic textFrame = shape.TextFrame;
                dynamic textRange = textFrame.TextRange;
                try
                {
                    int langId = Convert.ToInt32(textRange.LanguageID);
                    return new OperationResult
                    {
                        Success = true,
                        Action = "get-language",
                        Message = $"Language on shape '{shapeName}' (slide {slideIndex}): {langId} ({GetLanguageName(langId)})",
                        FilePath = ctx.PresentationPath
                    };
                }
                finally
                {
                    ComUtilities.Release(ref textRange!);
                    ComUtilities.Release(ref textFrame!);
                }
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    private static void CollectWordsFromSlide(dynamic slide, HashSet<string> words)
    {
        dynamic shapes = slide.Shapes;
        try
        {
            int count = (int)shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic shape = shapes.Item(i);
                try
                {
                    if (Convert.ToInt32(shape.HasTextFrame) != 0)
                    {
                        dynamic textRange = shape.TextFrame.TextRange;
                        try
                        {
                            string text = textRange.Text?.ToString() ?? "";
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                foreach (string word in text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
                                {
                                    words.Add(word);
                                }
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref textRange!);
                        }
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
        }
    }

    private static int SetLanguageOnSlide(dynamic slide, string shapeName, int languageId)
    {
        int affected = 0;

        if (!string.IsNullOrEmpty(shapeName))
        {
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                SetLanguageOnShape(shape, languageId);
                affected = 1;
            }
            finally
            {
                ComUtilities.Release(ref shape!);
            }
        }
        else
        {
            dynamic shapes = slide.Shapes;
            try
            {
                int count = (int)shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic shape = shapes.Item(i);
                    try
                    {
                        if (Convert.ToInt32(shape.HasTextFrame) != 0)
                        {
                            SetLanguageOnShape(shape, languageId);
                            affected++;
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
            }
        }

        return affected;
    }

    private static void SetLanguageOnShape(dynamic shape, int languageId)
    {
        if (Convert.ToInt32(shape.HasTextFrame) == 0)
            throw new InvalidOperationException($"Shape '{shape.Name}' does not have a text frame.");

        dynamic textFrame = shape.TextFrame;
        dynamic textRange = textFrame.TextRange;
        try
        {
            textRange.LanguageID = languageId;
        }
        finally
        {
            ComUtilities.Release(ref textRange!);
            ComUtilities.Release(ref textFrame!);
        }
    }

    private static string GetLanguageName(int languageId) => languageId switch
    {
        0 => "No Proofing",
        1033 => "English (US)",
        2057 => "English (UK)",
        1031 => "German",
        1036 => "French",
        1034 => "Spanish",
        1040 => "Italian",
        1041 => "Japanese",
        1042 => "Korean",
        2052 => "Chinese (Simplified)",
        1028 => "Chinese (Traditional)",
        1046 => "Portuguese (Brazil)",
        2070 => "Portuguese (Portugal)",
        1049 => "Russian",
        1025 => "Arabic",
        1037 => "Hebrew",
        1043 => "Dutch",
        1053 => "Swedish",
        1044 => "Norwegian (Bokmal)",
        1045 => "Polish",
        1055 => "Turkish",
        _ => $"MsoLanguageID({languageId})"
    };
}
