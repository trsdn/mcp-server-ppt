using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Text;

public class TextCommands : ITextCommands
{
    public TextResult GetText(IPptBatch batch, int slideIndex, string shapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                var result = new TextResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    ShapeId = (int)shape.Id,
                    ShapeName = shape.Name?.ToString() ?? ""
                };

                if (Convert.ToInt32(shape.HasTextFrame) == 0)
                {
                    result.Text = "";
                    return result;
                }

                dynamic textFrame = shape.TextFrame;
                dynamic textRange = textFrame.TextRange;
                try
                {
                    result.Text = textRange.Text?.ToString() ?? "";

                    ReadParagraphs(textRange, result.Paragraphs);

                    return result;
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

    public OperationResult SetText(IPptBatch batch, int slideIndex, string shapeName, string text)
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

                shape.TextFrame.TextRange.Text = text;
                return new OperationResult
                {
                    Success = true,
                    Action = "set",
                    Message = $"Set text on shape '{shapeName}' (slide {slideIndex})",
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

    public OperationResult Find(IPptBatch batch, string searchText, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            var matches = new List<string>();

            void SearchSlide(dynamic s, int idx)
            {
                dynamic shapes = s.Shapes;
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
                                string text = shape.TextFrame.TextRange.Text?.ToString() ?? "";
                                if (text.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                                {
                                    matches.Add($"Slide {idx}, Shape '{shape.Name}': found '{searchText}'");
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

            if (slideIndex > 0)
            {
                dynamic slide = pres.Slides.Item(slideIndex);
                try
                {
                    SearchSlide(slide, slideIndex);
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
                            SearchSlide(slide, i);
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

            return new OperationResult
            {
                Success = true,
                Action = "find",
                Message = matches.Count > 0
                    ? $"Found {matches.Count} match(es):\n" + string.Join("\n", matches)
                    : $"No matches found for '{searchText}'",
                FilePath = ctx.PresentationPath
            };
        });
    }

    public OperationResult Replace(IPptBatch batch, string searchText, string replaceText, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            int replacements = 0;

            void ReplaceInSlide(dynamic s)
            {
                dynamic shapes = s.Shapes;
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
                                    if (text.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                                    {
                                        // Use Replace method via TextRange
                                        dynamic found = textRange.Find(searchText);
                                        while (found != null && Convert.ToInt32(found.Length) > 0)
                                        {
                                            found.Text = replaceText;
                                            replacements++;
                                            try
                                            {
                                                found = textRange.Find(searchText);
                                            }
                                            catch
                                            {
                                                break;
                                            }
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

            if (slideIndex > 0)
            {
                dynamic slide = pres.Slides.Item(slideIndex);
                try
                {
                    ReplaceInSlide(slide);
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
                            ReplaceInSlide(slide);
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

            return new OperationResult
            {
                Success = true,
                Action = "replace",
                Message = $"Replaced {replacements} occurrence(s) of '{searchText}' with '{replaceText}'",
                FilePath = ctx.PresentationPath
            };
        });
    }

    public OperationResult Format(IPptBatch batch, int slideIndex, string shapeName, string? fontName, float? fontSize, bool? bold, bool? italic, string? color, string? alignment, string? verticalAlignment)
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
                dynamic font = textFrame.TextRange.Font;
                if (fontName != null) font.Name = fontName;
                if (fontSize.HasValue) font.Size = fontSize.Value;
                if (bold.HasValue) font.Bold = bold.Value ? -1 : 0; // msoTrue/msoFalse
                if (italic.HasValue) font.Italic = italic.Value ? -1 : 0;
                if (color != null)
                {
                    // Parse hex color #RRGGBB → RGB int
                    if (color.StartsWith('#') && color.Length == 7)
                    {
                        int r = Convert.ToInt32(color[1..3], 16);
                        int g = Convert.ToInt32(color[3..5], 16);
                        int b = Convert.ToInt32(color[5..7], 16);
                        font.Color.RGB = r + (g << 8) + (b << 16); // PowerPoint uses BGR format
                    }
                }

                // Horizontal alignment for all paragraphs
                if (alignment != null)
                {
                    // ppAlignLeft=1, ppAlignCenter=2, ppAlignRight=3, ppAlignJustify=4
                    int ppAlign = alignment.ToLowerInvariant() switch
                    {
                        "left" => 1,
                        "center" => 2,
                        "right" => 3,
                        "justify" => 4,
                        _ => 1
                    };
                    dynamic paragraphs = textFrame.TextRange.Paragraphs();
                    int paraCount = (int)paragraphs.Count;
                    for (int p = 1; p <= paraCount; p++)
                    {
                        dynamic para = textFrame.TextRange.Paragraphs(p, 1);
                        try { para.ParagraphFormat.Alignment = ppAlign; }
                        finally { ComUtilities.Release(ref para!); }
                    }
                    ComUtilities.Release(ref paragraphs!);
                }

                // Vertical anchor: msoAnchorTop=1, msoAnchorMiddle=3, msoAnchorBottom=4
                if (verticalAlignment != null)
                {
                    textFrame.VerticalAnchor = verticalAlignment.ToLowerInvariant() switch
                    {
                        "top" => 1,
                        "middle" => 3,
                        "bottom" => 4,
                        _ => 1
                    };
                }

                ComUtilities.Release(ref font!);
                ComUtilities.Release(ref textFrame!);
                return new OperationResult
                {
                    Success = true,
                    Action = "format",
                    Message = $"Formatted text in shape '{shapeName}' (slide {slideIndex})",
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

    public OperationResult FormatAdvanced(IPptBatch batch, int slideIndex, string shapeName, bool? underline, bool? strikethrough, bool? subscript, bool? superscript)
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

                dynamic font = shape.TextFrame.TextRange.Font;
                try
                {
                    if (underline.HasValue)
                        font.Underline = underline.Value ? -1 : 0;
                    if (strikethrough.HasValue)
                        font.Strikethrough = strikethrough.Value ? -1 : 0;
                    if (subscript.HasValue)
                        font.Subscript = subscript.Value ? -1 : 0;
                    if (superscript.HasValue)
                        font.Superscript = superscript.Value ? -1 : 0;
                }
                finally
                {
                    ComUtilities.Release(ref font!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "format-advanced",
                    Message = $"Applied advanced formatting to shape '{shapeName}' (slide {slideIndex})",
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

    public OperationResult WordCount(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            int totalWords = 0;

            void CountInSlide(dynamic s)
            {
                dynamic shapes = s.Shapes;
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
                                        totalWords += text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
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

            if (slideIndex > 0)
            {
                dynamic slide = pres.Slides.Item(slideIndex);
                try
                {
                    CountInSlide(slide);
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
                            CountInSlide(slide);
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
            return new OperationResult
            {
                Success = true,
                Action = "word-count",
                Message = $"Total word count ({scope}): {totalWords}",
                FilePath = ctx.PresentationPath
            };
        });
    }

    public OperationResult AltTextAudit(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            var missing = new List<string>();

            void AuditSlide(dynamic s, int idx)
            {
                dynamic shapes = s.Shapes;
                try
                {
                    int count = (int)shapes.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        dynamic shape = shapes.Item(i);
                        try
                        {
                            string altText = shape.AlternativeText?.ToString() ?? "";
                            if (string.IsNullOrWhiteSpace(altText))
                            {
                                missing.Add($"Slide {idx}, Shape '{shape.Name}'");
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

            if (slideIndex > 0)
            {
                dynamic slide = pres.Slides.Item(slideIndex);
                try
                {
                    AuditSlide(slide, slideIndex);
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
                            AuditSlide(slide, i);
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

            return new OperationResult
            {
                Success = true,
                Action = "alt-text-audit",
                Message = missing.Count > 0
                    ? $"{missing.Count} shape(s) missing alt text:\n" + string.Join("\n", missing)
                    : "All shapes have alt text.",
                FilePath = ctx.PresentationPath
            };
        });
    }

    public OperationResult EmptyPlaceholderAudit(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            var empty = new List<string>();

            void AuditSlide(dynamic s, int idx)
            {
                dynamic placeholders = s.Shapes.Placeholders;
                try
                {
                    int count = (int)placeholders.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        dynamic ph = placeholders.Item(i);
                        try
                        {
                            if (Convert.ToInt32(ph.HasTextFrame) != 0)
                            {
                                dynamic textRange = ph.TextFrame.TextRange;
                                try
                                {
                                    string text = textRange.Text?.ToString() ?? "";
                                    if (string.IsNullOrWhiteSpace(text))
                                    {
                                        empty.Add($"Slide {idx}, Placeholder '{ph.Name}'");
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
                            ComUtilities.Release(ref ph!);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref placeholders!);
                }
            }

            if (slideIndex > 0)
            {
                dynamic slide = pres.Slides.Item(slideIndex);
                try
                {
                    AuditSlide(slide, slideIndex);
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
                            AuditSlide(slide, i);
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

            return new OperationResult
            {
                Success = true,
                Action = "empty-placeholder-audit",
                Message = empty.Count > 0
                    ? $"{empty.Count} empty placeholder(s) found:\n" + string.Join("\n", empty)
                    : "No empty placeholders found.",
                FilePath = ctx.PresentationPath
            };
        });
    }

    public OperationResult SetSpacing(IPptBatch batch, int slideIndex, string shapeName, float? lineSpacing, float? spaceBefore, float? spaceAfter, float? characterSpacing)
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
                    // Paragraph-level spacing
                    dynamic paragraphFormat = textRange.ParagraphFormat;
                    try
                    {
                        if (lineSpacing.HasValue) paragraphFormat.SpaceWithin = lineSpacing.Value;
                        if (spaceBefore.HasValue) paragraphFormat.SpaceBefore = spaceBefore.Value;
                        if (spaceAfter.HasValue) paragraphFormat.SpaceAfter = spaceAfter.Value;
                    }
                    finally
                    {
                        ComUtilities.Release(ref paragraphFormat!);
                    }

                    // Character-level spacing
                    if (characterSpacing.HasValue)
                    {
                        dynamic font = textRange.Font;
                        try
                        {
                            font.Spacing = characterSpacing.Value;
                        }
                        finally
                        {
                            ComUtilities.Release(ref font!);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref textRange!);
                    ComUtilities.Release(ref textFrame!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "set-spacing",
                    Message = $"Set spacing on shape '{shapeName}' (slide {slideIndex})",
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

    public OperationResult SetBullets(IPptBatch batch, int slideIndex, string shapeName, int bulletType, string? bulletCharacter, int indentLevel)
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
                    dynamic paragraphFormat = textRange.ParagraphFormat;
                    try
                    {
                        // ppBulletNone=0, ppBulletUnnumbered=1, ppBulletNumbered=2
                        dynamic bullet = paragraphFormat.Bullet;
                        try
                        {
                            bullet.Type = bulletType;

                            if (bulletType == 1 && !string.IsNullOrEmpty(bulletCharacter))
                                bullet.Character = Convert.ToInt32(bulletCharacter[0]);
                        }
                        finally
                        {
                            ComUtilities.Release(ref bullet!);
                        }

                        // ParagraphFormat.Level is 1-based (1-5)
                        paragraphFormat.Level = indentLevel + 1;
                    }
                    finally
                    {
                        ComUtilities.Release(ref paragraphFormat!);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref textRange!);
                    ComUtilities.Release(ref textFrame!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "set-bullets",
                    Message = $"Set bullets on shape '{shapeName}' (slide {slideIndex})",
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

    /// <summary>
    /// Read paragraph and run details from a COM TextRange into the provided list.
    /// </summary>
    private static void ReadParagraphs(dynamic textRange, List<TextParagraphInfo> paragraphs)
    {
        dynamic allParagraphs = textRange.Paragraphs();
        try
        {
            int paraCount = (int)allParagraphs.Count;
            for (int p = 1; p <= paraCount; p++)
            {
                dynamic para = textRange.Paragraphs(p, 1);
                try
                {
                    var paraInfo = new TextParagraphInfo
                    {
                        Index = p,
                        Text = para.Text?.ToString() ?? ""
                    };

                    try { paraInfo.Alignment = Convert.ToInt32(para.ParagraphFormat.Alignment); } catch { }

                    dynamic runs = para.Runs();
                    try
                    {
                        int runCount = (int)runs.Count;
                        for (int r = 1; r <= runCount; r++)
                        {
                            dynamic run = para.Runs(r, 1);
                            try
                            {
                                var runInfo = new TextRunInfo
                                {
                                    Text = run.Text?.ToString() ?? ""
                                };
                                try { runInfo.FontName = run.Font.Name?.ToString(); } catch { }
                                try { runInfo.FontSize = Convert.ToSingle(run.Font.Size); } catch { }
                                try { runInfo.Bold = Convert.ToInt32(run.Font.Bold) != 0; } catch { }
                                try { runInfo.Italic = Convert.ToInt32(run.Font.Italic) != 0; } catch { }
                                try
                                {
                                    int rgb = Convert.ToInt32(run.Font.Color.RGB);
                                    runInfo.Color = $"#{rgb:X6}";
                                }
                                catch { }

                                paraInfo.Runs.Add(runInfo);
                            }
                            finally
                            {
                                ComUtilities.Release(ref run!);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref runs!);
                    }

                    paragraphs.Add(paraInfo);
                }
                finally
                {
                    ComUtilities.Release(ref para!);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref allParagraphs!);
        }
    }
}
