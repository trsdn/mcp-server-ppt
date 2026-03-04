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

    public OperationResult InsertLink(IPptBatch batch, int slideIndex, string shapeName, string linkText, string url)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);
        ArgumentException.ThrowIfNullOrWhiteSpace(linkText);
        ArgumentException.ThrowIfNullOrWhiteSpace(url);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? textFrame = null;
            dynamic? textRange = null;
            dynamic? found = null;
            dynamic? actionSettings = null;
            dynamic? actionSetting = null;
            dynamic? hyperlink = null;
            try
            {
                textFrame = shape.TextFrame;
                textRange = textFrame.TextRange;
                found = textRange.Find(linkText);

                if (found == null)
                {
                    return new OperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Text '{linkText}' not found in shape '{shapeName}' on slide {slideIndex}",
                        FilePath = ctx.PresentationPath
                    };
                }

                // ppMouseClick = 1
                actionSettings = found.ActionSettings;
                actionSetting = actionSettings.Item(1);
                hyperlink = actionSetting.Hyperlink;
                hyperlink.Address = url;

                return new OperationResult
                {
                    Success = true,
                    Action = "insert-link",
                    Message = $"Added hyperlink '{url}' to text '{linkText}' in shape '{shapeName}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (hyperlink != null) ComUtilities.Release(ref hyperlink!);
                if (actionSetting != null) ComUtilities.Release(ref actionSetting!);
                if (actionSettings != null) ComUtilities.Release(ref actionSettings!);
                if (found != null) ComUtilities.Release(ref found!);
                if (textRange != null) ComUtilities.Release(ref textRange!);
                if (textFrame != null) ComUtilities.Release(ref textFrame!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult ChangeCase(IPptBatch batch, int slideIndex, string shapeName, int caseType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? textFrame = null;
            dynamic? textRange = null;
            try
            {
                if (Convert.ToInt32(shape.HasTextFrame) == 0)
                    throw new InvalidOperationException($"Shape '{shapeName}' does not have a text frame.");

                textFrame = shape.TextFrame;
                textRange = textFrame.TextRange;
                textRange.ChangeCase(caseType);

                string caseLabel = caseType switch
                {
                    1 => "Sentence case",
                    2 => "lowercase",
                    3 => "UPPERCASE",
                    4 => "Title Case",
                    5 => "tOGGLE cASE",
                    _ => $"case type {caseType}"
                };

                return new OperationResult
                {
                    Success = true,
                    Action = "change-case",
                    Message = $"Changed text to {caseLabel} in shape '{shapeName}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (textRange != null) ComUtilities.Release(ref textRange!);
                if (textFrame != null) ComUtilities.Release(ref textFrame!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult ReadSpacing(IPptBatch batch, int slideIndex, string shapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? textFrame = null;
            dynamic? textRange = null;
            dynamic? paragraphFormat = null;
            dynamic? font = null;
            try
            {
                if (Convert.ToInt32(shape.HasTextFrame) == 0)
                    throw new InvalidOperationException($"Shape '{shapeName}' does not have a text frame.");

                textFrame = shape.TextFrame;
                textRange = textFrame.TextRange;
                paragraphFormat = textRange.ParagraphFormat;
                font = textRange.Font;

                float spaceWithin = Convert.ToSingle(paragraphFormat.SpaceWithin);
                float spaceBefore = Convert.ToSingle(paragraphFormat.SpaceBefore);
                float spaceAfter = Convert.ToSingle(paragraphFormat.SpaceAfter);
                float charSpacing = Convert.ToSingle(font.Spacing);

                return new OperationResult
                {
                    Success = true,
                    Action = "read-spacing",
                    Message = $"SpaceWithin={spaceWithin}, SpaceBefore={spaceBefore}, SpaceAfter={spaceAfter}, CharacterSpacing={charSpacing}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (font != null) ComUtilities.Release(ref font!);
                if (paragraphFormat != null) ComUtilities.Release(ref paragraphFormat!);
                if (textRange != null) ComUtilities.Release(ref textRange!);
                if (textFrame != null) ComUtilities.Release(ref textFrame!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult ReadBullets(IPptBatch batch, int slideIndex, string shapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? textFrame = null;
            dynamic? textRange = null;
            dynamic? allParagraphs = null;
            try
            {
                if (Convert.ToInt32(shape.HasTextFrame) == 0)
                    throw new InvalidOperationException($"Shape '{shapeName}' does not have a text frame.");

                textFrame = shape.TextFrame;
                textRange = textFrame.TextRange;
                allParagraphs = textRange.Paragraphs();

                int paraCount = (int)allParagraphs.Count;
                var lines = new List<string>();

                for (int p = 1; p <= paraCount; p++)
                {
                    dynamic para = textRange.Paragraphs(p, 1);
                    dynamic? pf = null;
                    dynamic? bullet = null;
                    try
                    {
                        pf = para.ParagraphFormat;
                        bullet = pf.Bullet;

                        int bulletType = Convert.ToInt32(bullet.Type);
                        int bulletChar = 0;
                        try { bulletChar = Convert.ToInt32(bullet.Character); } catch { }
                        int level = Convert.ToInt32(pf.Level);

                        lines.Add($"Paragraph {p}: BulletType={bulletType}, BulletCharacter={bulletChar}, IndentLevel={level}");
                    }
                    finally
                    {
                        if (bullet != null) ComUtilities.Release(ref bullet!);
                        if (pf != null) ComUtilities.Release(ref pf!);
                        ComUtilities.Release(ref para!);
                    }
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "read-bullets",
                    Message = string.Join("; ", lines),
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (allParagraphs != null) ComUtilities.Release(ref allParagraphs!);
                if (textRange != null) ComUtilities.Release(ref textRange!);
                if (textFrame != null) ComUtilities.Release(ref textFrame!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult InsertSymbol(IPptBatch batch, int slideIndex, string shapeName, string fontName, int charNumber)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);
        ArgumentException.ThrowIfNullOrWhiteSpace(fontName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? textFrame = null;
            dynamic? textRange = null;
            try
            {
                if (Convert.ToInt32(shape.HasTextFrame) == 0)
                    throw new InvalidOperationException($"Shape '{shapeName}' does not have a text frame.");

                textFrame = shape.TextFrame;
                textRange = textFrame.TextRange;
                textRange.InsertSymbol(fontName, charNumber, true);

                return new OperationResult
                {
                    Success = true,
                    Action = "insert-symbol",
                    Message = $"Inserted symbol (font='{fontName}', char={charNumber}) in shape '{shapeName}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (textRange != null) ComUtilities.Release(ref textRange!);
                if (textFrame != null) ComUtilities.Release(ref textFrame!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult InsertDateTime(IPptBatch batch, int slideIndex, string shapeName, int dateTimeFormat)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? textFrame = null;
            dynamic? textRange = null;
            try
            {
                if (Convert.ToInt32(shape.HasTextFrame) == 0)
                    throw new InvalidOperationException($"Shape '{shapeName}' does not have a text frame.");

                textFrame = shape.TextFrame;
                textRange = textFrame.TextRange;
                textRange.InsertDateTime(dateTimeFormat, true);

                return new OperationResult
                {
                    Success = true,
                    Action = "insert-datetime",
                    Message = $"Inserted date/time (format={dateTimeFormat}) in shape '{shapeName}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (textRange != null) ComUtilities.Release(ref textRange!);
                if (textFrame != null) ComUtilities.Release(ref textFrame!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult InsertSlideNumber(IPptBatch batch, int slideIndex, string shapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? textFrame = null;
            dynamic? textRange = null;
            try
            {
                if (Convert.ToInt32(shape.HasTextFrame) == 0)
                    throw new InvalidOperationException($"Shape '{shapeName}' does not have a text frame.");

                textFrame = shape.TextFrame;
                textRange = textFrame.TextRange;
                textRange.InsertSlideNumber();

                return new OperationResult
                {
                    Success = true,
                    Action = "insert-slide-number",
                    Message = $"Inserted slide number in shape '{shapeName}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (textRange != null) ComUtilities.Release(ref textRange!);
                if (textFrame != null) ComUtilities.Release(ref textFrame!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }
}
