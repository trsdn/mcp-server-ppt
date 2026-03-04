using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Accessibility;

public class AccessibilityCommands : IAccessibilityCommands
{
    public AccessibilityAuditResult Audit(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            var result = new AccessibilityAuditResult
            {
                Success = true,
                FilePath = ctx.PresentationPath
            };

            dynamic pres = ctx.Presentation;
            dynamic slides = pres.Slides;
            try
            {
                int slideCount = (int)slides.Count;
                result.TotalSlides = slideCount;

                for (int si = 1; si <= slideCount; si++)
                {
                    dynamic slide = slides.Item(si);
                    try
                    {
                        AuditSlide(slide, si, result.Issues);
                    }
                    finally
                    {
                        ComUtilities.Release(ref slide!);
                    }
                }

                result.IssueCount = result.Issues.Count;

                if (result.Issues.Count == 0)
                {
                    result.Message = "No accessibility issues found.";
                }
                else
                {
                    result.Message = $"Found {result.Issues.Count} accessibility issue(s) across {slideCount} slide(s).";
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref slides!);
            }
        });
    }

    public ReadingOrderResult GetReadingOrder(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                var result = new ReadingOrderResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    SlideIndex = slideIndex
                };

                dynamic shapes = slide.Shapes;
                try
                {
                    int count = (int)shapes.Count;
                    var entries = new List<ReadingOrderEntry>(count);

                    for (int i = 1; i <= count; i++)
                    {
                        dynamic shape = shapes.Item(i);
                        try
                        {
                            int shapeType = Convert.ToInt32(shape.Type);
                            entries.Add(new ReadingOrderEntry
                            {
                                ShapeName = shape.Name?.ToString() ?? "",
                                ShapeType = ShapeHelpers.GetShapeTypeName(shapeType),
                                ZOrderPosition = (int)shape.ZOrderPosition
                            });
                        }
                        finally
                        {
                            ComUtilities.Release(ref shape!);
                        }
                    }

                    // Sort by ZOrderPosition (reading order)
                    entries.Sort((a, b) => a.ZOrderPosition.CompareTo(b.ZOrderPosition));

                    for (int i = 0; i < entries.Count; i++)
                    {
                        entries[i].Position = i + 1;
                    }

                    result.Shapes = entries;
                    return result;
                }
                finally
                {
                    ComUtilities.Release(ref shapes!);
                }
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetReadingOrder(IPptBatch batch, int slideIndex, string shapeNames)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                var names = shapeNames
                    .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .ToList();

                if (names.Count == 0)
                    throw new ArgumentException("shapeNames must contain at least one shape name.");

                dynamic shapes = slide.Shapes;
                try
                {
                    // Send each shape to back in reverse order so the first name ends up with lowest ZOrder
                    // msoSendToBack = 1
                    for (int i = names.Count - 1; i >= 0; i--)
                    {
                        dynamic shape = shapes.Item(names[i]);
                        try
                        {
                            shape.ZOrder(1); // msoSendToBack
                        }
                        finally
                        {
                            ComUtilities.Release(ref shape!);
                        }
                    }

                    // Now bring each forward in order so they stack correctly
                    // msoSendToBack already placed them; now bring them to front in order
                    // to get the desired reading order: first name = lowest ZOrder
                    for (int i = 0; i < names.Count; i++)
                    {
                        dynamic shape = shapes.Item(names[i]);
                        try
                        {
                            shape.ZOrder(0); // msoBringToFront
                        }
                        finally
                        {
                            ComUtilities.Release(ref shape!);
                        }
                    }

                    return new OperationResult
                    {
                        Success = true,
                        Action = "set-reading-order",
                        Message = $"Set reading order for {names.Count} shape(s) on slide {slideIndex}.",
                        FilePath = ctx.PresentationPath
                    };
                }
                finally
                {
                    ComUtilities.Release(ref shapes!);
                }
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    private static void AuditSlide(dynamic slide, int slideIndex, List<AccessibilityIssue> issues)
    {
        bool hasTitle = false;

        // Check placeholders for title
        dynamic? placeholders = null;
        try
        {
            placeholders = slide.Shapes.Placeholders;
            int phCount = (int)placeholders.Count;
            for (int pi = 1; pi <= phCount; pi++)
            {
                dynamic ph = placeholders.Item(pi);
                try
                {
                    int phType = Convert.ToInt32(ph.PlaceholderFormat.Type);
                    // ppPlaceholderTitle = 1, ppPlaceholderCenterTitle = 3
                    if (phType == 1 || phType == 3)
                    {
                        hasTitle = true;

                        // Check if title placeholder has text
                        bool hasText = false;
                        try
                        {
                            if (Convert.ToInt32(ph.HasTextFrame) != 0)
                            {
                                string? text = ph.TextFrame.TextRange.Text?.ToString();
                                hasText = !string.IsNullOrWhiteSpace(text);
                            }
                        }
                        catch { }

                        if (!hasText)
                        {
                            issues.Add(new AccessibilityIssue
                            {
                                SlideIndex = slideIndex,
                                IssueType = "EmptyTitlePlaceholder",
                                ShapeName = ph.Name?.ToString(),
                                Description = "Title placeholder exists but has no text."
                            });
                        }
                    }
                    else
                    {
                        // Check other placeholders for empty text
                        try
                        {
                            if (Convert.ToInt32(ph.HasTextFrame) != 0)
                            {
                                string? text = ph.TextFrame.TextRange.Text?.ToString();
                                if (string.IsNullOrWhiteSpace(text))
                                {
                                    issues.Add(new AccessibilityIssue
                                    {
                                        SlideIndex = slideIndex,
                                        IssueType = "EmptyPlaceholder",
                                        ShapeName = ph.Name?.ToString(),
                                        Description = "Placeholder has no text content."
                                    });
                                }
                            }
                        }
                        catch { }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref ph!);
                }
            }
        }
        catch { }
        finally
        {
            if (placeholders != null) ComUtilities.Release(ref placeholders!);
        }

        if (!hasTitle)
        {
            issues.Add(new AccessibilityIssue
            {
                SlideIndex = slideIndex,
                IssueType = "MissingTitle",
                Description = "Slide has no title placeholder."
            });
        }

        // Check all shapes for missing alt text
        dynamic? shapes = null;
        try
        {
            shapes = slide.Shapes;
            int shapeCount = (int)shapes.Count;
            for (int i = 1; i <= shapeCount; i++)
            {
                dynamic shape = shapes.Item(i);
                try
                {
                    int shapeType = Convert.ToInt32(shape.Type);
                    // Skip placeholders (already checked), comments, and lines
                    if (shapeType == 14 || shapeType == 4 || shapeType == 9)
                        continue;

                    string? altText = null;
                    try { altText = shape.AlternativeText?.ToString(); } catch { }

                    if (string.IsNullOrWhiteSpace(altText))
                    {
                        string shapeName = shape.Name?.ToString() ?? "";
                        issues.Add(new AccessibilityIssue
                        {
                            SlideIndex = slideIndex,
                            IssueType = "MissingAltText",
                            ShapeName = shapeName,
                            Description = $"Shape '{shapeName}' ({ShapeHelpers.GetShapeTypeName(shapeType)}) has no alternative text."
                        });
                    }
                }
                finally
                {
                    ComUtilities.Release(ref shape!);
                }
            }
        }
        catch { }
        finally
        {
            if (shapes != null) ComUtilities.Release(ref shapes!);
        }
    }
}
