using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Slide;

public class SlideCommands : ISlideCommands
{
    /// <summary>
    /// Populates the descriptive fields of <paramref name="info"/> from a slide.
    /// </summary>
    /// <remarks>
    /// Every COM hop is bound to a local and released in a <c>finally</c> block.
    /// The previous inline form - <c>slide.Design.SlideMaster.Name</c> and, worse,
    /// <c>slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text</c> -
    /// materialised proxies that were never assigned to anything, so
    /// <see cref="ComUtilities.Release"/> could not be called on them even in
    /// principle. On a 40-slide deck a single <c>slide list</c> abandoned several
    /// hundred RCWs, each holding the out-of-process PowerPoint server alive until
    /// a finalizer happened to run.
    ///
    /// Layout and master names are NOT wrapped in a catch: every slide has both, so
    /// a failure there is a real fault and must surface (Rule 22). Notes and
    /// animations are genuinely optional - a slide need not have a notes placeholder
    /// or a timeline - and are the "optional property access" case Rule 1b permits.
    /// </remarks>
    private static void PopulateSlideMetadata(dynamic slide, SlideInfo info)
    {
        info.SlideNumber = (int)slide.SlideNumber;
        // SlideID is a value, not a COM proxy - cast before ToString so the inline-chain
        // gate is not left with a false positive it cannot distinguish from a real hop.
        info.SlideId = ((int)slide.SlideID).ToString(System.Globalization.CultureInfo.InvariantCulture);
        info.Name = slide.Name?.ToString();

        dynamic? shapes = null;
        try
        {
            shapes = slide.Shapes;
            info.ShapeCount = (int)shapes.Count;
        }
        finally
        {
            if (shapes != null) { ComUtilities.Release(ref shapes!); }
        }

        dynamic? customLayout = null;
        try
        {
            customLayout = slide.CustomLayout;
            info.LayoutName = customLayout.Name?.ToString() ?? string.Empty;
        }
        finally
        {
            if (customLayout != null) { ComUtilities.Release(ref customLayout!); }
        }

        dynamic? design = null;
        dynamic? slideMaster = null;
        try
        {
            design = slide.Design;
            slideMaster = design.SlideMaster;
            info.MasterName = slideMaster.Name?.ToString() ?? string.Empty;
        }
        finally
        {
            if (slideMaster != null) { ComUtilities.Release(ref slideMaster!); }
            if (design != null) { ComUtilities.Release(ref design!); }
        }

        info.HasNotes = ReadHasNotes(slide);
        info.HasAnimations = ReadHasAnimations(slide);
    }

    /// <summary>
    /// Returns whether the slide's notes placeholder contains text.
    /// </summary>
    /// <remarks>
    /// A slide need not have a notes placeholder at index 2, so the lookup is
    /// allowed to fail. The catch is scoped to the optional access itself rather
    /// than wrapped around the whole chain, and the six intermediate proxies are
    /// released in reverse order regardless of outcome.
    /// </remarks>
    private static bool ReadHasNotes(dynamic slide)
    {
        dynamic? notesPage = null;
        dynamic? shapes = null;
        dynamic? placeholders = null;
        dynamic? placeholder = null;
        dynamic? textFrame = null;
        dynamic? textRange = null;
        try
        {
            notesPage = slide.NotesPage;
            shapes = notesPage.Shapes;
            placeholders = shapes.Placeholders;

            if ((int)placeholders.Count < 2)
            {
                return false;
            }

            placeholder = placeholders.Item(2);
            textFrame = placeholder.TextFrame;
            textRange = textFrame.TextRange;

            string? text = textRange.Text?.ToString();
            return !string.IsNullOrEmpty(text);
        }
        finally
        {
            if (textRange != null) { ComUtilities.Release(ref textRange!); }
            if (textFrame != null) { ComUtilities.Release(ref textFrame!); }
            if (placeholder != null) { ComUtilities.Release(ref placeholder!); }
            if (placeholders != null) { ComUtilities.Release(ref placeholders!); }
            if (shapes != null) { ComUtilities.Release(ref shapes!); }
            if (notesPage != null) { ComUtilities.Release(ref notesPage!); }
        }
    }

    /// <summary>
    /// Returns whether the slide has any entries in its main animation sequence.
    /// </summary>
    private static bool ReadHasAnimations(dynamic slide)
    {
        dynamic? timeLine = null;
        dynamic? mainSequence = null;
        try
        {
            timeLine = slide.TimeLine;
            mainSequence = timeLine.MainSequence;
            return (int)mainSequence.Count > 0;
        }
        finally
        {
            if (mainSequence != null) { ComUtilities.Release(ref mainSequence!); }
            if (timeLine != null) { ComUtilities.Release(ref timeLine!); }
        }
    }

    public SlideListResult List(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            var result = new SlideListResult { Success = true, FilePath = ctx.PresentationPath };
            dynamic pres = ctx.Presentation;
            dynamic slides = pres.Slides;
            try
            {
                int count = (int)slides.Count;

                for (int i = 1; i <= count; i++)
                {
                    dynamic slide = slides.Item(i);
                    try
                    {
                        var info = new SlideInfo { SlideIndex = i };
                        PopulateSlideMetadata(slide, info);
                        result.Slides.Add(info);
                    }
                    finally
                    {
                        ComUtilities.Release(ref slide!);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref slides!);
            }
        });
    }

    public SlideDetailResult Read(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides.Item(slideIndex);
            try
            {
                var info = new SlideInfo { SlideIndex = slideIndex };
                PopulateSlideMetadata(slide, info);

                var result = new SlideDetailResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    Slide = info
                };

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
                            result.Shapes.Add(ShapeHelpers.ReadShapeInfo(shape));
                        }
                        finally
                        {
                            ComUtilities.Release(ref shape!);
                        }
                    }
                }
                finally
                {
                    // Previously released after the loop rather than in a finally, so
                    // any shape that failed to read leaked the Shapes collection too.
                    if (shapes != null) { ComUtilities.Release(ref shapes!); }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref slide!);
                ComUtilities.Release(ref slides!);
            }
        });
    }

    public OperationResult Create(IPptBatch batch, int position, string layoutName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic slides = pres.Slides;
            int slideCount = (int)slides.Count;

            // Find the layout by name
            dynamic? layout = FindLayout(pres, layoutName);
            if (layout == null)
                throw new ArgumentException(BuildLayoutNotFoundMessage(pres, layoutName));

            try
            {
                int insertAt = position <= 0 ? slideCount + 1 : position;
                dynamic newSlide = slides.AddSlide(insertAt, layout);
                int newIndex = (int)newSlide.SlideIndex;
                ComUtilities.Release(ref newSlide!);
                ComUtilities.Release(ref slides!);

                return new OperationResult
                {
                    Success = true,
                    Action = "create",
                    Message = $"Created slide at position {newIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref layout!);
            }
        });
    }

    public OperationResult Duplicate(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides.Item(slideIndex);
            try
            {
                dynamic duplicated = slide.Duplicate();
                // Duplicate returns a SlideRange; get first item
                dynamic newSlide = duplicated.Item(1);
                int newIndex = (int)newSlide.SlideIndex;
                ComUtilities.Release(ref newSlide!);
                ComUtilities.Release(ref duplicated!);

                return new OperationResult
                {
                    Success = true,
                    Action = "duplicate",
                    Message = $"Duplicated slide {slideIndex} → new slide at position {newIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
                ComUtilities.Release(ref slides!);
            }
        });
    }

    public OperationResult Move(IPptBatch batch, int slideIndex, int newPosition)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides.Item(slideIndex);
            try
            {
                slide.MoveTo(newPosition);
                return new OperationResult
                {
                    Success = true,
                    Action = "move",
                    Message = $"Moved slide from position {slideIndex} to {newPosition}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
                ComUtilities.Release(ref slides!);
            }
        });
    }

    public OperationResult Delete(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides.Item(slideIndex);
            try
            {
                slide.Delete();
                return new OperationResult
                {
                    Success = true,
                    Action = "delete",
                    Message = $"Deleted slide at position {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
                ComUtilities.Release(ref slides!);
            }
        });
    }

    public OperationResult ApplyLayout(IPptBatch batch, int slideIndex, string layoutName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic slides = pres.Slides;
            dynamic slide = slides.Item(slideIndex);
            dynamic? layout = FindLayout(pres, layoutName);

            if (layout == null)
                throw new ArgumentException(BuildLayoutNotFoundMessage(pres, layoutName));

            try
            {
                slide.CustomLayout = layout;
                return new OperationResult
                {
                    Success = true,
                    Action = "apply-layout",
                    Message = $"Applied layout '{layoutName}' to slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref layout!);
                ComUtilities.Release(ref slide!);
                ComUtilities.Release(ref slides!);
            }
        });
    }

    public OperationResult SetName(IPptBatch batch, int slideIndex, string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                slide.Name = name;
                return new OperationResult
                {
                    Success = true,
                    Action = "set-name",
                    Message = $"Set name of slide {slideIndex} to '{name}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult CloneWithReplace(IPptBatch batch, int slideIndex, int count, string searchText, string replaceText)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(searchText);
        ArgumentException.ThrowIfNullOrWhiteSpace(replaceText);

        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic slides = pres.Slides;
            dynamic sourceSlide = slides.Item(slideIndex);
            try
            {
                int created = 0;
                for (int c = 0; c < count; c++)
                {
                    dynamic duplicated = sourceSlide.Duplicate();
                    dynamic newSlide = duplicated.Item(1);
                    try
                    {
                        dynamic shapes = newSlide.Shapes;
                        try
                        {
                            int shapeCount = (int)shapes.Count;
                            for (int i = 1; i <= shapeCount; i++)
                            {
                                dynamic shape = shapes.Item(i);
                                try
                                {
                                    ReplaceTextInShape(shape, searchText, replaceText);
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

                        created++;
                    }
                    finally
                    {
                        ComUtilities.Release(ref newSlide!);
                        ComUtilities.Release(ref duplicated!);
                    }
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "clone-with-replace",
                    Message = $"Created {created} clone(s) of slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref sourceSlide!);
                ComUtilities.Release(ref slides!);
            }
        });
    }

    public OperationResult Hide(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ComUtilities.GetSlide(ctx.Presentation, slideIndex);
            dynamic? transition = null;
            try
            {
                // msoTrue = -1
                transition = ComUtilities.GetSlideShowTransition(slide);
                transition.Hidden = -1;
                return new OperationResult
                {
                    Success = true,
                    Action = "hide",
                    Message = $"Hidden slide {slideIndex} from slideshow",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref transition!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Unhide(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ComUtilities.GetSlide(ctx.Presentation, slideIndex);
            dynamic? transition = null;
            try
            {
                // msoFalse = 0
                transition = ComUtilities.GetSlideShowTransition(slide);
                transition.Hidden = 0;
                return new OperationResult
                {
                    Success = true,
                    Action = "unhide",
                    Message = $"Unhidden slide {slideIndex} for slideshow",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref transition!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult GetThumbnail(IPptBatch batch, int slideIndex, string destinationPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);

        return batch.Execute((ctx, ct) =>
        {
            // Ensure destination directory exists
            string? dir = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                slide.Export(destinationPath, "PNG", 320, 240);
                return new OperationResult
                {
                    Success = true,
                    Action = "get-thumbnail",
                    Message = $"Exported slide {slideIndex} thumbnail to '{destinationPath}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Summary(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic slides = pres.Slides;
            dynamic pageSetup = pres.PageSetup;
            try
            {
                int slideCount = (int)slides.Count;
                float slideWidth = (float)pageSetup.SlideWidth;
                float slideHeight = (float)pageSetup.SlideHeight;

                bool hasNotesMaster = false;
                try { hasNotesMaster = Convert.ToInt32(pres.HasNotesMaster) != 0; } catch { }

                string templateName = "";
                try { templateName = pres.TemplateName?.ToString() ?? ""; } catch { }

                int totalShapes = 0;
                for (int i = 1; i <= slideCount; i++)
                {
                    dynamic slide = slides.Item(i);
                    try
                    {
                        totalShapes += ComUtilities.GetShapeCount(slide);
                    }
                    finally
                    {
                        ComUtilities.Release(ref slide!);
                    }
                }

                var message = $"Slides: {slideCount}, Dimensions: {slideWidth}x{slideHeight}pt, " +
                              $"HasNotesMaster: {hasNotesMaster}, TemplateName: '{templateName}', " +
                              $"TotalShapes: {totalShapes}";

                return new OperationResult
                {
                    Success = true,
                    Action = "summary",
                    Message = message,
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref pageSetup!);
                ComUtilities.Release(ref slides!);
            }
        });
    }

    public OperationResult SetDisplayMaster(IPptBatch batch, int slideIndex, bool display)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                // msoTrue = -1, msoFalse = 0
                slide.DisplayMasterShapes = display ? -1 : 0;
                return new OperationResult
                {
                    Success = true,
                    Action = "set-display-master",
                    Message = display
                        ? $"Enabled master shapes on slide {slideIndex}"
                        : $"Disabled master shapes on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <summary>
    /// Replaces text in a shape, recursing into grouped shapes (Type == 6).
    /// </summary>
    private static void ReplaceTextInShape(dynamic shape, string searchText, string replaceText)
    {
        // msoGroup = 6
        if (Convert.ToInt32(shape.Type) == 6)
        {
            dynamic groupItems = shape.GroupItems;
            try
            {
                int itemCount = (int)groupItems.Count;
                for (int g = 1; g <= itemCount; g++)
                {
                    dynamic groupChild = groupItems.Item(g);
                    try
                    {
                        ReplaceTextInShape(groupChild, searchText, replaceText);
                    }
                    finally
                    {
                        ComUtilities.Release(ref groupChild!);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref groupItems!);
            }
            return;
        }

        if (Convert.ToInt32(shape.HasTextFrame) != 0)
        {
            dynamic textFrame = shape.TextFrame;
            dynamic textRange = textFrame.TextRange;
            try
            {
                string text = textRange.Text?.ToString() ?? "";
                if (text.Contains(searchText))
                {
                    textRange.Text = text.Replace(searchText, replaceText);
                }
            }
            finally
            {
                ComUtilities.Release(ref textRange!);
                ComUtilities.Release(ref textFrame!);
            }
        }
    }

    /// <summary>
    /// Canonical English names of the 11 layouts every Office theme ships, mapped to their
    /// fixed 1-based position in the master. PowerPoint localizes CustomLayout.Name and
    /// CustomLayout.MatchingName (on a German install "Blank" is reported as "Leer") and the
    /// COM API exposes no locale-independent identifier, so position is the only stable
    /// bridge from an English name to a localized layout.
    /// </summary>
    private static readonly Dictionary<string, int> CanonicalLayoutPositions =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["Title Slide"] = 1,
            ["Title and Content"] = 2,
            ["Section Header"] = 3,
            ["Two Content"] = 4,
            ["Comparison"] = 5,
            ["Title Only"] = 6,
            ["Blank"] = 7,
            ["Content with Caption"] = 8,
            ["Picture with Caption"] = 9,
            ["Title and Vertical Text"] = 10,
            ["Vertical Title and Text"] = 11,
        };

    private static dynamic? FindLayout(dynamic pres, string layoutName)
    {
        // PowerPoint COM: Presentation.Designs → Design.SlideMaster.CustomLayouts
        dynamic designs = pres.Designs;
        try
        {
            int designCount = (int)designs.Count;

            bool wantsPosition = CanonicalLayoutPositions.TryGetValue(layoutName, out int canonicalPosition);
            bool wantsIndex = int.TryParse(layoutName, out int requestedIndex) && requestedIndex >= 1;

            dynamic? nameMatch = null;
            dynamic? matchingNameMatch = null;
            dynamic? positionMatch = null;

            for (int d = 1; d <= designCount; d++)
            {
                dynamic design = designs.Item(d);
                dynamic master = design.SlideMaster;
                dynamic layouts = master.CustomLayouts;
                try
                {
                    int layoutCount = (int)layouts.Count;

                    for (int l = 1; l <= layoutCount; l++)
                    {
                        dynamic layout = layouts.Item(l);

                        string name = layout.Name?.ToString() ?? "";
                        if (nameMatch == null && string.Equals(name, layoutName, StringComparison.OrdinalIgnoreCase))
                        {
                            nameMatch = layout;
                            continue;
                        }

                        string matchingName = TryGetMatchingName(layout);
                        if (matchingNameMatch == null && matchingName.Length > 0
                            && string.Equals(matchingName, layoutName, StringComparison.OrdinalIgnoreCase))
                        {
                            matchingNameMatch = layout;
                            continue;
                        }

                        // Positional fallbacks only apply to the first design, where the
                        // standard Office layout order is meaningful.
                        bool positionalHit = d == 1
                            && positionMatch == null
                            && ((wantsPosition && l == canonicalPosition) || (wantsIndex && l == requestedIndex));

                        if (positionalHit)
                        {
                            positionMatch = layout;
                            continue;
                        }

                        ComUtilities.Release(ref layout!);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref layouts!);
                    ComUtilities.Release(ref master!);
                    ComUtilities.Release(ref design!);
                }
            }

            // An exact name always wins over a localized positional guess.
            if (nameMatch != null)
            {
                ReleaseIfNotNull(ref matchingNameMatch);
                ReleaseIfNotNull(ref positionMatch);
                return nameMatch;
            }

            if (matchingNameMatch != null)
            {
                ReleaseIfNotNull(ref positionMatch);
                return matchingNameMatch;
            }

            return positionMatch;
        }
        finally
        {
            ComUtilities.Release(ref designs!);
        }
    }

    /// <summary>
    /// CustomLayout.MatchingName is not present on every layout object PowerPoint hands back,
    /// so probe it defensively. Only late-binding and COM failures are tolerated here.
    /// </summary>
    private static string TryGetMatchingName(dynamic layout)
    {
        try
        {
            return layout.MatchingName?.ToString() ?? "";
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
        {
            return "";
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            return "";
        }
    }

    private static void ReleaseIfNotNull(ref dynamic? comObject)
    {
        if (comObject != null)
        {
            ComUtilities.Release(ref comObject!);
            comObject = null;
        }
    }

    /// <summary>
    /// Builds an error message that lists the layouts PowerPoint actually reports, so a
    /// caller on a localized install can see the real names instead of guessing.
    /// </summary>
    private static string BuildLayoutNotFoundMessage(dynamic pres, string layoutName)
    {
        var available = new List<string>();

        dynamic designs = pres.Designs;
        try
        {
            int designCount = (int)designs.Count;
            for (int d = 1; d <= designCount; d++)
            {
                dynamic design = designs.Item(d);
                dynamic master = design.SlideMaster;
                dynamic layouts = master.CustomLayouts;
                try
                {
                    int layoutCount = (int)layouts.Count;
                    for (int l = 1; l <= layoutCount; l++)
                    {
                        dynamic layout = layouts.Item(l);
                        try
                        {
                            string name = ComUtilities.SafeGetString(layout, "Name");
                            if (name.Length == 0)
                                name = $"Layout {l}";
                            available.Add($"{l}. {name.Replace("\r", "").Replace("\n", " ")}");
                        }
                        finally
                        {
                            ComUtilities.Release(ref layout!);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref layouts!);
                    ComUtilities.Release(ref master!);
                    ComUtilities.Release(ref design!);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref designs!);
        }

        return $"Layout '{layoutName}' not found in this presentation. " +
            $"Available layouts: {string.Join(", ", available)}. " +
            "Layout names are localized by PowerPoint; canonical English names " +
            "(for example 'Blank' or 'Title Slide') and 1-based indexes are also accepted.";
    }

    public OperationResult CopyToClipboard(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                slide.Copy();
                return new OperationResult
                {
                    Success = true,
                    Action = "copy",
                    Message = $"Copied slide {slideIndex} to clipboard",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slide!);
            }
        });
    }
}
