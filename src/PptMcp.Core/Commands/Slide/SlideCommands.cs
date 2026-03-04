using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Slide;

public class SlideCommands : ISlideCommands
{
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
                        var info = new SlideInfo
                        {
                            SlideIndex = i,
                            SlideNumber = (int)slide.SlideNumber,
                            SlideId = slide.SlideID.ToString(),
                            ShapeCount = (int)slide.Shapes.Count,
                        };

                        try { info.LayoutName = slide.CustomLayout.Name?.ToString() ?? ""; } catch { info.LayoutName = ""; }
                        try { info.MasterName = slide.Design.SlideMaster.Name?.ToString() ?? ""; } catch { info.MasterName = ""; }
                        try { info.HasNotes = slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text?.ToString()?.Length > 0; } catch { info.HasNotes = false; }
                        try { info.HasAnimations = (int)slide.TimeLine.MainSequence.Count > 0; } catch { info.HasAnimations = false; }
                        try { info.Name = slide.Name?.ToString(); } catch { }

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
                var result = new SlideDetailResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    Slide = new SlideInfo
                    {
                        SlideIndex = slideIndex,
                        SlideNumber = (int)slide.SlideNumber,
                        SlideId = slide.SlideID.ToString(),
                        ShapeCount = (int)slide.Shapes.Count,
                    }
                };

                try { result.Slide.LayoutName = slide.CustomLayout.Name?.ToString() ?? ""; } catch { result.Slide.LayoutName = ""; }
                try { result.Slide.MasterName = slide.Design.SlideMaster.Name?.ToString() ?? ""; } catch { result.Slide.MasterName = ""; }
                try { result.Slide.Name = slide.Name?.ToString(); } catch { }

                dynamic shapes = slide.Shapes;
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
                ComUtilities.Release(ref shapes!);

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
                throw new ArgumentException($"Layout '{layoutName}' not found in this presentation.");

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
                throw new ArgumentException($"Layout '{layoutName}' not found in this presentation.");

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
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                // msoTrue = -1
                slide.SlideShowTransition.Hidden = -1;
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
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Unhide(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                // msoFalse = 0
                slide.SlideShowTransition.Hidden = 0;
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
                        totalShapes += (int)slide.Shapes.Count;
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

    private static dynamic? FindLayout(dynamic pres, string layoutName)
    {
        // PowerPoint COM: Presentation.Designs → Design.SlideMaster.CustomLayouts
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
                        string name = layout.Name?.ToString() ?? "";
                        if (string.Equals(name, layoutName, StringComparison.OrdinalIgnoreCase))
                        {
                            return layout;
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

            return null;
        }
        finally
        {
            ComUtilities.Release(ref designs!);
        }
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
