using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Master;

public class MasterCommands : IMasterCommands
{
    public MasterListResult List(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            var result = new MasterListResult { Success = true, FilePath = ctx.PresentationPath };
            dynamic pres = ctx.Presentation;
            dynamic masters = pres.SlideMasters;
            try
            {
                int masterCount = (int)masters.Count;

                for (int m = 1; m <= masterCount; m++)
                {
                    dynamic master = masters.Item(m);
                    try
                    {
                        var masterInfo = new MasterInfo
                        {
                            Name = master.Name?.ToString() ?? $"Master {m}"
                        };

                        dynamic layouts = master.CustomLayouts;
                        try
                        {
                            int layoutCount = (int)layouts.Count;
                            for (int l = 1; l <= layoutCount; l++)
                            {
                                dynamic layout = layouts.Item(l);
                                try
                                {
                                    masterInfo.Layouts.Add(new LayoutInfo
                                    {
                                        Name = layout.Name?.ToString() ?? $"Layout {l}",
                                        Index = l
                                    });
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
                        }

                        result.Masters.Add(masterInfo);
                    }
                    finally
                    {
                        ComUtilities.Release(ref master!);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref masters!);
            }
        });
    }

    public OperationResult ListShapes(IPptBatch batch, int masterIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic masters = ((dynamic)ctx.Presentation).SlideMasters;
            dynamic master = masters.Item(masterIndex);
            dynamic shapes = master.Shapes;
            try
            {
                int count = (int)shapes.Count;
                var lines = new List<string>(count);

                for (int i = 1; i <= count; i++)
                {
                    dynamic shape = shapes.Item(i);
                    try
                    {
                        string name = shape.Name?.ToString() ?? $"Shape {i}";
                        int shapeType = Convert.ToInt32(shape.Type);
                        string typeName = ShapeHelpers.GetShapeTypeName(shapeType);
                        lines.Add($"{name} ({typeName})");
                    }
                    finally
                    {
                        ComUtilities.Release(ref shape!);
                    }
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "list-shapes",
                    Message = count > 0
                        ? $"Master {masterIndex} has {count} shape(s):\n" + string.Join("\n", lines)
                        : $"Master {masterIndex} has no shapes.",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref shapes!);
                ComUtilities.Release(ref master!);
                ComUtilities.Release(ref masters!);
            }
        });
    }

    public OperationResult EditShapeText(IPptBatch batch, int masterIndex, string shapeName, string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic masters = ((dynamic)ctx.Presentation).SlideMasters;
            dynamic master = masters.Item(masterIndex);
            dynamic shape = master.Shapes.Item(shapeName);
            try
            {
                if (Convert.ToInt32(shape.HasTextFrame) == 0)
                    throw new InvalidOperationException($"Shape '{shapeName}' on master {masterIndex} does not have a text frame.");

                shape.TextFrame.TextRange.Text = text;

                return new OperationResult
                {
                    Success = true,
                    Action = "edit-shape-text",
                    Message = $"Set text on shape '{shapeName}' (master {masterIndex})",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref master!);
                ComUtilities.Release(ref masters!);
            }
        });
    }

    public OperationResult ListLayouts(IPptBatch batch, int masterIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic masters = ((dynamic)ctx.Presentation).SlideMasters;
            dynamic master = masters.Item(masterIndex);
            dynamic layouts = master.CustomLayouts;
            try
            {
                int count = (int)layouts.Count;
                var layoutInfos = new List<LayoutInfo>(count);

                for (int i = 1; i <= count; i++)
                {
                    dynamic layout = layouts.Item(i);
                    try
                    {
                        string name = layout.Name?.ToString() ?? $"Layout {i}";
                        string? matchingName = null;
                        try { matchingName = layout.MatchingName?.ToString(); } catch { }

                        layoutInfos.Add(new LayoutInfo
                        {
                            Name = name,
                            Index = i,
                            MatchingName = string.IsNullOrEmpty(matchingName) ? null : matchingName
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref layout!);
                    }
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "list-layouts",
                    Message = $"Master {masterIndex} has {count} custom layout(s):\n" +
                        string.Join("\n", layoutInfos.Select(l => l.MatchingName != null
                            ? $"  {l.Index}. {l.Name} (matching: {l.MatchingName})"
                            : $"  {l.Index}. {l.Name}")),
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref layouts!);
                ComUtilities.Release(ref master!);
                ComUtilities.Release(ref masters!);
            }
        });
    }

    public OperationResult DeleteUnused(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic designs = pres.Designs;
            dynamic slides = pres.Slides;
            try
            {
                int slideCount = (int)slides.Count;
                int designCount = (int)designs.Count;

                // Build a set of design names that are in use
                var usedDesignNames = new HashSet<string>();
                for (int s = 1; s <= slideCount; s++)
                {
                    dynamic slide = slides.Item(s);
                    try
                    {
                        string designName = slide.Design.Name?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(designName))
                            usedDesignNames.Add(designName);
                    }
                    finally
                    {
                        ComUtilities.Release(ref slide!);
                    }
                }

                // Delete unused designs in reverse order to avoid index shifts
                int deletedCount = 0;
                for (int d = designCount; d >= 1; d--)
                {
                    // Never delete the last remaining design
                    int currentCount = (int)designs.Count;
                    if (currentCount <= 1)
                        break;

                    dynamic design = designs.Item(d);
                    try
                    {
                        string name = design.Name?.ToString() ?? "";
                        if (!usedDesignNames.Contains(name))
                        {
                            design.Delete();
                            deletedCount++;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref design!);
                    }
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "delete-unused",
                    Message = deletedCount > 0
                        ? $"Deleted {deletedCount} unused master(s)"
                        : "No unused masters found",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref slides!);
                ComUtilities.Release(ref designs!);
            }
        });
    }
}
