using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.SmartArt;

public class SmartArtCommands : ISmartArtCommands
{
    public SmartArtInfoResult GetInfo(IPptBatch batch, int slideIndex, string shapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                if (Convert.ToInt32(shape.HasSmartArt) == 0)
                    throw new InvalidOperationException($"Shape '{shapeName}' is not a SmartArt diagram.");

                dynamic smartArt = shape.SmartArt;
                try
                {
                    var result = new SmartArtInfoResult
                    {
                        Success = true,
                        FilePath = ctx.PresentationPath,
                        SlideIndex = slideIndex,
                        ShapeName = shapeName,
                    };

                    try { result.LayoutName = smartArt.Layout.Name?.ToString() ?? ""; } catch { }

                    dynamic nodes = smartArt.AllNodes;
                    try
                    {
                        int count = (int)nodes.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            dynamic node = nodes.Item(i);
                            try
                            {
                                string text = "";
                                try { text = node.TextFrame2.TextRange.Text?.ToString() ?? ""; } catch { }
                                result.Nodes.Add(new SmartArtNodeInfo
                                {
                                    Index = i,
                                    Text = text,
                                    Level = Convert.ToInt32(node.Level)
                                });
                            }
                            finally
                            {
                                ComUtilities.Release(ref node!);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref nodes!);
                    }

                    return result;
                }
                finally
                {
                    ComUtilities.Release(ref smartArt!);
                }
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult AddNode(IPptBatch batch, int slideIndex, string shapeName, string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);
        ArgumentException.ThrowIfNullOrWhiteSpace(text);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                if (Convert.ToInt32(shape.HasSmartArt) == 0)
                    throw new InvalidOperationException($"Shape '{shapeName}' is not a SmartArt diagram.");

                dynamic? smartArt = null;
                dynamic? nodes = null;
                dynamic? newNode = null;
                try
                {
                    smartArt = shape.SmartArt;
                    nodes = smartArt.AllNodes;
                    // AddNode() adds after the last node
                    newNode = nodes.Add();
                    newNode.TextFrame2.TextRange.Text = text;

                    return new OperationResult
                    {
                        Success = true,
                        Action = "add-node",
                        Message = $"Added node '{text}' to SmartArt '{shapeName}' on slide {slideIndex}",
                        FilePath = ctx.PresentationPath
                    };
                }
                finally
                {
                    if (newNode != null) ComUtilities.Release(ref newNode!);
                    if (nodes != null) ComUtilities.Release(ref nodes!);
                    if (smartArt != null) ComUtilities.Release(ref smartArt!);
                }
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetLayout(IPptBatch batch, int slideIndex, string shapeName, int layoutIndex)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                if (Convert.ToInt32(shape.HasSmartArt) == 0)
                    throw new InvalidOperationException($"Shape '{shapeName}' is not a SmartArt diagram.");

                dynamic? smartArt = null;
                dynamic? app = null;
                dynamic? layouts = null;
                dynamic? layout = null;
                try
                {
                    smartArt = shape.SmartArt;
                    app = ctx.Presentation.Application;
                    layouts = app.SmartArtLayouts;
                    layout = layouts.Item(layoutIndex);
                    smartArt.Layout = layout;

                    return new OperationResult
                    {
                        Success = true,
                        Action = "set-layout",
                        Message = $"Set SmartArt layout to index {layoutIndex} on shape '{shapeName}' on slide {slideIndex}",
                        FilePath = ctx.PresentationPath
                    };
                }
                finally
                {
                    if (layout != null) ComUtilities.Release(ref layout!);
                    if (layouts != null) ComUtilities.Release(ref layouts!);
                    if (app != null) ComUtilities.Release(ref app!);
                    if (smartArt != null) ComUtilities.Release(ref smartArt!);
                }
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetStyle(IPptBatch batch, int slideIndex, string shapeName, int styleIndex)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                if (Convert.ToInt32(shape.HasSmartArt) == 0)
                    throw new InvalidOperationException($"Shape '{shapeName}' is not a SmartArt diagram.");

                dynamic? smartArt = null;
                dynamic? app = null;
                dynamic? styles = null;
                dynamic? style = null;
                try
                {
                    smartArt = shape.SmartArt;
                    app = ctx.Presentation.Application;
                    styles = app.SmartArtQuickStyles;
                    style = styles.Item(styleIndex);
                    smartArt.QuickStyle = style;

                    return new OperationResult
                    {
                        Success = true,
                        Action = "set-style",
                        Message = $"Set SmartArt quick style to index {styleIndex} on shape '{shapeName}' on slide {slideIndex}",
                        FilePath = ctx.PresentationPath
                    };
                }
                finally
                {
                    if (style != null) ComUtilities.Release(ref style!);
                    if (styles != null) ComUtilities.Release(ref styles!);
                    if (app != null) ComUtilities.Release(ref app!);
                    if (smartArt != null) ComUtilities.Release(ref smartArt!);
                }
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult DeleteNode(IPptBatch batch, int slideIndex, string shapeName, int nodeIndex)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                if (Convert.ToInt32(shape.HasSmartArt) == 0)
                    throw new InvalidOperationException($"Shape '{shapeName}' is not a SmartArt diagram.");

                dynamic? smartArt = null;
                dynamic? nodes = null;
                dynamic? node = null;
                try
                {
                    smartArt = shape.SmartArt;
                    nodes = smartArt.AllNodes;
                    node = nodes.Item(nodeIndex);
                    node.Delete();

                    return new OperationResult
                    {
                        Success = true,
                        Action = "delete-node",
                        Message = $"Deleted node {nodeIndex} from SmartArt '{shapeName}' on slide {slideIndex}",
                        FilePath = ctx.PresentationPath
                    };
                }
                finally
                {
                    if (node != null) ComUtilities.Release(ref node!);
                    if (nodes != null) ComUtilities.Release(ref nodes!);
                    if (smartArt != null) ComUtilities.Release(ref smartArt!);
                }
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult ChangeNodeLevel(IPptBatch batch, int slideIndex, string shapeName, int nodeIndex, bool promote)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                if (Convert.ToInt32(shape.HasSmartArt) == 0)
                    throw new InvalidOperationException($"Shape '{shapeName}' is not a SmartArt diagram.");

                dynamic? smartArt = null;
                dynamic? nodes = null;
                dynamic? node = null;
                try
                {
                    smartArt = shape.SmartArt;
                    nodes = smartArt.AllNodes;
                    node = nodes.Item(nodeIndex);

                    if (promote)
                        node.Promote();
                    else
                        node.Demote();

                    string action = promote ? "promoted" : "demoted";

                    return new OperationResult
                    {
                        Success = true,
                        Action = "change-level",
                        Message = $"Node {nodeIndex} {action} in SmartArt '{shapeName}' on slide {slideIndex}",
                        FilePath = ctx.PresentationPath
                    };
                }
                finally
                {
                    if (node != null) ComUtilities.Release(ref node!);
                    if (nodes != null) ComUtilities.Release(ref nodes!);
                    if (smartArt != null) ComUtilities.Release(ref smartArt!);
                }
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }
}
