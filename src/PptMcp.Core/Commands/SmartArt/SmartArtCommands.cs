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
}
