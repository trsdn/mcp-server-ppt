using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Shape;

public class ShapeCommands : IShapeCommands
{
    public ShapeListResult List(IPptBatch batch, int slideIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shapes = slide.Shapes;
            try
            {
                int count = (int)shapes.Count;

                var result = new ShapeListResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    SlideIndex = slideIndex
                };

                for (int i = 1; i <= count; i++)
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

                return result;
            }
            finally
            {
                ComUtilities.Release(ref shapes!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public ShapeDetailResult Read(IPptBatch batch, int slideIndex, string shapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                return new ShapeDetailResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    Shape = ShapeHelpers.ReadShapeInfo(shape)
                };
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult AddTextbox(IPptBatch batch, int slideIndex, float left, float top, float width, float height, string text)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            // msoTextOrientationHorizontal = 1
            dynamic shape = slide.Shapes.AddTextbox(1, left, top, width, height);
            try
            {
                shape.TextFrame.TextRange.Text = text;
                string name = shape.Name?.ToString() ?? "";
                return new OperationResult
                {
                    Success = true,
                    Action = "add-textbox",
                    Message = $"Added textbox '{name}' on slide {slideIndex}",
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

    public OperationResult AddShape(IPptBatch batch, int slideIndex, int autoShapeType, float left, float top, float width, float height)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.AddShape(autoShapeType, left, top, width, height);
            try
            {
                string name = shape.Name?.ToString() ?? "";
                return new OperationResult
                {
                    Success = true,
                    Action = "add-shape",
                    Message = $"Added shape '{name}' (type {autoShapeType}) on slide {slideIndex}",
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

    public OperationResult MoveResize(IPptBatch batch, int slideIndex, string shapeName, float? left, float? top, float? width, float? height)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                if (left.HasValue) shape.Left = left.Value;
                if (top.HasValue) shape.Top = top.Value;
                if (width.HasValue) shape.Width = width.Value;
                if (height.HasValue) shape.Height = height.Value;

                return new OperationResult
                {
                    Success = true,
                    Action = "move-resize",
                    Message = $"Updated position/size of shape '{shapeName}' on slide {slideIndex}",
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

    public OperationResult Delete(IPptBatch batch, int slideIndex, string shapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                shape.Delete();
                return new OperationResult
                {
                    Success = true,
                    Action = "delete",
                    Message = $"Deleted shape '{shapeName}' from slide {slideIndex}",
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

    public OperationResult ZOrder(IPptBatch batch, int slideIndex, string shapeName, int zOrderCmd)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                shape.ZOrder(zOrderCmd);
                return new OperationResult
                {
                    Success = true,
                    Action = "z-order",
                    Message = $"Changed z-order of shape '{shapeName}' on slide {slideIndex}",
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

    public OperationResult SetFill(IPptBatch batch, int slideIndex, string shapeName, string colorHex)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);
        ArgumentException.ThrowIfNullOrWhiteSpace(colorHex);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                if (colorHex.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    // msoFillBackground = 5 (transparent/no fill)
                    shape.Fill.Visible = 0; // msoFalse
                }
                else
                {
                    shape.Fill.Visible = -1; // msoTrue
                    shape.Fill.Solid();
                    shape.Fill.ForeColor.RGB = HexToOleColor(colorHex);
                }
                return new OperationResult
                {
                    Success = true,
                    Action = "set-fill",
                    Message = $"Set fill of shape '{shapeName}' to '{colorHex}' on slide {slideIndex}",
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

    public OperationResult SetLine(IPptBatch batch, int slideIndex, string shapeName, string colorHex, float lineWidth)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);
        ArgumentException.ThrowIfNullOrWhiteSpace(colorHex);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                if (colorHex.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    shape.Line.Visible = 0; // msoFalse
                }
                else
                {
                    shape.Line.Visible = -1; // msoTrue
                    shape.Line.ForeColor.RGB = HexToOleColor(colorHex);
                    if (lineWidth > 0)
                        shape.Line.Weight = lineWidth;
                }
                return new OperationResult
                {
                    Success = true,
                    Action = "set-line",
                    Message = $"Set line of shape '{shapeName}' to '{colorHex}' (weight {lineWidth}pt) on slide {slideIndex}",
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

    public OperationResult SetRotation(IPptBatch batch, int slideIndex, string shapeName, float degrees)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                shape.Rotation = degrees;
                return new OperationResult
                {
                    Success = true,
                    Action = "set-rotation",
                    Message = $"Rotated shape '{shapeName}' to {degrees}° on slide {slideIndex}",
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

    public OperationResult Group(IPptBatch batch, int slideIndex, string shapeNames)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeNames);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic? group = null;
            try
            {
                // Select shapes by name, then group selection
                string[] names = shapeNames.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                dynamic? first = null;
                try
                {
                    first = slide.Shapes.Item(names[0]);
                    first.Select(true); // Replace=true to start new selection
                }
                finally
                {
                    if (first != null) ComUtilities.Release(ref first!);
                }
                for (int i = 1; i < names.Length; i++)
                {
                    dynamic? s = null;
                    try
                    {
                        s = slide.Shapes.Item(names[i]);
                        s.Select(false); // Replace=false to add to selection
                    }
                    finally
                    {
                        if (s != null) ComUtilities.Release(ref s!);
                    }
                }
                group = slide.Shapes.SelectAll();
                // Actually we need to use the selection to group
                dynamic app = ((dynamic)ctx.Presentation).Application;
                dynamic? sel = null;
                dynamic? grouped = null;
                try
                {
                    sel = app.ActiveWindow.Selection;
                    grouped = sel.ShapeRange.Group();
                    string groupName = grouped.Name?.ToString() ?? "";
                    return new OperationResult
                    {
                        Success = true,
                        Action = "group",
                        Message = $"Grouped {names.Length} shapes into '{groupName}' on slide {slideIndex}",
                        FilePath = ctx.PresentationPath
                    };
                }
                finally
                {
                    if (grouped != null) ComUtilities.Release(ref grouped!);
                    if (sel != null) ComUtilities.Release(ref sel!);
                }
            }
            finally
            {
                if (group != null) ComUtilities.Release(ref group!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Ungroup(IPptBatch batch, int slideIndex, string shapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                shape.Ungroup();
                return new OperationResult
                {
                    Success = true,
                    Action = "ungroup",
                    Message = $"Ungrouped shape '{shapeName}' on slide {slideIndex}",
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

    public OperationResult SetAltText(IPptBatch batch, int slideIndex, string shapeName, string altText)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                shape.AlternativeText = altText;
                return new OperationResult
                {
                    Success = true,
                    Action = "set-alt-text",
                    Message = $"Set alt text of shape '{shapeName}' to '{altText}' on slide {slideIndex}",
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

    /// <summary>Convert hex color string (#RRGGBB) to OLE color (R | G&lt;&lt;8 | B&lt;&lt;16).</summary>
    private static int HexToOleColor(string hex)
    {
        hex = hex.TrimStart('#');
        if (hex.Length == 3)
            hex = string.Concat(hex[0], hex[0], hex[1], hex[1], hex[2], hex[2]);
        int r = Convert.ToInt32(hex[..2], 16);
        int g = Convert.ToInt32(hex[2..4], 16);
        int b = Convert.ToInt32(hex[4..6], 16);
        return r | (g << 8) | (b << 16);
    }
}
