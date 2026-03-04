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

    public OperationResult CopyToSlide(IPptBatch batch, int slideIndex, string shapeName, int targetSlideIndex)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            dynamic srcSlide = pres.Slides.Item(slideIndex);
            dynamic shape = srcSlide.Shapes.Item(shapeName);
            try
            {
                shape.Copy();
                dynamic targetSlide = pres.Slides.Item(targetSlideIndex);
                dynamic pasted = targetSlide.Shapes.Paste();
                string newName = "";
                try { newName = pasted.Item(1).Name?.ToString() ?? ""; } catch { }
                ComUtilities.Release(ref pasted!);
                ComUtilities.Release(ref targetSlide!);

                return new OperationResult
                {
                    Success = true,
                    Action = "copy-to-slide",
                    Message = $"Copied shape '{shapeName}' from slide {slideIndex} to slide {targetSlideIndex}" +
                              (string.IsNullOrEmpty(newName) ? "" : $" as '{newName}'"),
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref srcSlide!);
            }
        });
    }

    public OperationResult SetShadow(IPptBatch batch, int slideIndex, string shapeName, bool visible, float offsetX, float offsetY)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                dynamic shadow = shape.Shadow;
                try
                {
                    shadow.Visible = visible ? -1 : 0;
                    if (visible)
                    {
                        shadow.OffsetX = offsetX;
                        shadow.OffsetY = offsetY;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref shadow!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "set-shadow",
                    Message = visible
                        ? $"Set shadow on shape '{shapeName}' (offset {offsetX},{offsetY})"
                        : $"Removed shadow from shape '{shapeName}'",
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

    public OperationResult AddConnector(IPptBatch batch, int slideIndex, int connectorType, string startShapeName, string endShapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(startShapeName);
        ArgumentException.ThrowIfNullOrWhiteSpace(endShapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic? startShape = null;
            dynamic? endShape = null;
            dynamic? connector = null;
            try
            {
                startShape = slide.Shapes.Item(startShapeName);
                endShape = slide.Shapes.Item(endShapeName);

                // AddConnector(Type, BeginX, BeginY, EndX, EndY)
                // Type: 1=msoConnectorStraight, 2=msoConnectorElbow, 3=msoConnectorCurve
                float sx = Convert.ToSingle(startShape.Left) + Convert.ToSingle(startShape.Width) / 2;
                float sy = Convert.ToSingle(startShape.Top) + Convert.ToSingle(startShape.Height) / 2;
                float ex = Convert.ToSingle(endShape.Left) + Convert.ToSingle(endShape.Width) / 2;
                float ey = Convert.ToSingle(endShape.Top) + Convert.ToSingle(endShape.Height) / 2;

                connector = slide.Shapes.AddConnector(connectorType, sx, sy, ex, ey);
                dynamic cf = connector.ConnectorFormat;
                try
                {
                    cf.BeginConnect(startShape, 1);
                    cf.EndConnect(endShape, 1);
                }
                finally
                {
                    ComUtilities.Release(ref cf!);
                }

                string name = connector.Name?.ToString() ?? "";
                return new OperationResult
                {
                    Success = true,
                    Action = "add-connector",
                    Message = $"Added connector '{name}' between '{startShapeName}' and '{endShapeName}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (connector != null) ComUtilities.Release(ref connector!);
                if (endShape != null) ComUtilities.Release(ref endShape!);
                if (startShape != null) ComUtilities.Release(ref startShape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult MergeShapes(IPptBatch batch, int slideIndex, string shapeNames, int mergeType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeNames);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                string[] names = shapeNames.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                if (names.Length < 2)
                    throw new ArgumentException("At least 2 shape names are required for merge.");

                dynamic shapes = slide.Shapes;
                try
                {
                    object[] nameArray = names.Cast<object>().ToArray();
                    dynamic range = shapes.Range(nameArray);
                    try
                    {
                        // MergeShapes: 1=Union, 2=Combine, 3=Fragment, 4=Intersect, 5=Subtract
                        range.MergeShapes(mergeType);

                        string mergeName = mergeType switch
                        {
                            1 => "union",
                            2 => "combine",
                            3 => "fragment",
                            4 => "intersect",
                            5 => "subtract",
                            _ => $"type({mergeType})"
                        };

                        return new OperationResult
                        {
                            Success = true,
                            Action = "merge",
                            Message = $"Merged {names.Length} shapes using {mergeName} on slide {slideIndex}",
                            FilePath = ctx.PresentationPath
                        };
                    }
                    finally
                    {
                        ComUtilities.Release(ref range!);
                    }
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

    public OperationResult Duplicate(IPptBatch batch, int slideIndex, string shapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? dup = null;
            try
            {
                dup = shape.Duplicate();
                // Duplicate returns a ShapeRange; get first item
                dynamic newShape = dup.Item(1);
                string newName = newShape.Name?.ToString() ?? "";
                ComUtilities.Release(ref newShape!);

                return new OperationResult
                {
                    Success = true,
                    Action = "duplicate",
                    Message = $"Duplicated shape '{shapeName}' as '{newName}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (dup != null) ComUtilities.Release(ref dup!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult Flip(IPptBatch batch, int slideIndex, string shapeName, int flipType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                // msoFlipHorizontal=0, msoFlipVertical=1
                shape.Flip(flipType);
                string dir = flipType == 0 ? "horizontally" : "vertically";
                return new OperationResult
                {
                    Success = true,
                    Action = "flip",
                    Message = $"Flipped shape '{shapeName}' {dir} on slide {slideIndex}",
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

    public OperationResult SetTextFrame(IPptBatch batch, int slideIndex, string shapeName, float? marginLeft, float? marginRight, float? marginTop, float? marginBottom, bool? wordWrap, int? autoSize)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? textFrame = null;
            try
            {
                textFrame = shape.TextFrame;
                if (marginLeft.HasValue) textFrame.MarginLeft = marginLeft.Value;
                if (marginRight.HasValue) textFrame.MarginRight = marginRight.Value;
                if (marginTop.HasValue) textFrame.MarginTop = marginTop.Value;
                if (marginBottom.HasValue) textFrame.MarginBottom = marginBottom.Value;
                if (wordWrap.HasValue) textFrame.WordWrap = wordWrap.Value ? -1 : 0; // msoTrue=-1, msoFalse=0
                if (autoSize.HasValue) textFrame.AutoSize = autoSize.Value; // ppAutoSizeNone=0, ppAutoSizeShapeToFitText=1, ppAutoSizeTextToFitShape=2

                return new OperationResult
                {
                    Success = true,
                    Action = "set-text-frame",
                    Message = $"Updated text frame properties of shape '{shapeName}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (textFrame != null) ComUtilities.Release(ref textFrame!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

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

    public OperationResult SetGradientFill(IPptBatch batch, int slideIndex, string shapeName, string color1, string color2, int gradientStyle)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);
        ArgumentException.ThrowIfNullOrWhiteSpace(color1);
        ArgumentException.ThrowIfNullOrWhiteSpace(color2);

        if (gradientStyle < 1 || gradientStyle > 6)
            throw new ArgumentOutOfRangeException(nameof(gradientStyle), "gradientStyle must be 1-6 (1=Horizontal, 2=Vertical, 3=DiagonalUp, 4=DiagonalDown, 5=FromCorner, 6=FromCenter)");

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                // TwoColorGradient(style, variant) - variant 1 is default direction
                shape.Fill.TwoColorGradient(gradientStyle, 1);
                shape.Fill.ForeColor.RGB = HexToOleColor(color1);
                shape.Fill.BackColor.RGB = HexToOleColor(color2);

                return new OperationResult
                {
                    Success = true,
                    Action = "set-gradient-fill",
                    Message = $"Set gradient fill on shape '{shapeName}' from '{color1}' to '{color2}' (style {gradientStyle}) on slide {slideIndex}",
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

    public OperationResult SetGlow(IPptBatch batch, int slideIndex, string shapeName, float radius, string colorHex)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);
        ArgumentException.ThrowIfNullOrWhiteSpace(colorHex);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                dynamic glow = shape.Glow;
                try
                {
                    glow.Radius = radius;
                    if (radius > 0)
                    {
                        glow.Color.RGB = HexToOleColor(colorHex);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref glow!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "set-glow",
                    Message = radius > 0
                        ? $"Set glow on shape '{shapeName}' with radius {radius}pt and color '{colorHex}' on slide {slideIndex}"
                        : $"Removed glow from shape '{shapeName}' on slide {slideIndex}",
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

    public OperationResult SetReflection(IPptBatch batch, int slideIndex, string shapeName, int reflectionType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        if (reflectionType < 0 || reflectionType > 9)
            throw new ArgumentOutOfRangeException(nameof(reflectionType), "reflectionType must be 0-9 (0=None, 1-9=msoReflectionType1 through msoReflectionType9)");

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                dynamic reflection = shape.Reflection;
                try
                {
                    reflection.Type = reflectionType;
                }
                finally
                {
                    ComUtilities.Release(ref reflection!);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "set-reflection",
                    Message = reflectionType > 0
                        ? $"Set reflection type {reflectionType} on shape '{shapeName}' on slide {slideIndex}"
                        : $"Removed reflection from shape '{shapeName}' on slide {slideIndex}",
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

    public OperationResult SetOpacity(IPptBatch batch, int slideIndex, string shapeName, float opacity)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        if (opacity < 0.0f || opacity > 1.0f)
            throw new ArgumentOutOfRangeException(nameof(opacity), "opacity must be between 0.0 (transparent) and 1.0 (opaque)");

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            try
            {
                // COM uses Transparency (0=opaque, 1=transparent), which is the inverse of opacity
                shape.Fill.Transparency = 1.0f - opacity;
                return new OperationResult
                {
                    Success = true,
                    Action = "set-opacity",
                    Message = $"Set opacity of shape '{shapeName}' to {opacity:F2} on slide {slideIndex}",
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
}
