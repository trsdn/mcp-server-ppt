using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.ShapeAlign;

public class ShapeAlignCommands : IShapeAlignCommands
{
    public OperationResult Align(IPptBatch batch, int slideIndex, string shapeNames, int alignType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeNames);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                string[] names = shapeNames.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                dynamic shapes = slide.Shapes;
                try
                {
                    // Build a ShapeRange from the named shapes
                    object[] nameArray = names.Cast<object>().ToArray();
                    dynamic range = shapes.Range(nameArray);
                    try
                    {
                        // msoAlignLeft=0, msoAlignCenter=1, msoAlignRight=2,
                        // msoAlignTop=3, msoAlignMiddle=4, msoAlignBottom=5
                        // RelativeTo: msoFalse=0 (relative to slide)
                        range.Align(alignType, 0);

                        string alignName = GetAlignTypeName(alignType);
                        return new OperationResult
                        {
                            Success = true,
                            Action = "align",
                            Message = $"Aligned {names.Length} shape(s) {alignName} on slide {slideIndex}",
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

    public OperationResult Distribute(IPptBatch batch, int slideIndex, string shapeNames, int distributeType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeNames);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            try
            {
                string[] names = shapeNames.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                dynamic shapes = slide.Shapes;
                try
                {
                    object[] nameArray = names.Cast<object>().ToArray();
                    dynamic range = shapes.Range(nameArray);
                    try
                    {
                        // msoDistributeHorizontally=0, msoDistributeVertically=1
                        range.Distribute(distributeType, 0);

                        string distName = distributeType == 0 ? "horizontally" : "vertically";
                        return new OperationResult
                        {
                            Success = true,
                            Action = "distribute",
                            Message = $"Distributed {names.Length} shape(s) {distName} on slide {slideIndex}",
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

    private static string GetAlignTypeName(int alignType) => alignType switch
    {
        0 => "left",
        1 => "center",
        2 => "right",
        3 => "top",
        4 => "middle",
        5 => "bottom",
        _ => $"type({alignType})"
    };
}
