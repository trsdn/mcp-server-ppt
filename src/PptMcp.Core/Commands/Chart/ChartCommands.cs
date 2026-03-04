using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Chart;

public class ChartCommands : IChartCommands
{
    public OperationResult Create(IPptBatch batch, int slideIndex, int chartType, float left, float top, float width, float height)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic? shape = null;
            try
            {
                // AddChart(Type, Left, Top, Width, Height)
                shape = slide.Shapes.AddChart(chartType, left, top, width, height);
                string name = shape?.Name?.ToString() ?? "";
                return new OperationResult
                {
                    Success = true,
                    Action = "create",
                    Message = $"Created chart '{name}' (type {chartType}) on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public ChartInfoResult GetInfo(IPptBatch batch, int slideIndex, string shapeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? chart = null;
            try
            {
                chart = shape.Chart;
                string? title = null;
                try
                {
                    if ((bool)chart.HasTitle)
                        title = chart.ChartTitle.Text?.ToString();
                }
                catch { /* Title not accessible */ }

                bool hasLegend = false;
                try { hasLegend = (bool)chart.HasLegend; } catch { }

                int seriesCount = 0;
                try
                {
                    dynamic seriesCol = chart.SeriesCollection();
                    seriesCount = (int)seriesCol.Count;
                    ComUtilities.Release(ref seriesCol!);
                }
                catch { }

                int chartTypeVal = Convert.ToInt32(chart.ChartType);

                return new ChartInfoResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    ShapeId = (int)shape.Id,
                    ShapeName = shape.Name?.ToString() ?? "",
                    ChartType = chartTypeVal,
                    ChartTypeName = GetChartTypeName(chartTypeVal),
                    Title = title,
                    HasLegend = hasLegend,
                    SeriesCount = seriesCount
                };
            }
            finally
            {
                if (chart != null) ComUtilities.Release(ref chart!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetTitle(IPptBatch batch, int slideIndex, string shapeName, string title)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? chart = null;
            try
            {
                chart = shape.Chart;
                chart.HasTitle = true;
                chart.ChartTitle.Text = title;

                return new OperationResult
                {
                    Success = true,
                    Action = "set-title",
                    Message = $"Set chart title to '{title}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (chart != null) ComUtilities.Release(ref chart!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult SetType(IPptBatch batch, int slideIndex, string shapeName, int chartType)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shapeName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? chart = null;
            try
            {
                chart = shape.Chart;
                chart.ChartType = chartType;

                return new OperationResult
                {
                    Success = true,
                    Action = "set-type",
                    Message = $"Changed chart type to {chartType} on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (chart != null) ComUtilities.Release(ref chart!);
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
                    Message = $"Deleted chart shape '{shapeName}' from slide {slideIndex}",
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

    private static string GetChartTypeName(int chartType) => chartType switch
    {
        1 => "xlArea",
        4 => "xlLine",
        5 => "xlPie",
        51 => "xlColumnClustered",
        52 => "xlColumnStacked",
        54 => "xlBarClustered",
        65 => "xlBarStacked",
        72 => "xlDoughnut",
        -4169 => "xl3DColumn",
        -4120 => "xlXYScatter",
        _ => $"Unknown({chartType})"
    };
}
