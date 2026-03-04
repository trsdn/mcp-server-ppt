using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Chart;

/// <summary>
/// Embedded chart operations: create, get info, set title, set type, delete.
/// </summary>
[ServiceCategory("chart")]
[McpTool("chart", Title = "Chart Operations", Destructive = true, Category = "charts")]
public interface IChartCommands
{
    /// <summary>
    /// Create an embedded chart on a slide.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="chartType">XlChartType integer (e.g., 4=xlLine, 5=xlPie, 51=xlColumnClustered, -4169=xl3DColumn)</param>
    /// <param name="left">Position from left in points</param>
    /// <param name="top">Position from top in points</param>
    /// <param name="width">Width in points</param>
    /// <param name="height">Height in points</param>
    [ServiceAction("create")]
    OperationResult Create(IPptBatch batch, int slideIndex, int chartType, float left, float top, float width, float height);

    /// <summary>
    /// Get information about a chart shape.
    /// </summary>
    [ServiceAction("get-info")]
    ChartInfoResult GetInfo(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// Set the chart title.
    /// </summary>
    [ServiceAction("set-title")]
    OperationResult SetTitle(IPptBatch batch, int slideIndex, string shapeName, string title);

    /// <summary>
    /// Change the chart type.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the chart shape</param>
    /// <param name="chartType">XlChartType integer</param>
    [ServiceAction("set-type")]
    OperationResult SetType(IPptBatch batch, int slideIndex, string shapeName, int chartType);

    /// <summary>
    /// Delete a chart shape.
    /// </summary>
    [ServiceAction("delete")]
    OperationResult Delete(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// Set chart data from a 2D array of values.
    /// Opens the embedded data worksheet, writes values, then closes.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the chart shape</param>
    /// <param name="values">2D array of values (rows × columns)</param>
    [ServiceAction("set-data")]
    OperationResult SetData(IPptBatch batch, int slideIndex, string shapeName, List<List<object?>> values);

    /// <summary>
    /// Set chart legend visibility and position.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the chart shape</param>
    /// <param name="visible">Whether the legend is visible</param>
    /// <param name="position">Legend position: -4107=Bottom, -4131=Left, -4152=Right, -4160=Top, -4161=TopRight</param>
    [ServiceAction("set-legend")]
    OperationResult SetLegend(IPptBatch batch, int slideIndex, string shapeName, bool visible, int position);

    /// <summary>
    /// Read chart data from the embedded data worksheet and return as a text grid.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the chart shape</param>
    [ServiceAction("read-data")]
    OperationResult ReadData(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// Set the title of a chart axis.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the chart shape</param>
    /// <param name="axisType">Axis type: 1=Category(X), 2=Value(Y)</param>
    /// <param name="title">Title text for the axis</param>
    [ServiceAction("set-axis-title")]
    OperationResult SetAxisTitle(IPptBatch batch, int slideIndex, string shapeName, int axisType, string title);

    /// <summary>
    /// Show or hide the chart data table.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the chart shape</param>
    /// <param name="visible">Whether the data table is visible</param>
    [ServiceAction("toggle-data-table")]
    OperationResult ToggleDataTable(IPptBatch batch, int slideIndex, string shapeName, bool visible);
}
