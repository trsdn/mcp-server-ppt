using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.SlideTable;

/// <summary>
/// Table shape operations: create, read, write cells, add/delete rows and columns, merge cells.
/// </summary>
[ServiceCategory("slidetable")]
[McpTool("slidetable", Title = "Table Operations", Destructive = true, Category = "tables")]
public interface ISlideTableCommands
{
    /// <summary>
    /// Create a table shape on a slide.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="rows">Number of rows</param>
    /// <param name="columns">Number of columns</param>
    /// <param name="left">Position from left in points</param>
    /// <param name="top">Position from top in points</param>
    /// <param name="width">Width in points</param>
    /// <param name="height">Height in points</param>
    [ServiceAction("create")]
    OperationResult Create(IPptBatch batch, int slideIndex, int rows, int columns, float left, float top, float width, float height);

    /// <summary>
    /// Read all data from a table shape.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the table shape</param>
    [ServiceAction("read")]
    SlideTableResult Read(IPptBatch batch, int slideIndex, string shapeName);

    /// <summary>
    /// Write a value to a specific table cell.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the table shape</param>
    /// <param name="row">1-based row index</param>
    /// <param name="column">1-based column index</param>
    /// <param name="value">Cell value to set</param>
    [ServiceAction("write-cell")]
    OperationResult WriteCell(IPptBatch batch, int slideIndex, string shapeName, int row, int column, string value);

    /// <summary>
    /// Add a row to the table.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the table shape</param>
    /// <param name="position">1-based position to insert (-1 = at end)</param>
    [ServiceAction("add-row")]
    OperationResult AddRow(IPptBatch batch, int slideIndex, string shapeName, int position);

    /// <summary>
    /// Add a column to the table.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the table shape</param>
    /// <param name="position">1-based position to insert (-1 = at end)</param>
    [ServiceAction("add-column")]
    OperationResult AddColumn(IPptBatch batch, int slideIndex, string shapeName, int position);

    /// <summary>
    /// Delete a row from the table.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the table shape</param>
    /// <param name="row">1-based row index to delete</param>
    [ServiceAction("delete-row")]
    OperationResult DeleteRow(IPptBatch batch, int slideIndex, string shapeName, int row);

    /// <summary>
    /// Delete a column from the table.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the table shape</param>
    /// <param name="column">1-based column index to delete</param>
    [ServiceAction("delete-column")]
    OperationResult DeleteColumn(IPptBatch batch, int slideIndex, string shapeName, int column);

    /// <summary>
    /// Merge a range of cells in a table.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the table shape</param>
    /// <param name="startRow">1-based start row</param>
    /// <param name="startColumn">1-based start column</param>
    /// <param name="endRow">1-based end row</param>
    /// <param name="endColumn">1-based end column</param>
    [ServiceAction("merge-cells")]
    OperationResult MergeCells(IPptBatch batch, int slideIndex, string shapeName, int startRow, int startColumn, int endRow, int endColumn);

    /// <summary>
    /// Set formatting on a table cell (fill color, text alignment).
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the table shape</param>
    /// <param name="row">1-based row index</param>
    /// <param name="column">1-based column index</param>
    /// <param name="fillColor">Hex fill color (#RRGGBB) or null to skip</param>
    /// <param name="fontBold">Set bold (null = don't change)</param>
    /// <param name="fontSize">Set font size (0 = don't change)</param>
    /// <param name="textAlign">Text alignment: left, center, right (null = don't change)</param>
    [ServiceAction("format-cell")]
    OperationResult FormatCell(IPptBatch batch, int slideIndex, string shapeName, int row, int column, string? fillColor, bool? fontBold, float fontSize, string? textAlign);
}
