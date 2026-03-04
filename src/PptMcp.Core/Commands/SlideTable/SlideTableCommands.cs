using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.SlideTable;

public class SlideTableCommands : ISlideTableCommands
{
    public OperationResult Create(IPptBatch batch, int slideIndex, int rows, int columns, float left, float top, float width, float height)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.AddTable(rows, columns, left, top, width, height);
            try
            {
                string name = shape.Name?.ToString() ?? "";
                return new OperationResult
                {
                    Success = true,
                    Action = "create",
                    Message = $"Created table '{name}' ({rows}x{columns}) on slide {slideIndex}",
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

    public SlideTableResult Read(IPptBatch batch, int slideIndex, string shapeName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? table = null;
            try
            {
                table = shape.Table;
                int rowCount = (int)table.Rows.Count;
                int colCount = (int)table.Columns.Count;

                var data = new List<List<string?>>();
                for (int r = 1; r <= rowCount; r++)
                {
                    var row = new List<string?>();
                    for (int c = 1; c <= colCount; c++)
                    {
                        dynamic cell = table.Cell(r, c);
                        try
                        {
                            string? text = cell.Shape.TextFrame.TextRange.Text?.ToString();
                            row.Add(text);
                        }
                        finally
                        {
                            ComUtilities.Release(ref cell!);
                        }
                    }
                    data.Add(row);
                }

                return new SlideTableResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath,
                    ShapeId = (int)shape.Id,
                    ShapeName = shape.Name?.ToString() ?? "",
                    RowCount = rowCount,
                    ColumnCount = colCount,
                    Data = data
                };
            }
            finally
            {
                if (table != null) ComUtilities.Release(ref table!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult WriteCell(IPptBatch batch, int slideIndex, string shapeName, int row, int column, string value)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? table = null;
            dynamic? cell = null;
            try
            {
                table = shape.Table;
                cell = table.Cell(row, column);
                cell.Shape.TextFrame.TextRange.Text = value;

                return new OperationResult
                {
                    Success = true,
                    Action = "write-cell",
                    Message = $"Set cell ({row},{column}) in table '{shapeName}' on slide {slideIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (cell != null) ComUtilities.Release(ref cell!);
                if (table != null) ComUtilities.Release(ref table!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult AddRow(IPptBatch batch, int slideIndex, string shapeName, int position)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? table = null;
            dynamic? rows = null;
            try
            {
                table = shape.Table;
                rows = table.Rows;
                int insertAt = position <= 0 ? (int)rows.Count : position;
                rows.Add(insertAt);

                return new OperationResult
                {
                    Success = true,
                    Action = "add-row",
                    Message = $"Added row at position {insertAt} in table '{shapeName}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (rows != null) ComUtilities.Release(ref rows!);
                if (table != null) ComUtilities.Release(ref table!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult AddColumn(IPptBatch batch, int slideIndex, string shapeName, int position)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? table = null;
            dynamic? columns = null;
            try
            {
                table = shape.Table;
                columns = table.Columns;
                int insertAt = position <= 0 ? (int)columns.Count : position;
                columns.Add(insertAt);

                return new OperationResult
                {
                    Success = true,
                    Action = "add-column",
                    Message = $"Added column at position {insertAt} in table '{shapeName}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (columns != null) ComUtilities.Release(ref columns!);
                if (table != null) ComUtilities.Release(ref table!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult DeleteRow(IPptBatch batch, int slideIndex, string shapeName, int row)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? table = null;
            dynamic? targetRow = null;
            try
            {
                table = shape.Table;
                targetRow = table.Rows.Item(row);
                targetRow.Delete();

                return new OperationResult
                {
                    Success = true,
                    Action = "delete-row",
                    Message = $"Deleted row {row} from table '{shapeName}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (targetRow != null) ComUtilities.Release(ref targetRow!);
                if (table != null) ComUtilities.Release(ref table!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult DeleteColumn(IPptBatch batch, int slideIndex, string shapeName, int column)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? table = null;
            dynamic? targetCol = null;
            try
            {
                table = shape.Table;
                targetCol = table.Columns.Item(column);
                targetCol.Delete();

                return new OperationResult
                {
                    Success = true,
                    Action = "delete-column",
                    Message = $"Deleted column {column} from table '{shapeName}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (targetCol != null) ComUtilities.Release(ref targetCol!);
                if (table != null) ComUtilities.Release(ref table!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }

    public OperationResult MergeCells(IPptBatch batch, int slideIndex, string shapeName, int startRow, int startColumn, int endRow, int endColumn)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
            dynamic shape = slide.Shapes.Item(shapeName);
            dynamic? table = null;
            dynamic? cell1 = null;
            dynamic? cell2 = null;
            try
            {
                table = shape.Table;
                cell1 = table.Cell(startRow, startColumn);
                cell2 = table.Cell(endRow, endColumn);
                cell1.Merge(cell2);

                return new OperationResult
                {
                    Success = true,
                    Action = "merge-cells",
                    Message = $"Merged cells ({startRow},{startColumn})-({endRow},{endColumn}) in table '{shapeName}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                if (cell2 != null) ComUtilities.Release(ref cell2!);
                if (cell1 != null) ComUtilities.Release(ref cell1!);
                if (table != null) ComUtilities.Release(ref table!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref slide!);
            }
        });
    }
}
