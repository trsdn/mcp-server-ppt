using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptMcp.ComInterop;

/// <summary>
/// Low-level COM interop utilities for PowerPoint automation.
/// Provides helpers for managing COM object lifecycle.
/// </summary>
public static class ComUtilities
{
    /// <summary>
    /// Safely releases a COM object and sets the reference to null
    /// </summary>
    /// <param name="comObject">The COM object to release</param>
    /// <remarks>
    /// Use this helper to release intermediate COM objects (like slides, shapes)
    /// to prevent PowerPoint process from staying open. This is especially important when
    /// iterating through collections or accessing multiple COM properties.
    /// </remarks>
    /// <example>
    /// <code>
    /// dynamic? slides = null;
    /// try
    /// {
    ///     slides = presentation.Slides;
    ///     // Use slides...
    /// }
    /// finally
    /// {
    ///     ComUtilities.Release(ref slides);
    /// }
    /// </code>
    /// </example>
    public static void Release<T>(ref T? comObject) where T : class
    {
        if (comObject != null)
        {
            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch (Exception)
            {
                // Ignore errors during release — COM object may already be released or RPC disconnected
            }
            comObject = null;
        }
    }

    /// <summary>
    /// Safely attempts to quit a PowerPoint application COM object.
    /// This is a fire-and-forget cleanup helper - errors are swallowed.
    /// </summary>
    /// <param name="powerPoint">The PowerPoint.Application COM object</param>
    /// <remarks>
    /// Use this for cleanup scenarios where you want to quit PowerPoint but don't
    /// need to handle or report errors. For production shutdown with retry
    /// logic, use PptShutdownService.CloseAndQuit instead.
    /// </remarks>
    public static void TryQuitPowerPoint(PowerPoint.Application? powerPoint)
    {
        if (powerPoint == null) return;

        try
        {
            powerPoint.Quit();
        }
        catch (Exception)
        {
            // Swallow errors during cleanup — PowerPoint may already be gone
        }
    }

    /// <summary>
    /// Gets a slide from a presentation without abandoning the intermediate
    /// <c>Slides</c> collection.
    ///
    /// <para>
    /// Writing <c>presentation.Slides.Item(i)</c> inline materialises a <c>Slides</c>
    /// RCW that is never bound to a local, so <see cref="Release{T}"/> cannot be called
    /// on it even in principle. The proxy survives until a garbage collection that may
    /// never come while the STA thread is alive, which is the leak class tracked by
    /// GitHub #137 - and, per #148, a leaked collection proxy is enough to stop the STA
    /// thread exiting and turn session teardown into a 45-second timeout.
    /// </para>
    /// </summary>
    /// <param name="presentation">Presentation COM object.</param>
    /// <param name="slideIndex">1-based slide index. PowerPoint collections are not zero-based.</param>
    /// <returns>The slide COM object. The caller owns it and must release it.</returns>
    public static dynamic GetSlide(dynamic presentation, int slideIndex)
    {
        dynamic? slides = null;
        try
        {
            slides = presentation.Slides;
            return slides.Item(slideIndex);
        }
        finally
        {
            ComUtilities.Release(ref slides!);
        }
    }

    /// <summary>
    /// Gets a shape from a slide by name without abandoning the intermediate
    /// <c>Shapes</c> collection. See <see cref="GetSlide"/> for why the inline form
    /// leaks.
    /// </summary>
    /// <param name="slide">Slide COM object.</param>
    /// <param name="shapeName">Shape name.</param>
    /// <returns>The shape COM object. The caller owns it and must release it.</returns>
    public static dynamic GetShape(dynamic slide, string shapeName)
    {
        dynamic? shapes = null;
        try
        {
            shapes = slide.Shapes;
            return shapes.Item(shapeName);
        }
        finally
        {
            ComUtilities.Release(ref shapes!);
        }
    }

    /// <summary>
    /// Gets a shape from a slide by 1-based position. See <see cref="GetSlide"/> for
    /// why the inline form leaks.
    /// </summary>
    /// <param name="slide">Slide COM object.</param>
    /// <param name="shapeIndex">1-based shape index.</param>
    /// <returns>The shape COM object. The caller owns it and must release it.</returns>
    public static dynamic GetShapeAt(dynamic slide, int shapeIndex)
    {
        dynamic? shapes = null;
        try
        {
            shapes = slide.Shapes;
            return shapes.Item(shapeIndex);
        }
        finally
        {
            ComUtilities.Release(ref shapes!);
        }
    }

    /// <summary>
    /// Gets a shape's <c>TextRange</c> without abandoning the intermediate
    /// <c>TextFrame</c>. See <see cref="GetSlide"/> for why the inline form leaks.
    ///
    /// <para>
    /// The caller is responsible for checking <c>HasTextFrame</c> first; this method
    /// deliberately does not guard, because a shape that cannot hold text is a caller
    /// error rather than something to paper over with a null return.
    /// </para>
    /// </summary>
    /// <param name="shape">Shape COM object.</param>
    /// <returns>The text range COM object. The caller owns it and must release it.</returns>
    public static dynamic GetTextRange(dynamic shape)
    {
        dynamic? textFrame = null;
        try
        {
            textFrame = shape.TextFrame;
            return textFrame.TextRange;
        }
        finally
        {
            ComUtilities.Release(ref textFrame!);
        }
    }

    /// <summary>
    /// Gets a shape's text <c>Font</c> without abandoning the intermediate
    /// <c>TextFrame</c> and <c>TextRange</c>. See <see cref="GetSlide"/> for why the
    /// inline form leaks.
    /// </summary>
    /// <param name="shape">Shape COM object.</param>
    /// <returns>The font COM object. The caller owns it and must release it.</returns>
    public static dynamic GetTextFont(dynamic shape)
    {
        dynamic? textRange = null;
        try
        {
            textRange = GetTextRange(shape);
            return textRange.Font;
        }
        finally
        {
            ComUtilities.Release(ref textRange!);
        }
    }

    /// <summary>
    /// Reads the text of a table cell without abandoning the intermediate
    /// <c>Shape</c>, <c>TextFrame</c> and <c>TextRange</c> proxies.
    ///
    /// <para>
    /// The inline form <c>cell.Shape.TextFrame.TextRange.Text</c> abandons three RCWs
    /// per call, and table reads call it once per cell - so the leak scales with the
    /// area of the table. See <see cref="GetSlide"/> for why the inline form cannot be
    /// released.
    /// </para>
    /// </summary>
    /// <param name="cell">Table cell COM object.</param>
    /// <returns>The cell text, or an empty string when the cell has none.</returns>
    public static string GetCellText(dynamic cell)
    {
        dynamic? cellShape = null;
        dynamic? textRange = null;
        try
        {
            cellShape = cell.Shape;
            textRange = GetTextRange(cellShape);
            return textRange.Text?.ToString() ?? string.Empty;
        }
        finally
        {
            ComUtilities.Release(ref textRange!);
            ComUtilities.Release(ref cellShape!);
        }
    }

    /// <summary>
    /// Writes the text of a table cell without abandoning the intermediate
    /// <c>Shape</c>, <c>TextFrame</c> and <c>TextRange</c> proxies. See
    /// <see cref="GetCellText"/>.
    /// </summary>
    /// <param name="cell">Table cell COM object.</param>
    /// <param name="text">Text to write.</param>
    public static void SetCellText(dynamic cell, string text)
    {
        dynamic? cellShape = null;
        dynamic? textRange = null;
        try
        {
            cellShape = cell.Shape;
            textRange = GetTextRange(cellShape);
            textRange.Text = text;
        }
        finally
        {
            ComUtilities.Release(ref textRange!);
            ComUtilities.Release(ref cellShape!);
        }
    }

    /// <summary>
    /// Gets a table row by 1-based index without abandoning the intermediate
    /// <c>Rows</c> collection. See <see cref="GetSlide"/>.
    /// </summary>
    /// <param name="table">Table COM object.</param>
    /// <param name="rowIndex">1-based row index.</param>
    /// <returns>The row COM object. The caller owns it and must release it.</returns>
    public static dynamic GetTableRow(dynamic table, int rowIndex)
    {
        dynamic? rows = null;
        try
        {
            rows = table.Rows;
            return rows.Item(rowIndex);
        }
        finally
        {
            ComUtilities.Release(ref rows!);
        }
    }

    /// <summary>
    /// Gets a table column by 1-based index without abandoning the intermediate
    /// <c>Columns</c> collection. See <see cref="GetSlide"/>.
    /// </summary>
    /// <param name="table">Table COM object.</param>
    /// <param name="columnIndex">1-based column index.</param>
    /// <returns>The column COM object. The caller owns it and must release it.</returns>
    public static dynamic GetTableColumn(dynamic table, int columnIndex)
    {
        dynamic? columns = null;
        try
        {
            columns = table.Columns;
            return columns.Item(columnIndex);
        }
        finally
        {
            ComUtilities.Release(ref columns!);
        }
    }

    /// <summary>
    /// Gets a table's row count without abandoning the intermediate <c>Rows</c>
    /// collection. See <see cref="GetSlide"/>.
    /// </summary>
    /// <param name="table">Table COM object.</param>
    /// <returns>Number of rows.</returns>
    public static int GetTableRowCount(dynamic table)
    {
        dynamic? rows = null;
        try
        {
            rows = table.Rows;
            return (int)rows.Count;
        }
        finally
        {
            ComUtilities.Release(ref rows!);
        }
    }

    /// <summary>
    /// Gets a table's column count without abandoning the intermediate <c>Columns</c>
    /// collection. See <see cref="GetSlide"/>.
    /// </summary>
    /// <param name="table">Table COM object.</param>
    /// <returns>Number of columns.</returns>
    public static int GetTableColumnCount(dynamic table)
    {
        dynamic? columns = null;
        try
        {
            columns = table.Columns;
            return (int)columns.Count;
        }
        finally
        {
            ComUtilities.Release(ref columns!);
        }
    }

    /// <summary>
    /// Reads the <c>ForeColor.RGB</c> of a fill or line without abandoning the
    /// intermediate <c>ForeColor</c> proxy. See <see cref="GetSlide"/>.
    /// </summary>
    /// <param name="fillOrLine">A <c>FillFormat</c> or <c>LineFormat</c> COM object.</param>
    /// <returns>The OLE colour value, in PowerPoint's 0x00BBGGRR order.</returns>
    public static int GetForeColorRgb(dynamic fillOrLine)
    {
        dynamic? foreColor = null;
        try
        {
            foreColor = fillOrLine.ForeColor;
            return Convert.ToInt32(foreColor.RGB);
        }
        finally
        {
            ComUtilities.Release(ref foreColor!);
        }
    }

    /// <summary>
    /// Reads a chart's title text without abandoning the intermediate
    /// <c>ChartTitle</c> proxy. See <see cref="GetSlide"/>.
    /// </summary>
    /// <param name="chart">Chart COM object.</param>
    /// <returns>The chart title text, or null when it has none.</returns>
    public static string? GetChartTitleText(dynamic chart)
    {
        dynamic? chartTitle = null;
        try
        {
            chartTitle = chart.ChartTitle;
            return chartTitle.Text?.ToString();
        }
        finally
        {
            ComUtilities.Release(ref chartTitle!);
        }
    }

    /// <summary>
    /// Sets a chart's title text without abandoning the intermediate
    /// <c>ChartTitle</c> proxy. See <see cref="GetSlide"/>.
    /// </summary>
    /// <param name="chart">Chart COM object.</param>
    /// <param name="title">Title text.</param>
    public static void SetChartTitleText(dynamic chart, string title)
    {
        dynamic? chartTitle = null;
        try
        {
            chartTitle = chart.ChartTitle;
            chartTitle.Text = title;
        }
        finally
        {
            ComUtilities.Release(ref chartTitle!);
        }
    }

    /// <summary>
    /// Safely gets a string property from a COM object, returning empty string if null
    /// </summary>
    /// <param name="obj">COM object</param>
    /// <param name="propertyName">Property name</param>
    /// <returns>Property value or empty string</returns>
    public static string SafeGetString(dynamic? obj, string propertyName)
    {
        try
        {
            var value = propertyName switch
            {
                "Name" => obj.Name,
                "Description" => obj.Description,
                _ => null
            };
            return value?.ToString() ?? string.Empty;
        }
        catch (Exception)
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Safely gets an integer property from a COM object, returning 0 if null or invalid
    /// </summary>
    /// <param name="obj">COM object</param>
    /// <param name="propertyName">Property name</param>
    /// <returns>Property value or 0</returns>
    public static int SafeGetInt(dynamic? obj, string propertyName)
    {
        try
        {
            var value = propertyName switch
            {
                "Count" => obj.Count,
                _ => 0
            };
            return Convert.ToInt32(value);
        }
        catch (Exception)
        {
            return 0;
        }
    }

    [DllImport("kernel32.dll")]
    private static extern void Sleep(uint dwMilliseconds);

    /// <summary>
    /// Kernel-level sleep that does NOT pump the STA COM message queue.
    /// Unlike Thread.Sleep (which uses CoWaitForMultipleHandles internally and wakes early on
    /// every incoming COM event), this calls Win32 Sleep() directly via NtDelayExecution —
    /// the thread genuinely sleeps for the full interval regardless of COM callbacks.
    /// </summary>
    public static void KernelSleep(int milliseconds) =>
        Sleep((uint)Math.Max(0, milliseconds));
}


