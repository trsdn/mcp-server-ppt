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


