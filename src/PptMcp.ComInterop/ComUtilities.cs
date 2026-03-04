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


