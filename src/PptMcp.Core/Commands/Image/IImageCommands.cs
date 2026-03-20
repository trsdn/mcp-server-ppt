using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Image;

/// <summary>
/// Image operations: insert pictures into slides.
/// </summary>
[ServiceCategory("image")]
[McpTool("image", Title = "Image Operations", Destructive = true, Category = "media",
    Description = "Insert and adjust images on slides. Supports PNG, JPG, BMP, GIF, SVG, TIFF. "
    + "Positions in points (72pt = 1 inch). width/height: 0 = keep original size. "
    + "crop values in points from each edge. brightness/contrast: 0.0-1.0 (0.5 = normal).")]
public interface IImageCommands
{
    /// <summary>Insert a picture from a file path onto a slide.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="imagePath">Path to the image file</param>
    /// <param name="left">Position from left in points</param>
    /// <param name="top">Position from top in points</param>
    /// <param name="width">Width in points (0 = original)</param>
    /// <param name="height">Height in points (0 = original)</param>
    [ServiceAction("insert")]
    OperationResult Insert(IPptBatch batch, int slideIndex, string imagePath, float left, float top, float width, float height);

    /// <summary>Crop an image shape by specifying crop amounts on each side.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the picture shape</param>
    /// <param name="cropLeft">Crop from left in points (0 = no crop)</param>
    /// <param name="cropRight">Crop from right in points (0 = no crop)</param>
    /// <param name="cropTop">Crop from top in points (0 = no crop)</param>
    /// <param name="cropBottom">Crop from bottom in points (0 = no crop)</param>
    [ServiceAction("crop")]
    OperationResult Crop(IPptBatch batch, int slideIndex, string shapeName, float cropLeft, float cropRight, float cropTop, float cropBottom);

    /// <summary>Set brightness and contrast on a picture shape.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the picture shape</param>
    /// <param name="brightness">Brightness value (0.0 to 1.0)</param>
    /// <param name="contrast">Contrast value (0.0 to 1.0)</param>
    [ServiceAction("set-brightness-contrast")]
    OperationResult SetBrightnessContrast(IPptBatch batch, int slideIndex, string shapeName, float brightness, float contrast);

    /// <summary>Set a transparent color on a picture shape. Pixels matching this color become transparent.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="shapeName">Name of the picture shape</param>
    /// <param name="colorHex">Hex color string (#RRGGBB) to make transparent</param>
    [ServiceAction("set-transparent-color")]
    OperationResult SetTransparentColor(IPptBatch batch, int slideIndex, string shapeName, string colorHex);
}
