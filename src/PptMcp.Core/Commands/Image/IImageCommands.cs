using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Image;

/// <summary>
/// Image operations: insert pictures into slides.
/// </summary>
[ServiceCategory("image")]
[McpTool("image", Title = "Image Operations", Destructive = true, Category = "media")]
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
}
