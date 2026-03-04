using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Export;

/// <summary>
/// Export presentations to PDF, images, or other formats.
/// </summary>
[ServiceCategory("export")]
[McpTool("export", Title = "Export Operations", Destructive = false, Category = "export")]
public interface IExportCommands
{
    /// <summary>Export the presentation to PDF.</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="destinationPath">Output PDF file path</param>
    [ServiceAction("to-pdf")]
    ExportResult ToPdf(IPptBatch batch, string destinationPath);

    /// <summary>Export a single slide as an image (PNG).</summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="destinationPath">Output image file path</param>
    /// <param name="width">Image width in pixels (default: 1920)</param>
    /// <param name="height">Image height in pixels (default: 1080)</param>
    [ServiceAction("slide-to-image")]
    ExportResult SlideToImage(IPptBatch batch, int slideIndex, string destinationPath, int width, int height);

    /// <summary>
    /// Export the presentation as a video (MP4).
    /// Resolution: 1=FullHD(1080p), 2=HD(720p), 3=Standard(480p).
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="destinationPath">Output video file path (.mp4)</param>
    /// <param name="defaultSlideSeconds">Seconds per slide (default: 5)</param>
    /// <param name="resolution">1=1080p, 2=720p, 3=480p</param>
    [ServiceAction("to-video")]
    ExportResult ToVideo(IPptBatch batch, string destinationPath, int defaultSlideSeconds, int resolution);

    /// <summary>
    /// Print the presentation.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="copies">Number of copies (default: 1)</param>
    /// <param name="fromSlide">First slide to print (0 = from beginning)</param>
    /// <param name="toSlide">Last slide to print (0 = to end)</param>
    [ServiceAction("print")]
    OperationResult Print(IPptBatch batch, int copies, int fromSlide, int toSlide);

    /// <summary>
    /// Save the presentation as a different format.
    /// Format: 1=pptx, 2=pptm (macro-enabled), 3=potx (template),
    /// 4=ppsx (show), 5=pdf, 6=xps, 7=odp (OpenDocument).
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="destinationPath">Output file path</param>
    /// <param name="format">Format code (1-7)</param>
    [ServiceAction("save-as")]
    ExportResult SaveAs(IPptBatch batch, string destinationPath, int format);

    /// <summary>
    /// Export all slides as individual PNG images (slide_001.png, slide_002.png, etc.).
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="destinationDirectory">Directory to save images</param>
    /// <param name="width">Image width in pixels (default: 1920)</param>
    /// <param name="height">Image height in pixels (default: 1080)</param>
    [ServiceAction("all-slides-to-images")]
    ExportResult AllSlidesToImages(IPptBatch batch, string destinationDirectory, int width, int height);

    /// <summary>
    /// Extract all text from the presentation to a text file.
    /// Iterates all slides and shapes, writing text with slide headers.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="destinationPath">Output text file path</param>
    [ServiceAction("extract-text")]
    OperationResult ExtractText(IPptBatch batch, string destinationPath);

    /// <summary>
    /// Extract all images (pictures) from the presentation as PNG files.
    /// Exports shapes of type Picture (13) or LinkedPicture (11).
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="destinationDirectory">Directory to save extracted images</param>
    [ServiceAction("extract-images")]
    OperationResult ExtractImages(IPptBatch batch, string destinationDirectory);

    /// <summary>
    /// Save a copy of the presentation to a new path without changing the current file.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="destinationPath">Full path for the copy</param>
    [ServiceAction("save-copy")]
    ExportResult SaveCopy(IPptBatch batch, string destinationPath);
}
