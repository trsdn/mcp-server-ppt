using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Media;

/// <summary>
/// Media management: insert audio and video files into slides.
/// Supports linking or embedding media files.
/// </summary>
[ServiceCategory("media")]
[McpTool("media", Title = "Media Operations", Destructive = true, Category = "content")]
public interface IMediaCommands
{
    /// <summary>
    /// Insert an audio file onto a slide. Supports .mp3, .wav, .m4a, .wma.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="filePath">Full path to the audio file</param>
    /// <param name="left">Position from left in points</param>
    /// <param name="top">Position from top in points</param>
    /// <param name="linkToFile">If true, link to file instead of embedding (smaller file size)</param>
    /// <param name="saveWithDocument">If true, save media with document when linking</param>
    [ServiceAction("insert-audio")]
    OperationResult InsertAudio(IPptBatch batch, int slideIndex, string filePath, float left, float top, bool linkToFile, bool saveWithDocument);

    /// <summary>
    /// Insert a video file onto a slide. Supports .mp4, .avi, .mov, .wmv.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="slideIndex">1-based slide index</param>
    /// <param name="filePath">Full path to the video file</param>
    /// <param name="left">Position from left in points</param>
    /// <param name="top">Position from top in points</param>
    /// <param name="width">Width in points (0 = use video native width)</param>
    /// <param name="height">Height in points (0 = use video native height)</param>
    /// <param name="linkToFile">If true, link to file instead of embedding</param>
    [ServiceAction("insert-video")]
    OperationResult InsertVideo(IPptBatch batch, int slideIndex, string filePath, float left, float top, float width, float height, bool linkToFile);

    /// <summary>
    /// Get information about a media shape (audio or video) on a slide.
    /// </summary>
    [ServiceAction("get-info")]
    MediaInfoResult GetInfo(IPptBatch batch, int slideIndex, string shapeName);
}
