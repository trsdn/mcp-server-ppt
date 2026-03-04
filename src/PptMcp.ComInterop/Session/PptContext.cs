using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptMcp.ComInterop.Session;

/// <summary>
/// Provides access to PowerPoint COM objects for operations.
/// Simplifies passing PowerPoint application and presentation to operations.
/// </summary>
public sealed class PptContext
{
    /// <summary>
    /// Creates a new PptContext.
    /// </summary>
    /// <param name="presentationPath">Full path to the presentation</param>
    /// <param name="app">PowerPoint.Application COM object</param>
    /// <param name="presentation">PowerPoint.Presentation COM object</param>
    public PptContext(string presentationPath, PowerPoint.Application app, PowerPoint.Presentation presentation)
    {
        PresentationPath = presentationPath ?? throw new ArgumentNullException(nameof(presentationPath));
        App = app ?? throw new ArgumentNullException(nameof(app));
        Presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
    }

    /// <summary>
    /// Gets the full path to the presentation.
    /// </summary>
    public string PresentationPath { get; }

    /// <summary>
    /// Gets the PowerPoint.Application COM object.
    /// </summary>
    public PowerPoint.Application App { get; }

    /// <summary>
    /// Gets the PowerPoint.Presentation COM object.
    /// </summary>
    public PowerPoint.Presentation Presentation { get; }
}


