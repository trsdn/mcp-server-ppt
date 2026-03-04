using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Design;

/// <summary>
/// Theme and design operations: list designs, apply themes, get theme colors.
/// </summary>
[ServiceCategory("design")]
[McpTool("design", Title = "Design Operations", Destructive = true, Category = "design")]
public interface IDesignCommands
{
    /// <summary>
    /// List all designs (themes) in the presentation.
    /// </summary>
    [ServiceAction("list")]
    DesignListResult List(IPptBatch batch);

    /// <summary>
    /// Apply an Office theme file (.thmx) to the presentation.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="themePath">Full path to .thmx theme file</param>
    [ServiceAction("apply-theme")]
    OperationResult ApplyTheme(IPptBatch batch, string themePath);

    /// <summary>
    /// Get the theme color palette for a design.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="designIndex">1-based design index (0 = first design)</param>
    [ServiceAction("get-colors")]
    ThemeColorResult GetColors(IPptBatch batch, int designIndex);

    /// <summary>
    /// List all color schemes in the presentation.
    /// </summary>
    [ServiceAction("list-color-schemes")]
    ColorSchemeListResult ListColorSchemes(IPptBatch batch);

    /// <summary>
    /// Get the theme fonts (heading and body) for a design.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="designIndex">1-based design index (0 = first design)</param>
    [ServiceAction("get-fonts")]
    ThemeFontResult GetThemeFonts(IPptBatch batch, int designIndex);
}
