using PptMcp.ComInterop.Session;
using PptMcp.Core.Attributes;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Design;

/// <summary>
/// Design operations: themes, colors, fonts, and design knowledge catalog.
///
/// THEME OPERATIONS (require open presentation):
/// - list, apply-theme, get-colors, list-color-schemes, get-fonts
///
/// DESIGN KNOWLEDGE CATALOG (query on demand):
/// - list-archetypes / get-archetype: primary unified archetype surface with curated rules plus learned subtype/example coverage when local reference data is available
/// - list-palettes / get-palette: 8 color palettes with hex values
/// - list-style-profiles / get-style-profile: consulting/sales/startup configurations
/// - list-layout-grids / get-layout-grid: exact x/y/w/h positioning coordinates
/// - list-density-profiles / get-density-profile: D1-D5 content density rules
/// - get-context-model: meeting type → audience → density mapping
/// - get-deck-sequence: canonical slide sequences (decision, pitch, review)
/// - get-slide-patterns: advanced consulting layout patterns
/// - get-icon-shapes: native shape icon catalog
/// </summary>
[ServiceCategory("design")]
[McpTool("design", Title = "Design Operations", Destructive = true, Category = "design",
    Description = "Manage presentation themes and design knowledge. "
    + "THEME OPS: 'list' designs, 'apply-theme' .thmx files, 'get-colors'/'get-fonts' from theme. "
    + "DESIGN CATALOG (query on demand instead of reading full docs): "
    + "'list-archetypes'/'get-archetype' as the primary unified archetype surface with curated layout guidance plus learned subtype/example coverage when local reference data is available. "
    + "'list-palettes'/'get-palette' for 8 color palettes with hex values. "
    + "'list-style-profiles'/'get-style-profile' for style configurations. "
    + "'list-layout-grids'/'get-layout-grid' for exact positioning coordinates. "
    + "'list-density-profiles'/'get-density-profile' for D1-D5 content density rules. "
    + "'get-context-model' for meeting type → density mapping. "
    + "'get-deck-sequence' for canonical slide sequences. "
    + "'get-slide-patterns' for advanced consulting patterns. "
    + "'get-icon-shapes' for native shape icon catalog. "
    + "design_index: 1-based (0 = first/active design).")]
public interface IDesignCommands
{
    // ── Theme Operations (require open presentation) ──────

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

    // ── Design Knowledge Catalog (static reference data) ──

    /// <summary>
    /// List the unified archetype catalog.
    /// Returns curated archetypes plus learned-only families when the local sanitized reference catalog is available.
    /// </summary>
    [ServiceAction("list-archetypes")]
    ArchetypeListResult ListArchetypes(IPptBatch batch);

    /// <summary>
    /// Get full detail for one archetype.
    /// Returns curated layout guidance when available and also attaches learned subtypes plus concrete sanitized observed examples from the local reference catalog when present.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="archetypeId">Archetype id: big-number, kpi-card-dashboard, operational-kpi, column-bar-chart, chart-insight-callout, framework, simple-table, waterfall-chart, comparison, timeline-roadmap, process-diagram, executive-summary, recommendations, quote, map, appendix, annotated-diagram, chart-and-commentary, org-chart, title-slide</param>
    [ServiceAction("get-archetype")]
    ArchetypeDetailResult GetArchetype(IPptBatch batch, string archetypeId);

    /// <summary>
    /// List all 8 curated color palettes.
    /// </summary>
    [ServiceAction("list-palettes")]
    PaletteListResult ListPalettes(IPptBatch batch);

    /// <summary>
    /// Get hex values for all color roles in a palette.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="paletteId">Palette id: corporate-blue, slate-professional, modern-tech, confident-bold, warm-coral, academic-neutral, nature-calm, electric-minimal</param>
    [ServiceAction("get-palette")]
    PaletteDetailResult GetPalette(IPptBatch batch, string paletteId);

    /// <summary>
    /// List all 8 style profiles.
    /// </summary>
    [ServiceAction("list-style-profiles")]
    StyleProfileListResult ListStyleProfiles(IPptBatch batch);

    /// <summary>
    /// Get full style profile configuration: fonts, sizes, spacing, preferred archetypes.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="profileId">Profile id: consulting, corporate, sales, startup, keynote, educational, academic, creative</param>
    [ServiceAction("get-style-profile")]
    StyleProfileDetailResult GetStyleProfile(IPptBatch batch, string profileId);

    /// <summary>
    /// List all layout grids with names and descriptions.
    /// </summary>
    [ServiceAction("list-layout-grids")]
    LayoutGridListResult ListLayoutGrids(IPptBatch batch);

    /// <summary>
    /// Get exact positioning coordinates (x, y, w, h in points) for a layout grid.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="gridId">Grid id: single-column, two-column-equal, two-column-60-40, three-column, four-column, grid-2x2, grid-3x2, big-number, kpi-3-cards, kpi-4-cards, kpi-5-cards, chart-full, chart-callout, timeline, process-3-step, section-divider, title-dark-hero</param>
    [ServiceAction("get-layout-grid")]
    LayoutGridResult GetLayoutGrid(IPptBatch batch, string gridId);

    /// <summary>
    /// List all 5 density profiles (D1 Minimal to D5 Dense).
    /// </summary>
    [ServiceAction("list-density-profiles")]
    DensityProfileListResult ListDensityProfiles(IPptBatch batch);

    /// <summary>
    /// Get full density profile: word counts, element counts, whitespace ratio, best archetypes.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="densityId">Density id: D1, D2, D3, D4, D5</param>
    [ServiceAction("get-density-profile")]
    DensityProfileResult GetDensityProfile(IPptBatch batch, string densityId);

    /// <summary>
    /// Get the full context model: meeting types → audience levels → density mapping.
    /// </summary>
    [ServiceAction("get-context-model")]
    ContextModelResult GetContextModel(IPptBatch batch);

    /// <summary>
    /// Get a canonical deck sequence with slide-by-slide archetype and density recommendations.
    /// </summary>
    /// <param name="batch">Batch context</param>
    /// <param name="sequenceId">Sequence id: S1 (Decision), S2 (Performance Review), S3 (Pitch/Proposal), S4 (Transformation), S5 (Regulatory)</param>
    [ServiceAction("get-deck-sequence")]
    DeckSequenceDetailResult GetDeckSequence(IPptBatch batch, string sequenceId);

    /// <summary>
    /// Get advanced consulting-grade slide layout patterns with shape density benchmarks.
    /// </summary>
    [ServiceAction("get-slide-patterns")]
    SlidePatternListResult GetSlidePatterns(IPptBatch batch);

    /// <summary>
    /// Get the icon shape catalog: native PowerPoint shapes for building icons, arrows, connectors.
    /// </summary>
    [ServiceAction("get-icon-shapes")]
    IconShapeListResult GetIconShapes(IPptBatch batch);
}
