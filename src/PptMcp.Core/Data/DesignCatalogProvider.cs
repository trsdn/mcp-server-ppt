using System.Collections.Concurrent;
using System.Reflection;
using System.Text.Json;

namespace PptMcp.Core.Data;

/// <summary>
/// Provides access to embedded design catalog data (archetypes, palettes, grids, etc.).
/// Data is loaded once and cached for the lifetime of the application.
/// </summary>
internal static partial class DesignCatalogProvider
{
    private static readonly Assembly Assembly = typeof(DesignCatalogProvider).Assembly;
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    // Cached data
    private static List<ArchetypeEntry>? _archetypes;
    private static List<PaletteEntry>? _palettes;
    private static List<StyleProfileEntry>? _styleProfiles;
    private static LayoutGridData? _layoutGrids;
    private static List<DensityProfileEntry>? _densityProfiles;
    private static ContextModelData? _contextModel;
    private static List<DeckSequenceEntry>? _deckSequences;
    private static string? _archetypeRegistry;
    private static readonly ConcurrentDictionary<string, string?> _archetypeDetails = new(StringComparer.OrdinalIgnoreCase);
    private static string? _slidePatternsDetail;
    private static string? _iconShapesDetail;

    // ── Archetypes ─────────────────────────────────────

    public static List<ArchetypeEntry> GetArchetypes()
    {
        _archetypes ??= LoadJson<List<ArchetypeEntry>>("archetypes.json");
        return _archetypes;
    }

    public static string GetArchetypeRegistry()
    {
        _archetypeRegistry ??= LoadText("archetypes.registry.md");
        return _archetypeRegistry;
    }

    public static string? GetArchetypeDetail(string archetypeId)
    {
        return _archetypeDetails.GetOrAdd(archetypeId, id =>
            TryLoadText($"archetypes.{id}.md"));
    }

    // ── Color Palettes ─────────────────────────────────

    public static List<PaletteEntry> GetPalettes()
    {
        _palettes ??= LoadJson<List<PaletteEntry>>("color-palettes.json");
        return _palettes;
    }

    public static PaletteEntry? GetPalette(string paletteId)
    {
        return GetPalettes().Find(p => string.Equals(p.Id, paletteId, StringComparison.OrdinalIgnoreCase));
    }

    // ── Style Profiles ─────────────────────────────────

    public static List<StyleProfileEntry> GetStyleProfiles()
    {
        _styleProfiles ??= LoadJson<List<StyleProfileEntry>>("style-profiles.json");
        return _styleProfiles;
    }

    public static StyleProfileEntry? GetStyleProfile(string profileId)
    {
        return GetStyleProfiles().Find(p => string.Equals(p.Id, profileId, StringComparison.OrdinalIgnoreCase));
    }

    // ── Layout Grids ───────────────────────────────────

    public static LayoutGridData GetLayoutGridData()
    {
        _layoutGrids ??= LoadJson<LayoutGridData>("layout-grids.json");
        return _layoutGrids;
    }

    public static LayoutGridEntry? GetLayoutGrid(string gridId)
    {
        return GetLayoutGridData().Grids.Find(g => string.Equals(g.Id, gridId, StringComparison.OrdinalIgnoreCase));
    }

    // ── Density Profiles ───────────────────────────────

    public static List<DensityProfileEntry> GetDensityProfiles()
    {
        _densityProfiles ??= LoadJson<List<DensityProfileEntry>>("density-profiles.json");
        return _densityProfiles;
    }

    public static DensityProfileEntry? GetDensityProfile(string profileId)
    {
        return GetDensityProfiles().Find(p => string.Equals(p.Id, profileId, StringComparison.OrdinalIgnoreCase));
    }

    // ── Context Model ──────────────────────────────────

    public static ContextModelData GetContextModel()
    {
        _contextModel ??= LoadJson<ContextModelData>("context-model.json");
        return _contextModel;
    }

    // ── Deck Sequences ─────────────────────────────────

    public static List<DeckSequenceEntry> GetDeckSequences()
    {
        _deckSequences ??= LoadJson<List<DeckSequenceEntry>>("deck-sequences.json");
        return _deckSequences;
    }

    public static DeckSequenceEntry? GetDeckSequence(string sequenceId)
    {
        return GetDeckSequences().Find(s => string.Equals(s.Id, sequenceId, StringComparison.OrdinalIgnoreCase));
    }

    // ── Slide Patterns ─────────────────────────────────

    public static string GetSlidePatterns()
    {
        _slidePatternsDetail ??= LoadText("slide-patterns-detail.md");
        return _slidePatternsDetail;
    }

    // ── Icon Shapes ────────────────────────────────────

    public static string GetIconShapes()
    {
        _iconShapesDetail ??= LoadText("icon-shapes-detail.md");
        return _iconShapesDetail;
    }

    // ── Internal helpers ───────────────────────────────

    private static T LoadJson<T>(string fileName) where T : class
    {
        var resourceName = GetResourceName(fileName);
        using var stream = Assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException($"Embedded resource not found: {resourceName}");
        return JsonSerializer.Deserialize<T>(stream, JsonOptions)
            ?? throw new InvalidOperationException($"Failed to deserialize: {resourceName}");
    }

    private static string LoadText(string fileName)
    {
        var resourceName = GetResourceName(fileName);
        using var stream = Assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException($"Embedded resource not found: {resourceName}");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static string? TryLoadText(string fileName)
    {
        var resourceName = GetResourceName(fileName);
        using var stream = Assembly.GetManifestResourceStream(resourceName);
        if (stream is null)
            return null;
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static string GetResourceName(string fileName)
    {
        return $"PptMcp.Core.Data.{fileName.Replace('\\', '.').Replace('/', '.')}";
    }
}

// ── Data models ────────────────────────────────────────

internal sealed class ArchetypeEntry
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Section { get; set; } = "";
    public string When { get; set; } = "";
    public List<string> BestDensity { get; set; } = [];
    public List<string> Variants { get; set; } = [];
    public string ExampleTitle { get; set; } = "";
}

internal sealed class PaletteEntry
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string BestFor { get; set; } = "";
    public Dictionary<string, string> Colors { get; set; } = [];
}

internal sealed class StyleProfileEntry
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Description { get; set; } = "";
    public string BestFor { get; set; } = "";
    public string ColorScheme { get; set; } = "";
    public string Font { get; set; } = "";
    public string TitleStyle { get; set; } = "";
    public int TitleSize { get; set; }
    public int BodySize { get; set; }
    public int FootnoteSize { get; set; }
    public string BulletsPerSlide { get; set; } = "";
    public string WordsPerBullet { get; set; } = "";
    public string ContentDensity { get; set; } = "";
    public List<string> PreferredArchetypes { get; set; } = [];
    public string Whitespace { get; set; } = "";
    public string Background { get; set; } = "";
    public string ChartStyle { get; set; } = "";
    public string SpecialRules { get; set; } = "";
}

internal sealed class LayoutGridData
{
    public SafeZoneData SafeZones { get; set; } = new();
    public List<LayoutGridEntry> Grids { get; set; } = [];
}

internal sealed class SafeZoneData
{
    public int SlideWidth { get; set; }
    public int SlideHeight { get; set; }
    public int Margin { get; set; }
    public ZoneRect Title { get; set; } = new();
    public ZoneRect ContentArea { get; set; } = new();
    public ZoneRect SourceBar { get; set; } = new();
    public ZoneRect FooterLeft { get; set; } = new();
    public ZoneRect FooterRight { get; set; } = new();
}

internal sealed class ZoneRect
{
    public int X { get; set; }
    public int Y { get; set; }
    public int W { get; set; }
    public int H { get; set; }
}

internal sealed class LayoutGridEntry
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string BestFor { get; set; } = "";
    public int? Gap { get; set; }
    public int? Gutter { get; set; }
    public int? GapH { get; set; }
    public int? GapV { get; set; }
    public List<GridZone> Zones { get; set; } = [];
}

internal sealed class GridZone
{
    public string Name { get; set; } = "";
    public int? X { get; set; }
    public int? Y { get; set; }
    public int? W { get; set; }
    public int? H { get; set; }
    public string? Fill { get; set; }
    public string? Description { get; set; }
}

internal sealed class DensityProfileEntry
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string UsedFor { get; set; } = "";
    public string Audience { get; set; } = "";
    public string Mode { get; set; } = "";
    public string TextVolume { get; set; } = "";
    public string ElementCount { get; set; } = "";
    public string DataGranularity { get; set; } = "";
    public string AnnotationDepth { get; set; } = "";
    public string SourceCompleteness { get; set; } = "";
    public string WhiteSpaceRatio { get; set; } = "";
    public string Character { get; set; } = "";
    public List<string> BestArchetypes { get; set; } = [];
}

internal sealed class ContextModelData
{
    public List<MeetingTypeEntry> MeetingTypes { get; set; } = [];
    public List<AudienceLevelEntry> AudienceLevels { get; set; } = [];
    public List<ConsumptionModeEntry> ConsumptionModes { get; set; } = [];
    public DensitySelectionData DensitySelection { get; set; } = new();
}

internal sealed class MeetingTypeEntry
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Audience { get; set; } = "";
    public string TimePerSlide { get; set; } = "";
    public string Goal { get; set; } = "";
    public string DecisionPressure { get; set; } = "";
    public string PrimaryMode { get; set; } = "";
    public string? SecondaryMode { get; set; }
    public string DefaultDensity { get; set; } = "";
}

internal sealed class AudienceLevelEntry
{
    public string Id { get; set; } = "";
    public string Label { get; set; } = "";
    public string Roles { get; set; } = "";
    public string AttentionSpan { get; set; } = "";
    public string PreferredDensity { get; set; } = "";
    public string WantsToSee { get; set; } = "";
}

internal sealed class ConsumptionModeEntry
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public bool SpeakerPresent { get; set; }
    public bool SelfContained { get; set; }
    public string TextDensity { get; set; } = "";
    public string AnnotationDepth { get; set; } = "";
}

internal sealed class DensitySelectionData
{
    public string Description { get; set; } = "";
    public List<DensityRule> Rules { get; set; } = [];
    public string Default { get; set; } = "";
}

internal sealed class DensityRule
{
    public string Mode { get; set; } = "";
    public List<string> Audience { get; set; } = [];
    public string Density { get; set; } = "";
}

internal sealed class DeckSequenceEntry
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string UsedFor { get; set; } = "";
    public string Intent { get; set; } = "";
    public List<DeckSlideEntry> Slides { get; set; } = [];
}

internal sealed class DeckSlideEntry
{
    public object Position { get; set; } = "";
    public string Purpose { get; set; } = "";
    public string Archetype { get; set; } = "";
    public string Density { get; set; } = "";
}
