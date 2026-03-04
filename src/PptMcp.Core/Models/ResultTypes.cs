using System.Text.Json.Serialization;

namespace PptMcp.Core.Models;

/// <summary>
/// Base result type for all Core operations.
/// Exceptions propagate naturally — batch.Execute() re-throws them via TaskCompletionSource.
/// </summary>
public abstract class ResultBase
{
    public bool Success { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ErrorMessage { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FilePath { get; set; }
}

/// <summary>
/// Result for operations that don't return data (create, delete, etc.)
/// </summary>
public class OperationResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Action { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }
}

/// <summary>
/// Result for rename operations
/// </summary>
public class RenameResult : ResultBase
{
    public string ObjectType { get; set; } = string.Empty;
    public string OldName { get; set; } = string.Empty;
    public string NewName { get; set; } = string.Empty;
}

// ── File / Session ────────────────────────────────────────

public class FileValidationInfo : ResultBase
{
    public bool Exists { get; set; }
    public string FileName { get; set; } = string.Empty;
    public long FileSizeBytes { get; set; }
    public bool IsReadOnly { get; set; }
    public bool IsMacroEnabled { get; set; }
    public int SlideCount { get; set; }
}

// ── Slide ─────────────────────────────────────────────────

public class SlideListResult : ResultBase
{
    public List<SlideInfo> Slides { get; set; } = [];
}

public class SlideInfo
{
    public int SlideIndex { get; set; }
    public int SlideNumber { get; set; }
    public string SlideId { get; set; } = string.Empty;
    public string LayoutName { get; set; } = string.Empty;
    public string MasterName { get; set; } = string.Empty;
    public int ShapeCount { get; set; }
    public bool HasNotes { get; set; }
    public bool HasAnimations { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; set; }
}

public class SlideDetailResult : ResultBase
{
    public SlideInfo? Slide { get; set; }
    public List<ShapeInfo> Shapes { get; set; } = [];
}

// ── Shape ─────────────────────────────────────────────────

public class ShapeListResult : ResultBase
{
    public int SlideIndex { get; set; }
    public List<ShapeInfo> Shapes { get; set; } = [];
}

public class ShapeInfo
{
    public int ShapeId { get; set; }
    public string Name { get; set; } = string.Empty;
    public string ShapeType { get; set; } = string.Empty;
    public float Left { get; set; }
    public float Top { get; set; }
    public float Width { get; set; }
    public float Height { get; set; }
    public int ZOrderPosition { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Text { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AlternativeText { get; set; }

    public bool HasTextFrame { get; set; }
    public bool HasTable { get; set; }
    public bool HasChart { get; set; }
    public bool IsGroup { get; set; }
    public bool IsPlaceholder { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? PlaceholderType { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<ShapeInfo>? GroupItems { get; set; }
}

public class ShapeDetailResult : ResultBase
{
    public ShapeInfo? Shape { get; set; }
}

// ── Text ──────────────────────────────────────────────────

public class TextResult : ResultBase
{
    public int ShapeId { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public List<TextParagraphInfo> Paragraphs { get; set; } = [];
}

public class TextParagraphInfo
{
    public int Index { get; set; }
    public string Text { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Alignment { get; set; }

    public List<TextRunInfo> Runs { get; set; } = [];
}

public class TextRunInfo
{
    public string Text { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FontName { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public float? FontSize { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Bold { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Italic { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Color { get; set; }
}

// ── Table (in shapes) ────────────────────────────────────

public class SlideTableResult : ResultBase
{
    public int ShapeId { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
    public List<List<string?>> Data { get; set; } = [];
}

// ── Master / Layout ───────────────────────────────────────

public class MasterListResult : ResultBase
{
    public List<MasterInfo> Masters { get; set; } = [];
}

public class MasterInfo
{
    public string Name { get; set; } = string.Empty;
    public List<LayoutInfo> Layouts { get; set; } = [];
}

public class LayoutInfo
{
    public string Name { get; set; } = string.Empty;
    public int Index { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? MatchingName { get; set; }
}

// ── Notes ─────────────────────────────────────────────────

public class NotesResult : ResultBase
{
    public int SlideIndex { get; set; }
    public string Text { get; set; } = string.Empty;
}

// ── Transition ────────────────────────────────────────────

public class TransitionResult : ResultBase
{
    public int SlideIndex { get; set; }
    public string TransitionType { get; set; } = string.Empty;
    public float Duration { get; set; }
    public bool AdvanceOnClick { get; set; }
    public float AdvanceAfterTime { get; set; }
}

// ── Animation ─────────────────────────────────────────────

public class AnimationListResult : ResultBase
{
    public int SlideIndex { get; set; }
    public List<AnimationInfo> Animations { get; set; } = [];
}

public class AnimationInfo
{
    public int Index { get; set; }
    public int ShapeId { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public string EffectType { get; set; } = string.Empty;
    public string Timing { get; set; } = string.Empty;
    public float Duration { get; set; }
    public float Delay { get; set; }
}

// ── Export ─────────────────────────────────────────────────

public class ExportResult : ResultBase
{
    public string OutputPath { get; set; } = string.Empty;
    public string Format { get; set; } = string.Empty;
}

// ── Chart ──────────────────────────────────────────────────

public class ChartInfoResult : ResultBase
{
    public int ShapeId { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public int ChartType { get; set; }
    public string ChartTypeName { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Title { get; set; }

    public bool HasLegend { get; set; }
    public int SeriesCount { get; set; }
}

// ── Design / Theme ────────────────────────────────────────

public class DesignListResult : ResultBase
{
    public List<DesignInfo> Designs { get; set; } = [];
}

public class DesignInfo
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public int LayoutCount { get; set; }
}

public class ThemeColorResult : ResultBase
{
    public string DesignName { get; set; } = string.Empty;
    public Dictionary<string, string> Colors { get; set; } = [];
}

// ── Theme Fonts ──────────────────────────────────────────

public class ThemeFontResult : ResultBase
{
    public string DesignName { get; set; } = string.Empty;
    public string HeadingFont { get; set; } = string.Empty;
    public string BodyFont { get; set; } = string.Empty;
}

// ── Slideshow ─────────────────────────────────────────────

public class SlideshowInfoResult : ResultBase
{
    public bool IsRunning { get; set; }
    public int CurrentSlide { get; set; }
    public int TotalSlides { get; set; }
}

// ── VBA ───────────────────────────────────────────────────

public class VbaModuleListResult : ResultBase
{
    public List<VbaModuleInfo> Modules { get; set; } = [];
}

public class VbaModuleInfo
{
    public string Name { get; set; } = string.Empty;
    public int ModuleType { get; set; }
    public string ModuleTypeName { get; set; } = string.Empty;
    public int LineCount { get; set; }
}

public class VbaModuleCodeResult : ResultBase
{
    public string ModuleName { get; set; } = string.Empty;
    public string Code { get; set; } = string.Empty;
    public int LineCount { get; set; }
}

// ── Window ────────────────────────────────────────────────

public class WindowInfoResult : ResultBase
{
    public int WindowState { get; set; }
    public string WindowStateName { get; set; } = string.Empty;
    public float Left { get; set; }
    public float Top { get; set; }
    public float Width { get; set; }
    public float Height { get; set; }
}

// ── Hyperlink ─────────────────────────────────────────────

public class HyperlinkResult : ResultBase
{
    public int SlideIndex { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public bool HasHyperlink { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Address { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SubAddress { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ScreenTip { get; set; }
}

public class HyperlinkListResult : ResultBase
{
    public List<HyperlinkInfo> Hyperlinks { get; set; } = [];
}

public class HyperlinkInfo
{
    public int Index { get; set; }
    public string Address { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SubAddress { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ScreenTip { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int SlideIndex { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ShapeName { get; set; }
}

// ── Section ───────────────────────────────────────────────

public class SectionListResult : ResultBase
{
    public List<SectionInfo> Sections { get; set; } = [];
}

public class SectionInfo
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public int FirstSlideIndex { get; set; }
    public int SlideCount { get; set; }
}

// ── Document Properties ───────────────────────────────────

public class DocumentPropertyResult : ResultBase
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Title { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Author { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Keywords { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Comments { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Company { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Category { get; set; }
}

// ── Media ─────────────────────────────────────────────────

public class MediaInfoResult : ResultBase
{
    public int SlideIndex { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public string MediaType { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SourceFile { get; set; }

    public float Left { get; set; }
    public float Top { get; set; }
    public float Width { get; set; }
    public float Height { get; set; }
}

// ── Comment ──────────────────────────────────────────────

public class CommentListResult : ResultBase
{
    public List<CommentInfo> Comments { get; set; } = [];
}

public class CommentInfo
{
    public int SlideIndex { get; set; }
    public int CommentIndex { get; set; }
    public string Author { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public float Left { get; set; }
    public float Top { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DateTime { get; set; }
}

// ── Placeholder ──────────────────────────────────────────

public class PlaceholderListResult : ResultBase
{
    public int SlideIndex { get; set; }
    public List<PlaceholderInfo> Placeholders { get; set; } = [];
}

public class PlaceholderInfo
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public int PlaceholderType { get; set; }
    public string PlaceholderTypeName { get; set; } = string.Empty;
    public bool HasTextFrame { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Text { get; set; }
}

// ── Background ───────────────────────────────────────────

public class BackgroundResult : ResultBase
{
    public int SlideIndex { get; set; }
    public bool FollowMasterBackground { get; set; }
    public string FillType { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Color { get; set; }
}

// ── Header/Footer ────────────────────────────────────────

public class HeaderFooterResult : ResultBase
{
    public bool ShowFooter { get; set; }
    public bool ShowSlideNumber { get; set; }
    public bool ShowDate { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FooterText { get; set; }
}

// ── SmartArt ─────────────────────────────────────────────

public class SmartArtInfoResult : ResultBase
{
    public int SlideIndex { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public string LayoutName { get; set; } = string.Empty;
    public List<SmartArtNodeInfo> Nodes { get; set; } = [];
}

public class SmartArtNodeInfo
{
    public int Index { get; set; }
    public string Text { get; set; } = string.Empty;
    public int Level { get; set; }
}

// ── Custom Show ──────────────────────────────────────────

public class CustomShowListResult : ResultBase
{
    public List<CustomShowInfo> Shows { get; set; } = [];
}

public class CustomShowInfo
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public int SlideCount { get; set; }
    public List<int> SlideIds { get; set; } = [];
}

// ── Page Setup ───────────────────────────────────────────

public class PageSetupResult : ResultBase
{
    public float SlideWidth { get; set; }
    public float SlideHeight { get; set; }
    public int SlideOrientation { get; set; }
    public int NotesOrientation { get; set; }
}

// ── Tags ─────────────────────────────────────────────────

public class TagListResult : ResultBase
{
    public int SlideIndex { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ShapeName { get; set; }

    public List<TagInfo> Tags { get; set; } = [];
}

public class TagInfo
{
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}

// ── Color Scheme ─────────────────────────────────────────

public class ColorSchemeListResult : ResultBase
{
    public List<ColorSchemeInfo> ColorSchemes { get; set; } = [];
}

public class ColorSchemeInfo
{
    public int Index { get; set; }
    public Dictionary<string, string> Colors { get; set; } = [];
}

// ── Accessibility ────────────────────────────────────────

public class AccessibilityAuditResult : OperationResult
{
    public int TotalSlides { get; set; }
    public int IssueCount { get; set; }
    public List<AccessibilityIssue> Issues { get; set; } = [];
}

public class AccessibilityIssue
{
    public int SlideIndex { get; set; }
    public string IssueType { get; set; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ShapeName { get; set; }

    public string Description { get; set; } = string.Empty;
}

public class ReadingOrderResult : ResultBase
{
    public int SlideIndex { get; set; }
    public List<ReadingOrderEntry> Shapes { get; set; } = [];
}

public class ReadingOrderEntry
{
    public int Position { get; set; }
    public string ShapeName { get; set; } = string.Empty;
    public string ShapeType { get; set; } = string.Empty;
    public int ZOrderPosition { get; set; }
}
