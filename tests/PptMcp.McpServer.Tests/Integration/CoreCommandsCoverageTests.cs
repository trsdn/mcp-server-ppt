// Suppress IDE0005 (unnecessary using) – explicit usings kept for clarity in test reflection code
#pragma warning disable IDE0005
using System.Reflection;
using PptMcp.Core.Commands.Accessibility;
using PptMcp.Core.Commands.Animation;
using PptMcp.Core.Commands.Background;
using PptMcp.Core.Commands.Chart;
using PptMcp.Core.Commands.Comment;
using PptMcp.Core.Commands.CustomShow;
using PptMcp.Core.Commands.Design;
using PptMcp.Core.Commands.DocumentProperty;
using PptMcp.Core.Commands.Export;
using PptMcp.Core.Commands.File;
using PptMcp.Core.Commands.HeaderFooter;
using PptMcp.Core.Commands.Hyperlink;
using PptMcp.Core.Commands.Image;
using PptMcp.Core.Commands.Master;
using PptMcp.Core.Commands.Media;
using PptMcp.Core.Commands.Notes;
using PptMcp.Core.Commands.PageSetup;
using PptMcp.Core.Commands.Placeholder;
using PptMcp.Core.Commands.Proofing;
using PptMcp.Core.Commands.Section;
using PptMcp.Core.Commands.Shape;
using PptMcp.Core.Commands.ShapeAlign;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Commands.SlideImport;
using PptMcp.Core.Commands.Slideshow;
using PptMcp.Core.Commands.SlideTable;
using PptMcp.Core.Commands.SmartArt;
using PptMcp.Core.Commands.Tag;
using PptMcp.Core.Commands.Text;
using PptMcp.Core.Commands.Transition;
using PptMcp.Core.Commands.Vba;
using PptMcp.Core.Commands.Window;
#pragma warning restore IDE0005
using PptMcp.Generated;
using Xunit;

namespace PptMcp.McpServer.Tests.Integration;

/// <summary>
/// CRITICAL: Automated verification that all Core Commands methods are exposed via generated actions.
/// These tests PREVENT regression by ensuring compile-time and runtime coverage.
/// </summary>
public class CoreCommandsCoverageTests
{
    // ── Existing coverage tests ──────────────────────────────

    [Fact]
    public void ISlideCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(ISlideCommands));
        var enumValueCount = Enum.GetValues<SlideAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"ISlideCommands has {coreMethodCount} [ServiceAction] methods but SlideAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IShapeCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IShapeCommands));
        var enumValueCount = Enum.GetValues<ShapeAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IShapeCommands has {coreMethodCount} [ServiceAction] methods but ShapeAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void ITextCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(ITextCommands));
        var enumValueCount = Enum.GetValues<TextAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"ITextCommands has {coreMethodCount} [ServiceAction] methods but TextAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void INotesCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(INotesCommands));
        var enumValueCount = Enum.GetValues<NotesAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"INotesCommands has {coreMethodCount} [ServiceAction] methods but NotesAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IMasterCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IMasterCommands));
        var enumValueCount = Enum.GetValues<MasterAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IMasterCommands has {coreMethodCount} [ServiceAction] methods but MasterAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IExportCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IExportCommands));
        var enumValueCount = Enum.GetValues<ExportAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IExportCommands has {coreMethodCount} [ServiceAction] methods but ExportAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void ITransitionCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(ITransitionCommands));
        var enumValueCount = Enum.GetValues<TransitionAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"ITransitionCommands has {coreMethodCount} [ServiceAction] methods but TransitionAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IImageCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IImageCommands));
        var enumValueCount = Enum.GetValues<ImageAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IImageCommands has {coreMethodCount} [ServiceAction] methods but ImageAction has only {enumValueCount} enum values.");
    }

    // ── NEW: Coverage tests for previously untested command areas ──

    [Fact]
    public void IAnimationCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IAnimationCommands));
        var enumValueCount = Enum.GetValues<AnimationAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IAnimationCommands has {coreMethodCount} [ServiceAction] methods but AnimationAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IChartCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IChartCommands));
        var enumValueCount = Enum.GetValues<ChartAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IChartCommands has {coreMethodCount} [ServiceAction] methods but ChartAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IDesignCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IDesignCommands));
        var enumValueCount = Enum.GetValues<DesignAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IDesignCommands has {coreMethodCount} [ServiceAction] methods but DesignAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IDocumentPropertyCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IDocumentPropertyCommands));
        var enumValueCount = Enum.GetValues<DocpropertyAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IDocumentPropertyCommands has {coreMethodCount} [ServiceAction] methods but DocpropertyAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IHyperlinkCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IHyperlinkCommands));
        var enumValueCount = Enum.GetValues<HyperlinkAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IHyperlinkCommands has {coreMethodCount} [ServiceAction] methods but HyperlinkAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IMediaCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IMediaCommands));
        var enumValueCount = Enum.GetValues<MediaAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IMediaCommands has {coreMethodCount} [ServiceAction] methods but MediaAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void ISectionCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(ISectionCommands));
        var enumValueCount = Enum.GetValues<SectionAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"ISectionCommands has {coreMethodCount} [ServiceAction] methods but SectionAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void ISlideshowCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(ISlideshowCommands));
        var enumValueCount = Enum.GetValues<SlideshowAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"ISlideshowCommands has {coreMethodCount} [ServiceAction] methods but SlideshowAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void ISlideTableCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(ISlideTableCommands));
        var enumValueCount = Enum.GetValues<SlidetableAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"ISlideTableCommands has {coreMethodCount} [ServiceAction] methods but SlidetableAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IVbaCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IVbaCommands));
        var enumValueCount = Enum.GetValues<VbaAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IVbaCommands has {coreMethodCount} [ServiceAction] methods but VbaAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IWindowCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IWindowCommands));
        var enumValueCount = Enum.GetValues<WindowAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IWindowCommands has {coreMethodCount} [ServiceAction] methods but WindowAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IFileCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IFileCommands));
        var enumValueCount = Enum.GetValues<FileAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IFileCommands has {coreMethodCount} [ServiceAction] methods but FileAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void ICommentCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(ICommentCommands));
        var enumValueCount = Enum.GetValues<CommentAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"ICommentCommands has {coreMethodCount} [ServiceAction] methods but CommentAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IPlaceholderCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IPlaceholderCommands));
        var enumValueCount = Enum.GetValues<PlaceholderAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IPlaceholderCommands has {coreMethodCount} [ServiceAction] methods but PlaceholderAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IBackgroundCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IBackgroundCommands));
        var enumValueCount = Enum.GetValues<BackgroundAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IBackgroundCommands has {coreMethodCount} [ServiceAction] methods but BackgroundAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IHeaderFooterCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IHeaderFooterCommands));
        var enumValueCount = Enum.GetValues<HeaderfooterAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IHeaderFooterCommands has {coreMethodCount} [ServiceAction] methods but HeaderfooterAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void ISmartArtCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(ISmartArtCommands));
        var enumValueCount = Enum.GetValues<SmartartAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"ISmartArtCommands has {coreMethodCount} [ServiceAction] methods but SmartartAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IShapeAlignCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IShapeAlignCommands));
        var enumValueCount = Enum.GetValues<ShapealignAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IShapeAlignCommands has {coreMethodCount} [ServiceAction] methods but ShapealignAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void ICustomShowCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(ICustomShowCommands));
        var enumValueCount = Enum.GetValues<CustomshowAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"ICustomShowCommands has {coreMethodCount} [ServiceAction] methods but CustomshowAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void IPageSetupCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(IPageSetupCommands));
        var enumValueCount = Enum.GetValues<PagesetupAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"IPageSetupCommands has {coreMethodCount} [ServiceAction] methods but PagesetupAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void ISlideImportCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(ISlideImportCommands));
        var enumValueCount = Enum.GetValues<SlideimportAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"ISlideImportCommands has {coreMethodCount} [ServiceAction] methods but SlideimportAction has only {enumValueCount} enum values.");
    }

    [Fact]
    public void ITagCommands_AllMethodsHaveEnumValues()
    {
        var coreMethodCount = GetServiceActionMethodCount(typeof(ITagCommands));
        var enumValueCount = Enum.GetValues<TagAction>().Length;
        Assert.True(enumValueCount >= coreMethodCount,
            $"ITagCommands has {coreMethodCount} [ServiceAction] methods but TagAction has only {enumValueCount} enum values.");
    }

    // ── Existing mapping tests ───────────────────────────────

    /// <summary>
    /// Verifies all generated action enums have ToActionString mappings via ServiceRegistry.
    /// </summary>
    [Fact]
    public void SlideAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<SlideAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Slide.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Slide.ToActionString(action));
        }
    }

    [Fact]
    public void ShapeAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<ShapeAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Shape.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Shape.ToActionString(action));
        }
    }

    [Fact]
    public void TextAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<TextAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Text.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Text.ToActionString(action));
        }
    }

    // ── NEW: Mapping tests for previously untested action enums ──

    [Fact]
    public void AnimationAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<AnimationAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Animation.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Animation.ToActionString(action));
        }
    }

    [Fact]
    public void ChartAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<ChartAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Chart.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Chart.ToActionString(action));
        }
    }

    [Fact]
    public void DesignAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<DesignAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Design.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Design.ToActionString(action));
        }
    }

    [Fact]
    public void DocpropertyAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<DocpropertyAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Docproperty.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Docproperty.ToActionString(action));
        }
    }

    [Fact]
    public void HyperlinkAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<HyperlinkAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Hyperlink.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Hyperlink.ToActionString(action));
        }
    }

    [Fact]
    public void MediaAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<MediaAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Media.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Media.ToActionString(action));
        }
    }

    [Fact]
    public void SectionAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<SectionAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Section.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Section.ToActionString(action));
        }
    }

    [Fact]
    public void SlideshowAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<SlideshowAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Slideshow.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Slideshow.ToActionString(action));
        }
    }

    [Fact]
    public void SlidetableAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<SlidetableAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Slidetable.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Slidetable.ToActionString(action));
        }
    }

    [Fact]
    public void VbaAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<VbaAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Vba.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Vba.ToActionString(action));
        }
    }

    [Fact]
    public void WindowAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<WindowAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Window.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Window.ToActionString(action));
        }
    }

    [Fact]
    public void NotesAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<NotesAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Notes.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Notes.ToActionString(action));
        }
    }

    [Fact]
    public void MasterAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<MasterAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Master.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Master.ToActionString(action));
        }
    }

    [Fact]
    public void ExportAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<ExportAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Export.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Export.ToActionString(action));
        }
    }

    [Fact]
    public void TransitionAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<TransitionAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Transition.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Transition.ToActionString(action));
        }
    }

    [Fact]
    public void ImageAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<ImageAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Image.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Image.ToActionString(action));
        }
    }

    [Fact]
    public void FileAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<FileAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.File.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.File.ToActionString(action));
        }
    }

    [Fact]
    public void CommentAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<CommentAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Comment.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Comment.ToActionString(action));
        }
    }

    [Fact]
    public void PlaceholderAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<PlaceholderAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Placeholder.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Placeholder.ToActionString(action));
        }
    }

    [Fact]
    public void BackgroundAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<BackgroundAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Background.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Background.ToActionString(action));
        }
    }

    [Fact]
    public void HeaderfooterAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<HeaderfooterAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Headerfooter.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Headerfooter.ToActionString(action));
        }
    }

    [Fact]
    public void SmartartAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<SmartartAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Smartart.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Smartart.ToActionString(action));
        }
    }

    [Fact]
    public void ShapealignAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<ShapealignAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Shapealign.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Shapealign.ToActionString(action));
        }
    }

    [Fact]
    public void CustomshowAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<CustomshowAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Customshow.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Customshow.ToActionString(action));
        }
    }

    [Fact]
    public void PagesetupAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<PagesetupAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Pagesetup.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Pagesetup.ToActionString(action));
        }
    }

    [Fact]
    public void SlideimportAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<SlideimportAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Slideimport.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Slideimport.ToActionString(action));
        }
    }

    [Fact]
    public void TagAction_AllEnumValuesHaveMappings()
    {
        foreach (var action in Enum.GetValues<TagAction>())
        {
            var exception = Record.Exception(() => ServiceRegistry.Tag.ToActionString(action));
            Assert.Null(exception);
            Assert.NotEmpty(ServiceRegistry.Tag.ToActionString(action));
        }
    }

    /// <summary>
    /// Helper: Counts methods with [ServiceAction] attribute in an interface.
    /// </summary>
    private static int GetServiceActionMethodCount(Type interfaceType)
    {
        return interfaceType
            .GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .Where(m => m.GetCustomAttributes()
                .Any(a => a.GetType().Name == "ServiceActionAttribute"))
            .Select(m => m.Name)
            .Distinct(StringComparer.Ordinal)
            .Count();
    }
}




