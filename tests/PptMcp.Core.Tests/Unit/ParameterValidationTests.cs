using PptMcp.Core.Commands.Animation;
using PptMcp.Core.Commands.Background;
using PptMcp.Core.Commands.Chart;
using PptMcp.Core.Commands.Comment;
using PptMcp.Core.Commands.CustomShow;
using PptMcp.Core.Commands.DocumentProperty;
using PptMcp.Core.Commands.Export;
using PptMcp.Core.Commands.Hyperlink;
using PptMcp.Core.Commands.Media;
using PptMcp.Core.Commands.Section;
using PptMcp.Core.Commands.Shape;
using PptMcp.Core.Commands.ShapeAlign;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Commands.SlideImport;
using PptMcp.Core.Commands.SlideTable;
using PptMcp.Core.Commands.SmartArt;
using PptMcp.Core.Commands.Tag;
using PptMcp.Core.Commands.Text;
using PptMcp.Core.Commands.Vba;
using Xunit;

namespace PptMcp.Core.Tests.Unit;

/// <summary>
/// Tests that Core Commands validate required parameters before executing.
/// These tests verify that ArgumentException/ArgumentNullException is thrown
/// for null/empty required parameters WITHOUT needing a PowerPoint COM connection.
/// </summary>
public class ParameterValidationTests
{
    // ── Hyperlink Commands ───────────────────────────────────

    [Fact]
    public void HyperlinkAdd_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new HyperlinkCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Add(null!, 1, null!, "https://example.com"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void HyperlinkAdd_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new HyperlinkCommands();
        Assert.Throws<ArgumentException>(() => commands.Add(null!, 1, shapeName, "https://example.com"));
    }

    [Fact]
    public void HyperlinkAdd_NullAddress_ThrowsArgumentNullException()
    {
        var commands = new HyperlinkCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Add(null!, 1, "Shape1", null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void HyperlinkAdd_EmptyAddress_ThrowsArgumentException(string address)
    {
        var commands = new HyperlinkCommands();
        Assert.Throws<ArgumentException>(() => commands.Add(null!, 1, "Shape1", address));
    }

    [Fact]
    public void HyperlinkRead_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new HyperlinkCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Read(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void HyperlinkRead_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new HyperlinkCommands();
        Assert.Throws<ArgumentException>(() => commands.Read(null!, 1, shapeName));
    }

    // ── VBA Commands ─────────────────────────────────────────

    [Fact]
    public void VbaView_NullModuleName_ThrowsArgumentNullException()
    {
        var commands = new VbaCommands();
        Assert.Throws<ArgumentNullException>(() => commands.View(null!, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void VbaView_EmptyModuleName_ThrowsArgumentException(string moduleName)
    {
        var commands = new VbaCommands();
        Assert.Throws<ArgumentException>(() => commands.View(null!, moduleName));
    }

    [Fact]
    public void VbaImport_NullModuleName_ThrowsArgumentNullException()
    {
        var commands = new VbaCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Import(null!, null!, "Sub Test()\nEnd Sub", 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void VbaImport_EmptyModuleName_ThrowsArgumentException(string moduleName)
    {
        var commands = new VbaCommands();
        Assert.Throws<ArgumentException>(() => commands.Import(null!, moduleName, "Sub Test()\nEnd Sub", 1));
    }

    [Fact]
    public void VbaImport_NullCode_ThrowsArgumentNullException()
    {
        var commands = new VbaCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Import(null!, "Module1", null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void VbaImport_EmptyCode_ThrowsArgumentException(string code)
    {
        var commands = new VbaCommands();
        Assert.Throws<ArgumentException>(() => commands.Import(null!, "Module1", code, 1));
    }

    [Fact]
    public void VbaDelete_NullModuleName_ThrowsArgumentNullException()
    {
        var commands = new VbaCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Delete(null!, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void VbaDelete_EmptyModuleName_ThrowsArgumentException(string moduleName)
    {
        var commands = new VbaCommands();
        Assert.Throws<ArgumentException>(() => commands.Delete(null!, moduleName));
    }

    [Fact]
    public void VbaRun_NullMacroName_ThrowsArgumentNullException()
    {
        var commands = new VbaCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Run(null!, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void VbaRun_EmptyMacroName_ThrowsArgumentException(string macroName)
    {
        var commands = new VbaCommands();
        Assert.Throws<ArgumentException>(() => commands.Run(null!, macroName));
    }

    // ── Section Commands ─────────────────────────────────────

    [Fact]
    public void SectionAdd_NullSectionName_ThrowsArgumentNullException()
    {
        var commands = new SectionCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Add(null!, null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SectionAdd_EmptySectionName_ThrowsArgumentException(string sectionName)
    {
        var commands = new SectionCommands();
        Assert.Throws<ArgumentException>(() => commands.Add(null!, sectionName, 1));
    }

    [Fact]
    public void SectionRename_NullNewName_ThrowsArgumentNullException()
    {
        var commands = new SectionCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Rename(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SectionRename_EmptyNewName_ThrowsArgumentException(string newName)
    {
        var commands = new SectionCommands();
        Assert.Throws<ArgumentException>(() => commands.Rename(null!, 1, newName));
    }

    // ── Animation Commands ───────────────────────────────────

    [Fact]
    public void AnimationAdd_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new AnimationCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Add(null!, 1, null!, 1, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void AnimationAdd_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new AnimationCommands();
        Assert.Throws<ArgumentException>(() => commands.Add(null!, 1, shapeName, 1, 1));
    }

    // ── Chart Commands ───────────────────────────────────────

    [Fact]
    public void ChartGetInfo_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentNullException>(() => commands.GetInfo(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ChartGetInfo_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentException>(() => commands.GetInfo(null!, 1, shapeName));
    }

    [Fact]
    public void ChartSetTitle_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetTitle(null!, 1, null!, "Title"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ChartSetTitle_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentException>(() => commands.SetTitle(null!, 1, shapeName, "Title"));
    }

    [Fact]
    public void ChartSetType_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetType(null!, 1, null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ChartSetType_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentException>(() => commands.SetType(null!, 1, shapeName, 1));
    }

    [Fact]
    public void ChartDelete_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Delete(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ChartDelete_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentException>(() => commands.Delete(null!, 1, shapeName));
    }

    // ── Export Commands ──────────────────────────────────────

    [Fact]
    public void ExportToPdf_NullDestinationPath_ThrowsArgumentNullException()
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ToPdf(null!, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ExportToPdf_EmptyDestinationPath_ThrowsArgumentException(string path)
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentException>(() => commands.ToPdf(null!, path));
    }

    [Fact]
    public void ExportSlideToImage_NullDestinationPath_ThrowsArgumentNullException()
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SlideToImage(null!, 1, null!, 1920, 1080));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ExportSlideToImage_EmptyDestinationPath_ThrowsArgumentException(string path)
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentException>(() => commands.SlideToImage(null!, 1, path, 1920, 1080));
    }

    [Fact]
    public void ExportSaveAs_NullDestinationPath_ThrowsArgumentNullException()
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SaveAs(null!, null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ExportSaveAs_EmptyDestinationPath_ThrowsArgumentException(string path)
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentException>(() => commands.SaveAs(null!, path, 1));
    }

    // ── Shape Commands ───────────────────────────────────────

    [Fact]
    public void ShapeRead_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Read(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeRead_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.Read(null!, 1, shapeName));
    }

    [Fact]
    public void ShapeMoveResize_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.MoveResize(null!, 1, null!, 0, 0, null, null));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeMoveResize_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.MoveResize(null!, 1, shapeName, 0, 0, null, null));
    }

    [Fact]
    public void ShapeDelete_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Delete(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeDelete_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.Delete(null!, 1, shapeName));
    }

    [Fact]
    public void ShapeZOrder_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ZOrder(null!, 1, null!, 0));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeZOrder_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.ZOrder(null!, 1, shapeName, 0));
    }

    [Fact]
    public void ShapeSetFill_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetFill(null!, 1, null!, "#FF0000"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetFill_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetFill(null!, 1, shapeName, "#FF0000"));
    }

    [Fact]
    public void ShapeSetFill_NullColorHex_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetFill(null!, 1, "Shape1", null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetFill_EmptyColorHex_ThrowsArgumentException(string colorHex)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetFill(null!, 1, "Shape1", colorHex));
    }

    [Fact]
    public void ShapeSetLine_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetLine(null!, 1, null!, "#FF0000", 1f));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetLine_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetLine(null!, 1, shapeName, "#FF0000", 1f));
    }

    [Fact]
    public void ShapeSetLine_NullColorHex_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetLine(null!, 1, "Shape1", null!, 1f));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetLine_EmptyColorHex_ThrowsArgumentException(string colorHex)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetLine(null!, 1, "Shape1", colorHex, 1f));
    }

    [Fact]
    public void ShapeSetRotation_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetRotation(null!, 1, null!, 45f));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetRotation_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetRotation(null!, 1, shapeName, 45f));
    }

    [Fact]
    public void ShapeGroup_NullShapeNames_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Group(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeGroup_EmptyShapeNames_ThrowsArgumentException(string shapeNames)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.Group(null!, 1, shapeNames));
    }

    [Fact]
    public void ShapeUngroup_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Ungroup(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeUngroup_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.Ungroup(null!, 1, shapeName));
    }

    [Fact]
    public void ShapeSetAltText_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetAltText(null!, 1, null!, "alt text"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetAltText_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetAltText(null!, 1, shapeName, "alt text"));
    }

    [Fact]
    public void ShapeCopyToSlide_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.CopyToSlide(null!, 1, null!, 2));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeCopyToSlide_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.CopyToSlide(null!, 1, shapeName, 2));
    }

    [Fact]
    public void ShapeSetShadow_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetShadow(null!, 1, null!, true, 3f, 3f));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetShadow_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetShadow(null!, 1, shapeName, true, 3f, 3f));
    }

    [Fact]
    public void ShapeAddConnector_NullStartShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.AddConnector(null!, 1, 1, null!, "End"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeAddConnector_EmptyStartShapeName_ThrowsArgumentException(string startShapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.AddConnector(null!, 1, 1, startShapeName, "End"));
    }

    [Fact]
    public void ShapeAddConnector_NullEndShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.AddConnector(null!, 1, 1, "Start", null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeAddConnector_EmptyEndShapeName_ThrowsArgumentException(string endShapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.AddConnector(null!, 1, 1, "Start", endShapeName));
    }

    [Fact]
    public void ShapeMergeShapes_NullShapeNames_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.MergeShapes(null!, 1, null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeMergeShapes_EmptyShapeNames_ThrowsArgumentException(string shapeNames)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.MergeShapes(null!, 1, shapeNames, 1));
    }

    [Fact]
    public void ShapeDuplicate_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Duplicate(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeDuplicate_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.Duplicate(null!, 1, shapeName));
    }

    [Fact]
    public void ShapeFlip_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Flip(null!, 1, null!, 0));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeFlip_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.Flip(null!, 1, shapeName, 0));
    }

    // ── Text Commands ────────────────────────────────────────

    [Fact]
    public void TextGetText_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.GetText(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextGetText_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.GetText(null!, 1, shapeName));
    }

    [Fact]
    public void TextSetText_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetText(null!, 1, null!, "text"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextSetText_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.SetText(null!, 1, shapeName, "text"));
    }

    [Fact]
    public void TextFormat_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Format(null!, 1, null!, null, null, null, null, null, null, null));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextFormat_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.Format(null!, 1, shapeName, null, null, null, null, null, null, null));
    }

    [Fact]
    public void TextFormatAdvanced_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.FormatAdvanced(null!, 1, null!, null, null, null, null));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextFormatAdvanced_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.FormatAdvanced(null!, 1, shapeName, null, null, null, null));
    }

    // ── Background Commands ──────────────────────────────────

    [Fact]
    public void BackgroundSetColor_NullColorHex_ThrowsArgumentNullException()
    {
        var commands = new BackgroundCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetColor(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void BackgroundSetColor_EmptyColorHex_ThrowsArgumentException(string colorHex)
    {
        var commands = new BackgroundCommands();
        Assert.Throws<ArgumentException>(() => commands.SetColor(null!, 1, colorHex));
    }

    [Fact]
    public void BackgroundSetImage_NullImagePath_ThrowsArgumentNullException()
    {
        var commands = new BackgroundCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetImage(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void BackgroundSetImage_EmptyImagePath_ThrowsArgumentException(string imagePath)
    {
        var commands = new BackgroundCommands();
        Assert.Throws<ArgumentException>(() => commands.SetImage(null!, 1, imagePath));
    }

    // ── SmartArt Commands ────────────────────────────────────

    [Fact]
    public void SmartArtGetInfo_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new SmartArtCommands();
        Assert.Throws<ArgumentNullException>(() => commands.GetInfo(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SmartArtGetInfo_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new SmartArtCommands();
        Assert.Throws<ArgumentException>(() => commands.GetInfo(null!, 1, shapeName));
    }

    [Fact]
    public void SmartArtAddNode_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new SmartArtCommands();
        Assert.Throws<ArgumentNullException>(() => commands.AddNode(null!, 1, null!, "text"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SmartArtAddNode_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new SmartArtCommands();
        Assert.Throws<ArgumentException>(() => commands.AddNode(null!, 1, shapeName, "text"));
    }

    [Fact]
    public void SmartArtAddNode_NullText_ThrowsArgumentNullException()
    {
        var commands = new SmartArtCommands();
        Assert.Throws<ArgumentNullException>(() => commands.AddNode(null!, 1, "SmartArt1", null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SmartArtAddNode_EmptyText_ThrowsArgumentException(string text)
    {
        var commands = new SmartArtCommands();
        Assert.Throws<ArgumentException>(() => commands.AddNode(null!, 1, "SmartArt1", text));
    }

    // ── Comment Commands ─────────────────────────────────────

    [Fact]
    public void CommentAdd_NullText_ThrowsArgumentNullException()
    {
        var commands = new CommentCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Add(null!, 1, null!, "Author", 0, 0));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void CommentAdd_EmptyText_ThrowsArgumentException(string text)
    {
        var commands = new CommentCommands();
        Assert.Throws<ArgumentException>(() => commands.Add(null!, 1, text, "Author", 0, 0));
    }

    [Fact]
    public void CommentAdd_NullAuthor_ThrowsArgumentNullException()
    {
        var commands = new CommentCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Add(null!, 1, "Comment text", null!, 0, 0));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void CommentAdd_EmptyAuthor_ThrowsArgumentException(string author)
    {
        var commands = new CommentCommands();
        Assert.Throws<ArgumentException>(() => commands.Add(null!, 1, "Comment text", author, 0, 0));
    }

    // ── Custom Show Commands ─────────────────────────────────

    [Fact]
    public void CustomShowCreate_NullShowName_ThrowsArgumentNullException()
    {
        var commands = new CustomShowCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Create(null!, null!, "1,2,3"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void CustomShowCreate_EmptyShowName_ThrowsArgumentException(string showName)
    {
        var commands = new CustomShowCommands();
        Assert.Throws<ArgumentException>(() => commands.Create(null!, showName, "1,2,3"));
    }

    [Fact]
    public void CustomShowCreate_NullSlideIndices_ThrowsArgumentNullException()
    {
        var commands = new CustomShowCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Create(null!, "Show1", null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void CustomShowCreate_EmptySlideIndices_ThrowsArgumentException(string slideIndices)
    {
        var commands = new CustomShowCommands();
        Assert.Throws<ArgumentException>(() => commands.Create(null!, "Show1", slideIndices));
    }

    [Fact]
    public void CustomShowDelete_NullShowName_ThrowsArgumentNullException()
    {
        var commands = new CustomShowCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Delete(null!, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void CustomShowDelete_EmptyShowName_ThrowsArgumentException(string showName)
    {
        var commands = new CustomShowCommands();
        Assert.Throws<ArgumentException>(() => commands.Delete(null!, showName));
    }

    // ── Tag Commands ─────────────────────────────────────────

    [Fact]
    public void TagSetTag_NullTagName_ThrowsArgumentNullException()
    {
        var commands = new TagCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetTag(null!, 1, null, null!, "value"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TagSetTag_EmptyTagName_ThrowsArgumentException(string tagName)
    {
        var commands = new TagCommands();
        Assert.Throws<ArgumentException>(() => commands.SetTag(null!, 1, null, tagName, "value"));
    }

    [Fact]
    public void TagDeleteTag_NullTagName_ThrowsArgumentNullException()
    {
        var commands = new TagCommands();
        Assert.Throws<ArgumentNullException>(() => commands.DeleteTag(null!, 1, null, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TagDeleteTag_EmptyTagName_ThrowsArgumentException(string tagName)
    {
        var commands = new TagCommands();
        Assert.Throws<ArgumentException>(() => commands.DeleteTag(null!, 1, null, tagName));
    }

    // ── Slide Import Commands ────────────────────────────────

    [Fact]
    public void SlideImportImportSlides_NullSourceFilePath_ThrowsArgumentNullException()
    {
        var commands = new SlideImportCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ImportSlides(null!, null!, "1,2", 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SlideImportImportSlides_EmptySourceFilePath_ThrowsArgumentException(string sourceFilePath)
    {
        var commands = new SlideImportCommands();
        Assert.Throws<ArgumentException>(() => commands.ImportSlides(null!, sourceFilePath, "1,2", 1));
    }

    // ── Media Commands ───────────────────────────────────────

    [Fact]
    public void MediaGetInfo_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new MediaCommands();
        Assert.Throws<ArgumentNullException>(() => commands.GetInfo(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void MediaGetInfo_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new MediaCommands();
        Assert.Throws<ArgumentException>(() => commands.GetInfo(null!, 1, shapeName));
    }

    // ── Slide Commands ───────────────────────────────────────

    [Fact]
    public void SlideSetName_NullName_ThrowsArgumentNullException()
    {
        var commands = new SlideCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetName(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SlideSetName_EmptyName_ThrowsArgumentException(string name)
    {
        var commands = new SlideCommands();
        Assert.Throws<ArgumentException>(() => commands.SetName(null!, 1, name));
    }

    // ── Document Property Commands ───────────────────────────

    [Fact]
    public void DocumentPropertyGetCustom_NullPropertyName_ThrowsArgumentNullException()
    {
        var commands = new DocumentPropertyCommands();
        Assert.Throws<ArgumentNullException>(() => commands.GetCustom(null!, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void DocumentPropertyGetCustom_EmptyPropertyName_ThrowsArgumentException(string propertyName)
    {
        var commands = new DocumentPropertyCommands();
        Assert.Throws<ArgumentException>(() => commands.GetCustom(null!, propertyName));
    }

    [Fact]
    public void DocumentPropertySetCustom_NullPropertyName_ThrowsArgumentNullException()
    {
        var commands = new DocumentPropertyCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetCustom(null!, null!, "value"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void DocumentPropertySetCustom_EmptyPropertyName_ThrowsArgumentException(string propertyName)
    {
        var commands = new DocumentPropertyCommands();
        Assert.Throws<ArgumentException>(() => commands.SetCustom(null!, propertyName, "value"));
    }

    // ── Shape Align Commands ─────────────────────────────────

    [Fact]
    public void ShapeAlignAlign_NullShapeNames_ThrowsArgumentNullException()
    {
        var commands = new ShapeAlignCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Align(null!, 1, null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeAlignAlign_EmptyShapeNames_ThrowsArgumentException(string shapeNames)
    {
        var commands = new ShapeAlignCommands();
        Assert.Throws<ArgumentException>(() => commands.Align(null!, 1, shapeNames, 1));
    }

    [Fact]
    public void ShapeAlignDistribute_NullShapeNames_ThrowsArgumentNullException()
    {
        var commands = new ShapeAlignCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Distribute(null!, 1, null!, 0));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeAlignDistribute_EmptyShapeNames_ThrowsArgumentException(string shapeNames)
    {
        var commands = new ShapeAlignCommands();
        Assert.Throws<ArgumentException>(() => commands.Distribute(null!, 1, shapeNames, 0));
    }

    // ── Slide Table Commands ─────────────────────────────────

    [Fact]
    public void SlideTableFormatCell_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new SlideTableCommands();
        Assert.Throws<ArgumentNullException>(() => commands.FormatCell(null!, 1, null!, 1, 1, null, null, 0, null));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SlideTableFormatCell_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new SlideTableCommands();
        Assert.Throws<ArgumentException>(() => commands.FormatCell(null!, 1, shapeName, 1, 1, null, null, 0, null));
    }
}
