using PptMcp.Core.Commands.Animation;
using PptMcp.Core.Commands.Background;
using PptMcp.Core.Commands.Chart;
using PptMcp.Core.Commands.Comment;
using PptMcp.Core.Commands.CustomShow;
using PptMcp.Core.Commands.DocumentProperty;
using PptMcp.Core.Commands.Export;
using PptMcp.Core.Commands.Hyperlink;
using PptMcp.Core.Commands.Image;
using PptMcp.Core.Commands.Master;
using PptMcp.Core.Commands.Media;
using PptMcp.Core.Commands.Proofing;
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

    // ── Chart Commands (Additional) ─────────────────────────

    [Fact]
    public void ChartSetData_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetData(null!, 1, null!, new List<List<object?>>()));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ChartSetData_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentException>(() => commands.SetData(null!, 1, shapeName, new List<List<object?>>()));
    }

    [Fact]
    public void ChartSetLegend_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetLegend(null!, 1, null!, true, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ChartSetLegend_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentException>(() => commands.SetLegend(null!, 1, shapeName, true, 1));
    }

    [Fact]
    public void ChartReadData_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ReadData(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ChartReadData_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentException>(() => commands.ReadData(null!, 1, shapeName));
    }

    [Fact]
    public void ChartSetAxisTitle_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetAxisTitle(null!, 1, null!, 1, "Title"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ChartSetAxisTitle_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentException>(() => commands.SetAxisTitle(null!, 1, shapeName, 1, "Title"));
    }

    [Fact]
    public void ChartSetAxisTitle_NullTitle_ThrowsArgumentNullException()
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetAxisTitle(null!, 1, "Chart1", 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ChartSetAxisTitle_EmptyTitle_ThrowsArgumentException(string title)
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentException>(() => commands.SetAxisTitle(null!, 1, "Chart1", 1, title));
    }

    [Fact]
    public void ChartToggleDataTable_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ToggleDataTable(null!, 1, null!, true));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ChartToggleDataTable_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ChartCommands();
        Assert.Throws<ArgumentException>(() => commands.ToggleDataTable(null!, 1, shapeName, true));
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

    // ── Export Commands (Additional) ─────────────────────────

    [Fact]
    public void ExportToVideo_NullDestinationPath_ThrowsArgumentNullException()
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ToVideo(null!, null!, 5, 1080));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ExportToVideo_EmptyDestinationPath_ThrowsArgumentException(string path)
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentException>(() => commands.ToVideo(null!, path, 5, 1080));
    }

    [Fact]
    public void ExportAllSlidesToImages_NullDestinationDirectory_ThrowsArgumentNullException()
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentNullException>(() => commands.AllSlidesToImages(null!, null!, 1920, 1080));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ExportAllSlidesToImages_EmptyDestinationDirectory_ThrowsArgumentException(string dir)
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentException>(() => commands.AllSlidesToImages(null!, dir, 1920, 1080));
    }

    [Fact]
    public void ExportExtractText_NullDestinationPath_ThrowsArgumentNullException()
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ExtractText(null!, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ExportExtractText_EmptyDestinationPath_ThrowsArgumentException(string path)
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentException>(() => commands.ExtractText(null!, path));
    }

    [Fact]
    public void ExportExtractImages_NullDestinationDirectory_ThrowsArgumentNullException()
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ExtractImages(null!, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ExportExtractImages_EmptyDestinationDirectory_ThrowsArgumentException(string dir)
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentException>(() => commands.ExtractImages(null!, dir));
    }

    [Fact]
    public void ExportSaveCopy_NullDestinationPath_ThrowsArgumentNullException()
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SaveCopy(null!, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ExportSaveCopy_EmptyDestinationPath_ThrowsArgumentException(string path)
    {
        var commands = new ExportCommands();
        Assert.Throws<ArgumentException>(() => commands.SaveCopy(null!, path));
    }

    // ── Image Commands ──────────────────────────────────────

    [Fact]
    public void ImageCrop_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ImageCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Crop(null!, 1, null!, 0f, 0f, 0f, 0f));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ImageCrop_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ImageCommands();
        Assert.Throws<ArgumentException>(() => commands.Crop(null!, 1, shapeName, 0f, 0f, 0f, 0f));
    }

    [Fact]
    public void ImageSetBrightnessContrast_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ImageCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetBrightnessContrast(null!, 1, null!, 0.5f, 0.5f));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ImageSetBrightnessContrast_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ImageCommands();
        Assert.Throws<ArgumentException>(() => commands.SetBrightnessContrast(null!, 1, shapeName, 0.5f, 0.5f));
    }

    [Fact]
    public void ImageSetTransparentColor_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ImageCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetTransparentColor(null!, 1, null!, "#FFFFFF"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ImageSetTransparentColor_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ImageCommands();
        Assert.Throws<ArgumentException>(() => commands.SetTransparentColor(null!, 1, shapeName, "#FFFFFF"));
    }

    [Fact]
    public void ImageSetTransparentColor_NullColorHex_ThrowsArgumentNullException()
    {
        var commands = new ImageCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetTransparentColor(null!, 1, "Image1", null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ImageSetTransparentColor_EmptyColorHex_ThrowsArgumentException(string colorHex)
    {
        var commands = new ImageCommands();
        Assert.Throws<ArgumentException>(() => commands.SetTransparentColor(null!, 1, "Image1", colorHex));
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

    // ── Shape Commands (Additional) ─────────────────────────

    [Fact]
    public void ShapeSetTextFrame_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetTextFrame(null!, 1, null!, null, null, null, null, null, null));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetTextFrame_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetTextFrame(null!, 1, shapeName, null, null, null, null, null, null));
    }

    [Fact]
    public void ShapeReadFill_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ReadFill(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeReadFill_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.ReadFill(null!, 1, shapeName));
    }

    [Fact]
    public void ShapeReadLine_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ReadLine(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeReadLine_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.ReadLine(null!, 1, shapeName));
    }

    [Fact]
    public void ShapeSetGradientFill_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetGradientFill(null!, 1, null!, "#FF0000", "#0000FF", 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetGradientFill_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetGradientFill(null!, 1, shapeName, "#FF0000", "#0000FF", 1));
    }

    [Fact]
    public void ShapeSetGradientFill_NullColor1_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetGradientFill(null!, 1, "Shape1", null!, "#0000FF", 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetGradientFill_EmptyColor1_ThrowsArgumentException(string color1)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetGradientFill(null!, 1, "Shape1", color1, "#0000FF", 1));
    }

    [Fact]
    public void ShapeSetGradientFill_NullColor2_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetGradientFill(null!, 1, "Shape1", "#FF0000", null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetGradientFill_EmptyColor2_ThrowsArgumentException(string color2)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetGradientFill(null!, 1, "Shape1", "#FF0000", color2, 1));
    }

    [Fact]
    public void ShapeSetGlow_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetGlow(null!, 1, null!, 5f, "#FF0000"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetGlow_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetGlow(null!, 1, shapeName, 5f, "#FF0000"));
    }

    [Fact]
    public void ShapeSetGlow_NullColorHex_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetGlow(null!, 1, "Shape1", 5f, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetGlow_EmptyColorHex_ThrowsArgumentException(string colorHex)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetGlow(null!, 1, "Shape1", 5f, colorHex));
    }

    [Fact]
    public void ShapeSetReflection_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetReflection(null!, 1, null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetReflection_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetReflection(null!, 1, shapeName, 1));
    }

    [Fact]
    public void ShapeSetOpacity_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetOpacity(null!, 1, null!, 0.5f));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetOpacity_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetOpacity(null!, 1, shapeName, 0.5f));
    }

    [Fact]
    public void ShapeSetActionSettings_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetActionSettings(null!, 1, null!, 1, null));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetActionSettings_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetActionSettings(null!, 1, shapeName, 1, null));
    }

    [Fact]
    public void ShapeScale_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Scale(null!, 1, null!, 1.5f, 1.5f));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeScale_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.Scale(null!, 1, shapeName, 1.5f, 1.5f));
    }

    [Fact]
    public void ShapeSetLockAspectRatio_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetLockAspectRatio(null!, 1, null!, true));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetLockAspectRatio_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetLockAspectRatio(null!, 1, shapeName, true));
    }

    [Fact]
    public void ShapeSetSoftEdge_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetSoftEdge(null!, 1, null!, 5f));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSetSoftEdge_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.SetSoftEdge(null!, 1, shapeName, 5f));
    }

    [Fact]
    public void ShapeReadShadow_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ReadShadow(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeReadShadow_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.ReadShadow(null!, 1, shapeName));
    }

    [Fact]
    public void ShapeAddTextEffect_NullText_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.AddTextEffect(null!, 1, 0, null!, "Arial", 24f, 100f, 100f));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeAddTextEffect_EmptyText_ThrowsArgumentException(string text)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.AddTextEffect(null!, 1, 0, text, "Arial", 24f, 100f, 100f));
    }

    [Fact]
    public void ShapeAddTextEffect_NullFontName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.AddTextEffect(null!, 1, 0, "Text", null!, 24f, 100f, 100f));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeAddTextEffect_EmptyFontName_ThrowsArgumentException(string fontName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.AddTextEffect(null!, 1, 0, "Text", fontName, 24f, 100f, 100f));
    }

    [Fact]
    public void ShapeSet3D_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.Set3D(null!, 1, null!, null, null, null, null, null));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeSet3D_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.Set3D(null!, 1, shapeName, null, null, null, null, null));
    }

    [Fact]
    public void ShapeCopyFormatting_NullSourceShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.CopyFormatting(null!, 1, null!, "Target"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeCopyFormatting_EmptySourceShapeName_ThrowsArgumentException(string sourceShapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.CopyFormatting(null!, 1, sourceShapeName, "Target"));
    }

    [Fact]
    public void ShapeCopyFormatting_NullTargetShapeName_ThrowsArgumentNullException()
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentNullException>(() => commands.CopyFormatting(null!, 1, "Source", null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ShapeCopyFormatting_EmptyTargetShapeName_ThrowsArgumentException(string targetShapeName)
    {
        var commands = new ShapeCommands();
        Assert.Throws<ArgumentException>(() => commands.CopyFormatting(null!, 1, "Source", targetShapeName));
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

    // ── Text Commands (Additional) ──────────────────────────

    [Fact]
    public void TextSetSpacing_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetSpacing(null!, 1, null!, null, null, null, null));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextSetSpacing_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.SetSpacing(null!, 1, shapeName, null, null, null, null));
    }

    [Fact]
    public void TextSetBullets_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetBullets(null!, 1, null!, 1, null, 0));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextSetBullets_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.SetBullets(null!, 1, shapeName, 1, null, 0));
    }

    [Fact]
    public void TextInsertLink_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.InsertLink(null!, 1, null!, "Click here", "https://example.com"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextInsertLink_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.InsertLink(null!, 1, shapeName, "Click here", "https://example.com"));
    }

    [Fact]
    public void TextInsertLink_NullLinkText_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.InsertLink(null!, 1, "Shape1", null!, "https://example.com"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextInsertLink_EmptyLinkText_ThrowsArgumentException(string linkText)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.InsertLink(null!, 1, "Shape1", linkText, "https://example.com"));
    }

    [Fact]
    public void TextInsertLink_NullUrl_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.InsertLink(null!, 1, "Shape1", "Click here", null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextInsertLink_EmptyUrl_ThrowsArgumentException(string url)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.InsertLink(null!, 1, "Shape1", "Click here", url));
    }

    [Fact]
    public void TextChangeCase_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ChangeCase(null!, 1, null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextChangeCase_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.ChangeCase(null!, 1, shapeName, 1));
    }

    [Fact]
    public void TextReadSpacing_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ReadSpacing(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextReadSpacing_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.ReadSpacing(null!, 1, shapeName));
    }

    [Fact]
    public void TextReadBullets_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ReadBullets(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextReadBullets_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.ReadBullets(null!, 1, shapeName));
    }

    [Fact]
    public void TextInsertSymbol_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.InsertSymbol(null!, 1, null!, "Wingdings", 65));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextInsertSymbol_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.InsertSymbol(null!, 1, shapeName, "Wingdings", 65));
    }

    [Fact]
    public void TextInsertSymbol_NullFontName_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.InsertSymbol(null!, 1, "Shape1", null!, 65));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextInsertSymbol_EmptyFontName_ThrowsArgumentException(string fontName)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.InsertSymbol(null!, 1, "Shape1", fontName, 65));
    }

    [Fact]
    public void TextInsertDateTime_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.InsertDateTime(null!, 1, null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextInsertDateTime_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.InsertDateTime(null!, 1, shapeName, 1));
    }

    [Fact]
    public void TextInsertSlideNumber_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentNullException>(() => commands.InsertSlideNumber(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TextInsertSlideNumber_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new TextCommands();
        Assert.Throws<ArgumentException>(() => commands.InsertSlideNumber(null!, 1, shapeName));
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

    // ── SmartArt Commands (Additional) ──────────────────────

    [Fact]
    public void SmartArtSetLayout_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new SmartArtCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetLayout(null!, 1, null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SmartArtSetLayout_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new SmartArtCommands();
        Assert.Throws<ArgumentException>(() => commands.SetLayout(null!, 1, shapeName, 1));
    }

    [Fact]
    public void SmartArtSetStyle_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new SmartArtCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetStyle(null!, 1, null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SmartArtSetStyle_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new SmartArtCommands();
        Assert.Throws<ArgumentException>(() => commands.SetStyle(null!, 1, shapeName, 1));
    }

    [Fact]
    public void SmartArtDeleteNode_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new SmartArtCommands();
        Assert.Throws<ArgumentNullException>(() => commands.DeleteNode(null!, 1, null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SmartArtDeleteNode_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new SmartArtCommands();
        Assert.Throws<ArgumentException>(() => commands.DeleteNode(null!, 1, shapeName, 1));
    }

    [Fact]
    public void SmartArtChangeNodeLevel_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new SmartArtCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ChangeNodeLevel(null!, 1, null!, 1, true));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SmartArtChangeNodeLevel_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new SmartArtCommands();
        Assert.Throws<ArgumentException>(() => commands.ChangeNodeLevel(null!, 1, shapeName, 1, true));
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

    // ── Media Commands (Additional) ─────────────────────────

    [Fact]
    public void MediaSetPlayback_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new MediaCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetPlayback(null!, 1, null!, null, null, null, null));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void MediaSetPlayback_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new MediaCommands();
        Assert.Throws<ArgumentException>(() => commands.SetPlayback(null!, 1, shapeName, null, null, null, null));
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

    // ── Slide Table Commands (Additional) ────────────────────

    [Fact]
    public void SlideTableReadCell_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new SlideTableCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ReadCell(null!, 1, null!, 1, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SlideTableReadCell_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new SlideTableCommands();
        Assert.Throws<ArgumentException>(() => commands.ReadCell(null!, 1, shapeName, 1, 1));
    }

    [Fact]
    public void SlideTableWriteRow_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new SlideTableCommands();
        Assert.Throws<ArgumentNullException>(() => commands.WriteRow(null!, 1, null!, 1, "a,b,c"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SlideTableWriteRow_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new SlideTableCommands();
        Assert.Throws<ArgumentException>(() => commands.WriteRow(null!, 1, shapeName, 1, "a,b,c"));
    }

    [Fact]
    public void SlideTableWriteRow_NullValues_ThrowsArgumentNullException()
    {
        var commands = new SlideTableCommands();
        Assert.Throws<ArgumentNullException>(() => commands.WriteRow(null!, 1, "Table1", 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SlideTableWriteRow_EmptyValues_ThrowsArgumentException(string values)
    {
        var commands = new SlideTableCommands();
        Assert.Throws<ArgumentException>(() => commands.WriteRow(null!, 1, "Table1", 1, values));
    }

    [Fact]
    public void SlideTableReadRow_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new SlideTableCommands();
        Assert.Throws<ArgumentNullException>(() => commands.ReadRow(null!, 1, null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SlideTableReadRow_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new SlideTableCommands();
        Assert.Throws<ArgumentException>(() => commands.ReadRow(null!, 1, shapeName, 1));
    }

    [Fact]
    public void SlideTableSetCellBorder_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new SlideTableCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetCellBorder(null!, 1, null!, 1, 1, "#000000", 1f));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SlideTableSetCellBorder_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new SlideTableCommands();
        Assert.Throws<ArgumentException>(() => commands.SetCellBorder(null!, 1, shapeName, 1, 1, "#000000", 1f));
    }

    [Fact]
    public void SlideTableSetCellBorder_NullColorHex_ThrowsArgumentNullException()
    {
        var commands = new SlideTableCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetCellBorder(null!, 1, "Table1", 1, 1, null!, 1f));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SlideTableSetCellBorder_EmptyColorHex_ThrowsArgumentException(string colorHex)
    {
        var commands = new SlideTableCommands();
        Assert.Throws<ArgumentException>(() => commands.SetCellBorder(null!, 1, "Table1", 1, 1, colorHex, 1f));
    }

    // ── Background Commands (Additional) ────────────────────

    [Fact]
    public void BackgroundSetGradient_NullColor1_ThrowsArgumentNullException()
    {
        var commands = new BackgroundCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetGradient(null!, 1, null!, "#0000FF", 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void BackgroundSetGradient_EmptyColor1_ThrowsArgumentException(string color1)
    {
        var commands = new BackgroundCommands();
        Assert.Throws<ArgumentException>(() => commands.SetGradient(null!, 1, color1, "#0000FF", 1));
    }

    [Fact]
    public void BackgroundSetGradient_NullColor2_ThrowsArgumentNullException()
    {
        var commands = new BackgroundCommands();
        Assert.Throws<ArgumentNullException>(() => commands.SetGradient(null!, 1, "#FF0000", null!, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void BackgroundSetGradient_EmptyColor2_ThrowsArgumentException(string color2)
    {
        var commands = new BackgroundCommands();
        Assert.Throws<ArgumentException>(() => commands.SetGradient(null!, 1, "#FF0000", color2, 1));
    }

    // ── Proofing Commands ───────────────────────────────────

    [Fact]
    public void ProofingGetLanguage_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new ProofingCommands();
        Assert.Throws<ArgumentNullException>(() => commands.GetLanguage(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ProofingGetLanguage_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new ProofingCommands();
        Assert.Throws<ArgumentException>(() => commands.GetLanguage(null!, 1, shapeName));
    }

    // ── Master Commands ─────────────────────────────────────

    [Fact]
    public void MasterEditShapeText_NullShapeName_ThrowsArgumentNullException()
    {
        var commands = new MasterCommands();
        Assert.Throws<ArgumentNullException>(() => commands.EditShapeText(null!, 1, null!, "text"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void MasterEditShapeText_EmptyShapeName_ThrowsArgumentException(string shapeName)
    {
        var commands = new MasterCommands();
        Assert.Throws<ArgumentException>(() => commands.EditShapeText(null!, 1, shapeName, "text"));
    }

    // ── Slide Commands (Additional) ─────────────────────────

    [Fact]
    public void SlideCloneWithReplace_NullSearchText_ThrowsArgumentNullException()
    {
        var commands = new SlideCommands();
        Assert.Throws<ArgumentNullException>(() => commands.CloneWithReplace(null!, 1, 1, null!, "replace"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SlideCloneWithReplace_EmptySearchText_ThrowsArgumentException(string searchText)
    {
        var commands = new SlideCommands();
        Assert.Throws<ArgumentException>(() => commands.CloneWithReplace(null!, 1, 1, searchText, "replace"));
    }

    [Fact]
    public void SlideCloneWithReplace_NullReplaceText_ThrowsArgumentNullException()
    {
        var commands = new SlideCommands();
        Assert.Throws<ArgumentNullException>(() => commands.CloneWithReplace(null!, 1, 1, "search", null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SlideCloneWithReplace_EmptyReplaceText_ThrowsArgumentException(string replaceText)
    {
        var commands = new SlideCommands();
        Assert.Throws<ArgumentException>(() => commands.CloneWithReplace(null!, 1, 1, "search", replaceText));
    }

    [Fact]
    public void SlideGetThumbnail_NullDestinationPath_ThrowsArgumentNullException()
    {
        var commands = new SlideCommands();
        Assert.Throws<ArgumentNullException>(() => commands.GetThumbnail(null!, 1, null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SlideGetThumbnail_EmptyDestinationPath_ThrowsArgumentException(string path)
    {
        var commands = new SlideCommands();
        Assert.Throws<ArgumentException>(() => commands.GetThumbnail(null!, 1, path));
    }
}
