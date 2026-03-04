using PptMcp.Core.Commands.Hyperlink;
using PptMcp.Core.Commands.Section;
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
}
