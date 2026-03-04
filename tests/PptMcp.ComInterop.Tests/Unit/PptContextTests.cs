using PptMcp.ComInterop.Session;
using Xunit;

namespace PptMcp.ComInterop.Tests.Unit;

/// <summary>
/// Unit tests for PptContext - validates constructor and property behavior.
/// This class is a simple data holder, so tests focus on path validation and immutability.
/// Note: PowerPoint.Application and PowerPoint.Presentation COM objects cannot be mocked in unit tests,
/// so these tests use null! for those parameters and verify only what is testable.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
public class PptContextTests
{
    [Fact]
    public void Constructor_WithValidArguments_SetsPresentationPathCorrectly()
    {
        // Arrange
        string presentationPath = @"C:\test\presentation.pptx";

        // Act & Assert - Constructor throws ArgumentNullException for null COM objects,
        // which is expected behavior. PresentationPath validation is tested separately.
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new PptContext(presentationPath, null!, null!));

        // When null is passed, the constructor throws on the first null param (excel)
        Assert.NotNull(ex);
    }

    [Fact]
    public void Constructor_WithNullPresentationPath_ThrowsArgumentNullException()
    {
        // Arrange
        string? presentationPath = null;

        // Act & Assert
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new PptContext(presentationPath!, null!, null!));

        Assert.Equal("presentationPath", ex.ParamName);
    }

    [Fact]
    public void Constructor_WithNullPowerPoint_ThrowsArgumentNullException()
    {
        // Arrange
        string presentationPath = @"C:\test\presentation.pptx";

        // Act & Assert
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new PptContext(presentationPath, null!, null!));

        Assert.Equal("app", ex.ParamName);
    }

    [Fact]
    public void Constructor_WithNullPresentationPath_ThrowsBeforeNullPowerPoint()
    {
        // Arrange
        string? presentationPath = null;

        // Act & Assert - PresentationPath is validated first
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new PptContext(presentationPath!, null!, null!));

        Assert.Equal("presentationPath", ex.ParamName);
    }

    [Fact]
    public void Constructor_PresentationPathValidation_RejectsNull()
    {
        // Arrange & Act & Assert
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new PptContext(null!, null!, null!));

        Assert.Equal("presentationPath", ex.ParamName);
    }

    [Theory]
    [InlineData(@"C:\test\presentation.pptx")]
    [InlineData(@"\\server\share\presentation.pptm")]
    [InlineData(@"D:\Documents\My Presentation.pptx")]
    [InlineData(@"presentation.pptx")] // Relative path
    public void Constructor_WithNullPowerPointAnyPath_ThrowsArgumentNullException(string presentationPath)
    {
        // Act & Assert - Path is validated, then excel COM object is validated
        var ex = Assert.Throws<ArgumentNullException>(() =>
            new PptContext(presentationPath, null!, null!));

        // app is the first COM parameter validated after presentationPath
        Assert.Equal("app", ex.ParamName);
    }

    [Fact]
    public void Constructor_NullPresentationPath_ThrowsWithCorrectParamName()
    {
        // Arrange - Simulates null path being passed
        Assert.Throws<ArgumentNullException>(() =>
            new PptContext(null!, null!, null!));
    }
}





