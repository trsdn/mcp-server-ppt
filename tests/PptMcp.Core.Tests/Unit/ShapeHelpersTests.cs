using PptMcp.Core.Commands.Slide;
using Xunit;

namespace PptMcp.Core.Tests.Unit;

/// <summary>
/// Unit tests for ShapeHelpers utility methods.
/// </summary>
public class ShapeHelpersTests
{
    [Theory]
    [InlineData(1, "AutoShape")]
    [InlineData(2, "Callout")]
    [InlineData(3, "Chart")]
    [InlineData(4, "Comment")]
    [InlineData(5, "FreeForm")]
    [InlineData(6, "Group")]
    [InlineData(7, "EmbeddedOLEObject")]
    [InlineData(8, "FormControl")]
    [InlineData(9, "Line")]
    [InlineData(10, "LinkedOLEObject")]
    [InlineData(11, "LinkedPicture")]
    [InlineData(12, "OLEControlObject")]
    [InlineData(13, "Picture")]
    [InlineData(14, "Placeholder")]
    [InlineData(15, "TextEffect")]
    [InlineData(16, "MediaObject")]
    [InlineData(17, "TextBox")]
    [InlineData(19, "Table")]
    [InlineData(20, "Canvas")]
    [InlineData(21, "Diagram")]
    [InlineData(22, "Ink")]
    [InlineData(23, "InkComment")]
    [InlineData(24, "SmartArt")]
    [InlineData(25, "Slicer")]
    [InlineData(26, "WebVideo")]
    [InlineData(27, "ContentApp")]
    [InlineData(28, "Graphic")]
    [InlineData(29, "LinkedGraphic")]
    [InlineData(30, "3DModel")]
    [InlineData(31, "Linked3DModel")]
    public void GetShapeTypeName_KnownTypes_ReturnsExpectedName(int msoType, string expectedName)
    {
        var result = ShapeHelpers.GetShapeTypeName(msoType);
        Assert.Equal(expectedName, result);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(18)]
    [InlineData(32)]
    [InlineData(100)]
    [InlineData(999)]
    public void GetShapeTypeName_UnknownTypes_ReturnsUnknownWithValue(int msoType)
    {
        var result = ShapeHelpers.GetShapeTypeName(msoType);
        Assert.StartsWith("Unknown(", result);
        Assert.Contains(msoType.ToString(System.Globalization.CultureInfo.InvariantCulture), result);
        Assert.EndsWith(")", result);
    }

    [Fact]
    public void GetShapeTypeName_MsoShapeType18_IsNotDefined()
    {
        // msoShapeType 18 is intentionally skipped in PowerPoint COM API
        var result = ShapeHelpers.GetShapeTypeName(18);
        Assert.Equal("Unknown(18)", result);
    }
}
