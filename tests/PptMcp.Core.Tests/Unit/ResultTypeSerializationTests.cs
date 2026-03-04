using System.Text.Json;
using PptMcp.Core.Models;
using Xunit;

namespace PptMcp.Core.Tests.Unit;

/// <summary>
/// Validates JSON serialization behavior of result types,
/// ensuring null properties are omitted and camelCase naming works correctly.
/// </summary>
public class ResultTypeSerializationTests
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    [Fact]
    public void OperationResult_Success_OmitsNullFields()
    {
        var result = new OperationResult { Success = true, Action = "create", Message = "Done" };
        var json = JsonSerializer.Serialize(result, JsonOptions);

        Assert.Contains("\"success\":true", json);
        Assert.Contains("\"action\":\"create\"", json);
        Assert.DoesNotContain("errorMessage", json);
        Assert.DoesNotContain("filePath", json);
    }

    [Fact]
    public void OperationResult_Failure_IncludesErrorMessage()
    {
        var result = new OperationResult { Success = false, ErrorMessage = "Not found" };
        var json = JsonSerializer.Serialize(result, JsonOptions);

        Assert.Contains("\"success\":false", json);
        Assert.Contains("\"errorMessage\":\"Not found\"", json);
    }

    [Fact]
    public void SlideListResult_WithSlides_SerializesCorrectly()
    {
        var result = new SlideListResult
        {
            Success = true,
            Slides =
            [
                new SlideInfo
                {
                    SlideIndex = 1,
                    SlideNumber = 1,
                    SlideId = "256",
                    LayoutName = "Title Slide",
                    MasterName = "Office Theme",
                    ShapeCount = 3
                }
            ]
        };
        var json = JsonSerializer.Serialize(result, JsonOptions);

        Assert.Contains("\"slideIndex\":1", json);
        Assert.Contains("\"layoutName\":\"Title Slide\"", json);
        Assert.Contains("\"shapeCount\":3", json);
    }

    [Fact]
    public void ShapeInfo_NullOptionalFields_AreOmitted()
    {
        var info = new ShapeInfo
        {
            ShapeId = 1,
            Name = "Rectangle 1",
            ShapeType = "AutoShape",
            Width = 100f,
            Height = 50f
        };
        var json = JsonSerializer.Serialize(info, JsonOptions);

        Assert.Contains("\"name\":\"Rectangle 1\"", json);
        Assert.DoesNotContain("\"text\":", json);
        Assert.DoesNotContain("\"alternativeText\":", json);
        Assert.DoesNotContain("\"placeholderType\":", json);
        Assert.DoesNotContain("\"groupItems\":", json);
    }

    [Fact]
    public void TextResult_WithParagraphs_SerializesNestedStructure()
    {
        var result = new TextResult
        {
            Success = true,
            ShapeId = 1,
            ShapeName = "Title 1",
            Text = "Hello World",
            Paragraphs =
            [
                new TextParagraphInfo
                {
                    Index = 0,
                    Text = "Hello World",
                    Runs =
                    [
                        new TextRunInfo { Text = "Hello ", Bold = true, FontSize = 24f },
                        new TextRunInfo { Text = "World", Italic = true }
                    ]
                }
            ]
        };
        var json = JsonSerializer.Serialize(result, JsonOptions);

        Assert.Contains("\"bold\":true", json);
        Assert.Contains("\"fontSize\":24", json);
        Assert.Contains("\"italic\":true", json);
    }

    [Fact]
    public void OperationResult_RoundTrip_PreservesAllFields()
    {
        var original = new OperationResult
        {
            Success = true,
            Action = "delete",
            Message = "Deleted slide 3",
            FilePath = @"C:\test\pres.pptx"
        };

        var json = JsonSerializer.Serialize(original, JsonOptions);
        var deserialized = JsonSerializer.Deserialize<OperationResult>(json, JsonOptions);

        Assert.NotNull(deserialized);
        Assert.Equal(original.Success, deserialized.Success);
        Assert.Equal(original.Action, deserialized.Action);
        Assert.Equal(original.Message, deserialized.Message);
        Assert.Equal(original.FilePath, deserialized.FilePath);
    }

    [Fact]
    public void DocumentPropertyResult_AllNulls_MinimalJson()
    {
        var result = new DocumentPropertyResult { Success = true };
        var json = JsonSerializer.Serialize(result, JsonOptions);

        // Should only have success, no null property fields
        Assert.Contains("\"success\":true", json);
        Assert.DoesNotContain("\"title\":", json);
        Assert.DoesNotContain("\"author\":", json);
        Assert.DoesNotContain("\"subject\":", json);
    }

    [Fact]
    public void HyperlinkInfo_ConditionalSerialization_WhenWritingDefault()
    {
        var info = new HyperlinkInfo
        {
            Index = 1,
            Address = "https://example.com"
        };
        var json = JsonSerializer.Serialize(info, JsonOptions);

        Assert.Contains("\"address\":\"https://example.com\"", json);
        // SlideIndex = 0 should be omitted (WhenWritingDefault)
        Assert.DoesNotContain("\"slideIndex\":", json);
    }
}
