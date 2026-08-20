using PptMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace PptMcp.CLI.Tests.Integration;

/// <summary>
/// Layout names reported by PowerPoint are localized ("Blank" is "Leer" on a German
/// install), so looking a layout up by its English name has to work regardless of the
/// UI language of the machine running the tests.
/// </summary>
[Collection("Service")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "Slide")]
[Trait("RequiresPowerPoint", "true")]
[Trait("Speed", "Fast")]
public sealed class SlideLayoutLookupTests
{
    private readonly ITestOutputHelper _output;

    public SlideLayoutLookupTests(ITestOutputHelper output) => _output = output;

    [Theory]
    [InlineData("Blank")]
    [InlineData("Title Slide")]
    [InlineData("Title and Content")]
    [InlineData("Two Content")]
    public async Task Create_WithCanonicalEnglishLayoutName_Succeeds(string layoutName)
    {
        var (sessionId, filePath) = await CreateSessionAsync();
        try
        {
            var (result, json) = await CliProcessHelper.RunJsonAsync(
                $"slide create --session {sessionId} --position 0 --layout-name \"{layoutName}\"");
            _output.WriteLine(result.Stdout);

            Assert.Equal(0, result.ExitCode);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean());

            var (listResult, listJson) = await CliProcessHelper.RunJsonAsync($"slide list --session {sessionId}");
            _output.WriteLine(listResult.Stdout);

            Assert.True(listJson.RootElement.GetProperty("success").GetBoolean());
            Assert.NotEmpty(listJson.RootElement.GetProperty("slides").EnumerateArray());
        }
        finally
        {
            await CloseSessionAsync(sessionId, filePath);
        }
    }

    [Fact]
    public async Task Create_WithLocalizedLayoutName_StillSucceeds()
    {
        var (sessionId, filePath) = await CreateSessionAsync();
        try
        {
            var (_, layoutsJson) = await CliProcessHelper.RunJsonAsync(
                $"master list --session {sessionId}");

            var firstLayoutName = layoutsJson.RootElement
                .GetProperty("masters").EnumerateArray().First()
                .GetProperty("layouts").EnumerateArray().First()
                .GetProperty("name").GetString()!;

            _output.WriteLine($"Using layout reported by PowerPoint: {firstLayoutName}");

            var (result, json) = await CliProcessHelper.RunJsonAsync(
                $"slide create --session {sessionId} --position 0 --layout-name \"{firstLayoutName}\"");
            _output.WriteLine(result.Stdout);

            Assert.Equal(0, result.ExitCode);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        }
        finally
        {
            await CloseSessionAsync(sessionId, filePath);
        }
    }

    [Fact]
    public async Task Create_WithUnknownLayoutName_ListsAvailableLayouts()
    {
        var (sessionId, filePath) = await CreateSessionAsync();
        try
        {
            var (result, json) = await CliProcessHelper.RunJsonAsync(
                $"slide create --session {sessionId} --position 0 --layout-name \"DefinitelyNotALayout\"");
            _output.WriteLine(result.Stdout);

            Assert.False(json.RootElement.GetProperty("success").GetBoolean());

            var error = json.RootElement.GetProperty("error").GetString()!;
            Assert.Contains("DefinitelyNotALayout", error, StringComparison.Ordinal);
            Assert.Contains("Available", error, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            await CloseSessionAsync(sessionId, filePath);
        }
    }

    private static async Task<(string SessionId, string FilePath)> CreateSessionAsync()
    {
        var filePath = Path.Join(Path.GetTempPath(), $"CliSlideLayoutTests_{Guid.NewGuid():N}.pptx");
        var (result, json) = await CliProcessHelper.RunJsonAsync($"session create \"{filePath}\"");

        Assert.Equal(0, result.ExitCode);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());

        return (json.RootElement.GetProperty("sessionId").GetString()!, filePath);
    }

    private static async Task CloseSessionAsync(string? sessionId, string filePath)
    {
        if (!string.IsNullOrWhiteSpace(sessionId))
        {
            await CliProcessHelper.RunAsync($"session close --session {sessionId} --save false");
        }

        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }
}
