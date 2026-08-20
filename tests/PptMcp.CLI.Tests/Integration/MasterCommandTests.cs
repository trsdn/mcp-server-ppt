using PptMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace PptMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "Master")]
[Trait("RequiresPowerPoint", "true")]
[Trait("Speed", "Fast")]
public sealed class MasterCommandTests
{
    private readonly ITestOutputHelper _output;

    public MasterCommandTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public async Task List_ReturnsMastersWithLayouts()
    {
        var (sessionId, filePath) = await CreateSessionAsync();
        try
        {
            var (result, json) = await CliProcessHelper.RunJsonAsync($"master list --session {sessionId}");
            _output.WriteLine(result.Stdout);

            Assert.Equal(0, result.ExitCode);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean());

            var masters = json.RootElement.GetProperty("masters");
            Assert.NotEmpty(masters.EnumerateArray());

            var first = masters.EnumerateArray().First();
            Assert.False(string.IsNullOrWhiteSpace(first.GetProperty("name").GetString()));
            Assert.NotEmpty(first.GetProperty("layouts").EnumerateArray());
        }
        finally
        {
            await CloseSessionAsync(sessionId, filePath);
        }
    }

    [Fact]
    public async Task ListLayouts_ForFirstMaster_ReturnsLayouts()
    {
        var (sessionId, filePath) = await CreateSessionAsync();
        try
        {
            var (result, json) = await CliProcessHelper.RunJsonAsync(
                $"master list-layouts --session {sessionId} --master-index 1");
            _output.WriteLine(result.Stdout);

            Assert.Equal(0, result.ExitCode);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean());
            Assert.Contains("layout", json.RootElement.GetProperty("message").GetString()!, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            await CloseSessionAsync(sessionId, filePath);
        }
    }

    [Fact]
    public async Task ListShapes_ForFirstMaster_Succeeds()
    {
        var (sessionId, filePath) = await CreateSessionAsync();
        try
        {
            var (result, json) = await CliProcessHelper.RunJsonAsync(
                $"master list-shapes --session {sessionId} --master-index 1");
            _output.WriteLine(result.Stdout);

            Assert.Equal(0, result.ExitCode);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        }
        finally
        {
            await CloseSessionAsync(sessionId, filePath);
        }
    }

    private static async Task<(string SessionId, string FilePath)> CreateSessionAsync()
    {
        var filePath = Path.Join(Path.GetTempPath(), $"CliMasterCommandTests_{Guid.NewGuid():N}.pptx");
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
