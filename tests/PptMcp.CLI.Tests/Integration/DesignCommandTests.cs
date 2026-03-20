using PptMcp.CLI.Tests.Helpers;
using PptMcp.Core.Data;
using PptMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace PptMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "Design")]
[Trait("RequiresPowerPoint", "true")]
[Trait("Speed", "Fast")]
public sealed class DesignCommandTests
{
    private readonly ITestOutputHelper _output;

    public DesignCommandTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public async Task ListArchetypes_ReturnsUnifiedFamilies()
    {
        using var referenceCatalogScope = DesignCatalogProvider.UseReferenceCatalogRootForTesting(ReferenceCatalogFixture.GetCatalogRoot());

        var (sessionId, filePath) = await CreateSessionAsync();
        try
        {
            var (result, json) = await CliProcessHelper.RunJsonAsync($"design list-archetypes --session {sessionId}");
            _output.WriteLine(result.Stdout);

            Assert.Equal(0, result.ExitCode);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean());

            var archetypes = json.RootElement.GetProperty("archetypes");
            Assert.Contains(
                archetypes.EnumerateArray(),
                entry => entry.GetProperty("id").GetString() == "framework"
                    && entry.GetProperty("hasCuratedLayoutGuidance").GetBoolean());
            Assert.Contains(
                archetypes.EnumerateArray(),
                entry => entry.GetProperty("id").GetString() == "org-chart"
                    && !entry.GetProperty("hasCuratedLayoutGuidance").GetBoolean());
        }
        finally
        {
            await CloseSessionAsync(sessionId, filePath);
        }
    }

    [Fact]
    public async Task GetArchetype_ForCuratedFamily_ReturnsSanitizedObservedExamples()
    {
        using var referenceCatalogScope = DesignCatalogProvider.UseReferenceCatalogRootForTesting(ReferenceCatalogFixture.GetCatalogRoot());

        var (sessionId, filePath) = await CreateSessionAsync();
        try
        {
            var (result, json) = await CliProcessHelper.RunJsonAsync(
                $"design get-archetype --session {sessionId} --archetype-id \"framework\"");
            _output.WriteLine(result.Stdout);

            Assert.Equal(0, result.ExitCode);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean());
            Assert.True(json.RootElement.GetProperty("hasCuratedLayoutGuidance").GetBoolean());
            Assert.Equal(3, json.RootElement.GetProperty("observedSlideCount").GetInt32());

            var observedExamples = json.RootElement.GetProperty("observedExamples");
            Assert.Contains(
                observedExamples.EnumerateArray(),
                example => example.GetProperty("id").GetString() == ReferenceCatalogFixture.FrameworkMatrixReferenceId);
            Assert.All(
                observedExamples.EnumerateArray(),
                example => Assert.False(example.TryGetProperty("sourceName", out _)));

            var observedSubtypes = json.RootElement.GetProperty("observedSubtypes");
            Assert.Contains(
                observedSubtypes.EnumerateArray(),
                subtype =>
                {
                    if (subtype.GetProperty("subArchetypeId").GetString() != "matrix-grid")
                    {
                        return false;
                    }

                    var exampleDetails = subtype.GetProperty("exampleDetails");
                    return exampleDetails.EnumerateArray()
                        .Any(example => example.GetProperty("id").GetString() == ReferenceCatalogFixture.FrameworkMatrixReferenceId);
                });
        }
        finally
        {
            await CloseSessionAsync(sessionId, filePath);
        }
    }

    [Fact]
    public async Task GetArchetype_ForLearnedOnlyFamily_ReturnsUnifiedObservedData()
    {
        using var referenceCatalogScope = DesignCatalogProvider.UseReferenceCatalogRootForTesting(ReferenceCatalogFixture.GetCatalogRoot());

        var (sessionId, filePath) = await CreateSessionAsync();
        try
        {
            var (result, json) = await CliProcessHelper.RunJsonAsync(
                $"design get-archetype --session {sessionId} --archetype-id \"org-chart\"");
            _output.WriteLine(result.Stdout);

            Assert.Equal(0, result.ExitCode);
            Assert.True(json.RootElement.GetProperty("success").GetBoolean());
            Assert.False(json.RootElement.GetProperty("hasCuratedLayoutGuidance").GetBoolean());
            Assert.Equal(1, json.RootElement.GetProperty("observedSlideCount").GetInt32());

            var observedExamples = json.RootElement.GetProperty("observedExamples");
            var example = Assert.Single(observedExamples.EnumerateArray());
            Assert.Equal(ReferenceCatalogFixture.OrgChartReferenceId, example.GetProperty("id").GetString());

            var observedSubtypes = json.RootElement.GetProperty("observedSubtypes");
            Assert.Contains(
                observedSubtypes.EnumerateArray(),
                subtype => subtype.GetProperty("subArchetypeId").GetString() == "hierarchy-tree");
        }
        finally
        {
            await CloseSessionAsync(sessionId, filePath);
        }
    }

    private static async Task<(string SessionId, string FilePath)> CreateSessionAsync()
    {
        var filePath = Path.Join(Path.GetTempPath(), $"CliDesignCommandTests_{Guid.NewGuid():N}.pptx");
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
