using System.Text.Json;
using PptMcp.Generated;
using PptMcp.Core.Data;
using PptMcp.Core.Tests.Helpers;
using PptMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

namespace PptMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Layer", "MCP")]
[Trait("Category", "Integration")]
[Trait("Feature", "Design")]
[Trait("RequiresPowerPoint", "true")]
public sealed class PptDesignToolTests(ITestOutputHelper output)
{
    [Fact]
    public void ListArchetypes_WithSession_ReturnsUnifiedCatalog()
    {
        using var referenceCatalogScope = DesignCatalogProvider.UseReferenceCatalogRootForTesting(ReferenceCatalogFixture.GetCatalogRoot());

        var tempPath = Path.Join(Path.GetTempPath(), $"PptDesignToolTest_{Guid.NewGuid():N}.pptx");
        string? sessionId = null;

        try
        {
            sessionId = CreateSession(tempPath);

            var result = PptDesignTool.PptDesign(
                action: DesignAction.ListArchetypes,
                session_id: sessionId,
                theme_path: null,
                design_index: 0,
                archetype_id: null,
                palette_id: null,
                profile_id: null,
                grid_id: null,
                density_id: null,
                sequence_id: null);

            output.WriteLine(result);

            var json = JsonDocument.Parse(result).RootElement;
            Assert.True(json.GetProperty("success").GetBoolean());
            Assert.Contains(
                json.GetProperty("archetypes").EnumerateArray(),
                entry => entry.GetProperty("id").GetString() == "org-chart");
        }
        finally
        {
            CloseSession(sessionId);
            DeleteFile(tempPath);
        }
    }

    [Fact]
    public void GetArchetype_WithSession_ReturnsSanitizedObservedExamples()
    {
        using var referenceCatalogScope = DesignCatalogProvider.UseReferenceCatalogRootForTesting(ReferenceCatalogFixture.GetCatalogRoot());

        var tempPath = Path.Join(Path.GetTempPath(), $"PptDesignToolTest_{Guid.NewGuid():N}.pptx");
        string? sessionId = null;

        try
        {
            sessionId = CreateSession(tempPath);

            var result = PptDesignTool.PptDesign(
                action: DesignAction.GetArchetype,
                session_id: sessionId,
                theme_path: null,
                design_index: 0,
                archetype_id: "framework",
                palette_id: null,
                profile_id: null,
                grid_id: null,
                density_id: null,
                sequence_id: null);

            output.WriteLine(result);

            var json = JsonDocument.Parse(result).RootElement;
            Assert.True(json.GetProperty("success").GetBoolean());
            Assert.True(json.GetProperty("hasCuratedLayoutGuidance").GetBoolean());
            Assert.Contains(
                json.GetProperty("observedExamples").EnumerateArray(),
                example => example.GetProperty("id").GetString() == ReferenceCatalogFixture.FrameworkMatrixReferenceId);

            var matrixSubtype = Assert.Single(
                json.GetProperty("observedSubtypes").EnumerateArray(),
                subtype => subtype.GetProperty("subArchetypeId").GetString() == "matrix-grid");
            Assert.Contains(
                matrixSubtype.GetProperty("exampleDetails").EnumerateArray(),
                example => example.GetProperty("id").GetString() == ReferenceCatalogFixture.FrameworkMatrixReferenceId);
        }
        finally
        {
            CloseSession(sessionId);
            DeleteFile(tempPath);
        }
    }

    [Fact]
    public void GetArchetype_WithSession_ReturnsUnifiedObservedDataForLearnedOnlyFamily()
    {
        using var referenceCatalogScope = DesignCatalogProvider.UseReferenceCatalogRootForTesting(ReferenceCatalogFixture.GetCatalogRoot());

        var tempPath = Path.Join(Path.GetTempPath(), $"PptDesignToolTest_{Guid.NewGuid():N}.pptx");
        string? sessionId = null;

        try
        {
            sessionId = CreateSession(tempPath);

            var result = PptDesignTool.PptDesign(
                action: DesignAction.GetArchetype,
                session_id: sessionId,
                theme_path: null,
                design_index: 0,
                archetype_id: "org-chart",
                palette_id: null,
                profile_id: null,
                grid_id: null,
                density_id: null,
                sequence_id: null);

            output.WriteLine(result);

            var json = JsonDocument.Parse(result).RootElement;
            Assert.True(json.GetProperty("success").GetBoolean());
            Assert.False(json.GetProperty("hasCuratedLayoutGuidance").GetBoolean());
            Assert.Equal(1, json.GetProperty("observedSlideCount").GetInt32());
            Assert.Equal(
                ReferenceCatalogFixture.OrgChartReferenceId,
                Assert.Single(json.GetProperty("observedExamples").EnumerateArray()).GetProperty("id").GetString());
        }
        finally
        {
            CloseSession(sessionId);
            DeleteFile(tempPath);
        }
    }

    private static string CreateSession(string path)
    {
        var result = PptFileTool.PptFile(
            PptFileAction.Create,
            path: path,
            session_id: null,
            save: false,
            show: false,
            timeout_seconds: 300);

        var json = JsonDocument.Parse(result).RootElement;
        Assert.True(json.GetProperty("success").GetBoolean());
        return json.GetProperty("session_id").GetString()!;
    }

    private static void CloseSession(string? sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return;
        }

        PptFileTool.PptFile(
            PptFileAction.Close,
            path: null,
            session_id: sessionId,
            save: false,
            show: false,
            timeout_seconds: 300);
    }

    private static void DeleteFile(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }
}
