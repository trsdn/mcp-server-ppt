using System.Reflection;
using PptMcp.Core.Commands.Design;
using PptMcp.Core.Data;
using PptMcp.Core.Tests.Helpers;
using Xunit;

namespace PptMcp.Core.Tests.Unit;

public class DesignReferenceCatalogTests
{
    [Fact]
    public void ReferenceCatalogProvider_LoadsManifestAndSubtypeCatalog()
    {
        using var _ = DesignCatalogProvider.UseReferenceCatalogRootForTesting(ReferenceCatalogFixture.GetCatalogRoot());

        var manifest = DesignCatalogProvider.GetReferenceManifest();
        var catalog = DesignCatalogProvider.GetReferenceSubArchetypeCatalog();

        Assert.Equal(6, manifest.Count);
        Assert.Equal(6, catalog.AnnotatedEntries);
        Assert.Equal(4, catalog.DistinctSubArchetypes);
        Assert.Empty(catalog.UnsplitArchetypes);
        Assert.All(manifest, entry => Assert.False(string.IsNullOrWhiteSpace(entry.SubArchetypeId)));
    }

    [Fact]
    public void ListArchetypes_UnifiesCuratedAndLearnedFamilies()
    {
        using var _ = DesignCatalogProvider.UseReferenceCatalogRootForTesting(ReferenceCatalogFixture.GetCatalogRoot());

        var result = new DesignCommands().ListArchetypes(batch: null!);

        Assert.True(result.Success);

        var framework = Assert.Single(result.Archetypes, item => item.Id == "framework");
        Assert.True(framework.HasCuratedLayoutGuidance);
        Assert.Equal(3, framework.ObservedSlideCount);
        Assert.Equal(2, framework.ObservedSubtypeCount);
        Assert.Contains(ReferenceCatalogFixture.FrameworkMatrixReferenceId, framework.ObservedExampleSlides);

        var orgChart = Assert.Single(result.Archetypes, item => item.Id == "org-chart");
        Assert.False(orgChart.HasCuratedLayoutGuidance);
        Assert.Equal(1, orgChart.ObservedSlideCount);
        Assert.Equal(1, orgChart.ObservedSubtypeCount);
        Assert.Contains(ReferenceCatalogFixture.OrgChartReferenceId, orgChart.ObservedExampleSlides);
    }

    [Fact]
    public void GetArchetype_ForCuratedFamily_ContainsSanitizedObservedCoverage()
    {
        using var _ = DesignCatalogProvider.UseReferenceCatalogRootForTesting(ReferenceCatalogFixture.GetCatalogRoot());

        var result = new DesignCommands().GetArchetype(batch: null!, archetypeId: "framework");

        Assert.True(result.Success);
        Assert.True(result.HasCuratedLayoutGuidance);
        Assert.Equal(3, result.ObservedSlideCount);
        Assert.Equal(3, result.ObservedExamples.Count);
        Assert.Contains(result.ObservedExamples, example => example.Id == ReferenceCatalogFixture.FrameworkMatrixReferenceId);
        Assert.All(result.ObservedExamples, example =>
        {
            Assert.StartsWith("ref-", example.Id);
            Assert.False(string.IsNullOrWhiteSpace(example.Rationale));
        });

        var matrixSubtype = Assert.Single(result.ObservedSubtypes, subtype => subtype.SubArchetypeId == "matrix-grid");
        Assert.Contains(ReferenceCatalogFixture.FrameworkMatrixReferenceId, matrixSubtype.ExampleSlides);
        Assert.All(matrixSubtype.ExampleSlides, exampleId => Assert.StartsWith("ref-", exampleId));
        Assert.Contains(matrixSubtype.ExampleDetails, example => example.Id == ReferenceCatalogFixture.FrameworkMatrixReferenceId);
        Assert.All(matrixSubtype.ExampleDetails, example => Assert.Equal("framework", example.ArchetypeId));

        var pillarsSubtype = Assert.Single(result.ObservedSubtypes, subtype => subtype.SubArchetypeId == "pillars-model");
        Assert.All(pillarsSubtype.ExampleSlides, exampleId => Assert.StartsWith("ref-", exampleId));

        var auditSample = Assert.Single(result.AuditSamples);
        Assert.Equal(ReferenceCatalogFixture.FrameworkAuditReferenceId, auditSample.ReferenceId);
        Assert.Equal("framework", auditSample.CurrentArchetypeId);
        Assert.Equal("comparison", auditSample.SuggestedArchetypeId);
        Assert.NotEmpty(result.Detail);
    }

    [Fact]
    public void GetArchetype_ForLearnedOnlyFamily_SucceedsWithObservedExamples()
    {
        using var _ = DesignCatalogProvider.UseReferenceCatalogRootForTesting(ReferenceCatalogFixture.GetCatalogRoot());

        var result = new DesignCommands().GetArchetype(batch: null!, archetypeId: "org-chart");

        Assert.True(result.Success);
        Assert.Equal("org-chart", result.Id);
        Assert.False(result.HasCuratedLayoutGuidance);
        Assert.Equal(1, result.ObservedSlideCount);
        Assert.Single(result.ObservedExamples);
        Assert.Equal(ReferenceCatalogFixture.OrgChartReferenceId, result.ObservedExamples[0].Id);

        var hierarchySubtype = Assert.Single(result.ObservedSubtypes, subtype => subtype.SubArchetypeId == "hierarchy-tree");
        Assert.Single(hierarchySubtype.ExampleSlides);
        Assert.Equal(ReferenceCatalogFixture.OrgChartReferenceId, hierarchySubtype.ExampleSlides[0]);
        Assert.Single(hierarchySubtype.ExampleDetails);
        Assert.Equal(ReferenceCatalogFixture.OrgChartReferenceId, hierarchySubtype.ExampleDetails[0].Id);

        Assert.Contains("No curated layout-coordinate guide", result.Detail);
    }

    [Fact]
    public void ReferenceCatalogProvider_UsesEvalAssetRepoRootOverride()
    {
        var tempRepoRoot = Path.Combine(Path.GetTempPath(), $"PptMcpEvalAssets_{Guid.NewGuid():N}");
        var stagedCatalogRoot = Path.Combine(tempRepoRoot, "eval", "data", "archetype-references");
        Directory.CreateDirectory(stagedCatalogRoot);

        foreach (var fileName in new[] { "manifest.json", "sub-archetypes.json", "new-archetypes.json" })
        {
            File.Copy(
                Path.Combine(ReferenceCatalogFixture.GetCatalogRoot(), fileName),
                Path.Combine(stagedCatalogRoot, fileName));
        }

        var previousAssetRepoRoot = Environment.GetEnvironmentVariable("PPTMCP_EVAL_ASSET_REPO_ROOT");
        var previousReferenceRoot = Environment.GetEnvironmentVariable("PPTMCP_REFERENCE_DATA_ROOT");

        try
        {
            Environment.SetEnvironmentVariable("PPTMCP_REFERENCE_DATA_ROOT", null);
            Environment.SetEnvironmentVariable("PPTMCP_EVAL_ASSET_REPO_ROOT", tempRepoRoot);
            ResetReferenceCatalogCache();

            Assert.True(DesignCatalogProvider.TryGetReferenceCatalogAvailability(out var errorMessage), errorMessage);

            var manifest = DesignCatalogProvider.GetReferenceManifest();
            Assert.Equal(6, manifest.Count);
            Assert.Contains(manifest, entry => entry.Id == ReferenceCatalogFixture.FrameworkMatrixRawId);
        }
        finally
        {
            Environment.SetEnvironmentVariable("PPTMCP_EVAL_ASSET_REPO_ROOT", previousAssetRepoRoot);
            Environment.SetEnvironmentVariable("PPTMCP_REFERENCE_DATA_ROOT", previousReferenceRoot);
            ResetReferenceCatalogCache();

            if (Directory.Exists(tempRepoRoot))
            {
                Directory.Delete(tempRepoRoot, recursive: true);
            }
        }
    }

    private static void ResetReferenceCatalogCache()
    {
        var resetMethod = typeof(DesignCatalogProvider).GetMethod(
            "ResetReferenceCatalogCache",
            BindingFlags.NonPublic | BindingFlags.Static);

        Assert.NotNull(resetMethod);
        resetMethod.Invoke(null, null);
    }
}
