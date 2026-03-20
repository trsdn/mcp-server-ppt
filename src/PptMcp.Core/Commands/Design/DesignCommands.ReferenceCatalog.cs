using PptMcp.Core.Data;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Design;

public partial class DesignCommands
{
    private static ReferenceSlideInfo ToReferenceSlideInfo(ReferenceManifestEntry entry)
    {
        return new ReferenceSlideInfo
        {
            Id = DesignCatalogProvider.GetPublicReferenceId(entry),
            ArchetypeId = entry.ArchetypeId,
            SubArchetypeId = entry.SubArchetypeId,
            Rationale = entry.Rationale
        };
    }

    private static (string Name, string Summary) GetLearnedArchetypeMetadata(string archetypeId)
    {
        var coreEntry = DesignCatalogProvider.GetArchetypes()
            .Find(entry => string.Equals(entry.Id, archetypeId, StringComparison.OrdinalIgnoreCase));
        if (coreEntry != null)
        {
            return (coreEntry.Name, coreEntry.When);
        }

        var learnedEntry = DesignCatalogProvider.GetNewArchetype(archetypeId);
        if (learnedEntry != null)
        {
            return (HumanizeIdentifier(archetypeId), learnedEntry.Description);
        }

        var humanized = HumanizeIdentifier(archetypeId);
        return (humanized, $"Learned reference family from the curated slide corpus: {humanized}.");
    }

    private static Dictionary<string, int> GetCountsByArchetype(List<ReferenceManifestEntry> manifest)
    {
        return manifest
            .GroupBy(entry => entry.ArchetypeId, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);
    }

    private static ReferenceTopLevelEntry? GetReferenceTopLevel(ReferenceSubArchetypeCatalog catalog, string archetypeId)
    {
        return catalog.TopLevels.Find(entry => string.Equals(entry.ArchetypeId, archetypeId, StringComparison.OrdinalIgnoreCase));
    }

    private static List<ReferenceSubtypeInfo> BuildSubtypeInfos(
        ReferenceTopLevelEntry topLevel,
        List<ReferenceManifestEntry> manifest)
    {
        var manifestBySourceName = manifest.ToDictionary(entry => entry.SourceName, StringComparer.OrdinalIgnoreCase);

        return
        [
            .. topLevel.Subtypes
                .OrderByDescending(subtype => subtype.Count)
                .ThenBy(subtype => subtype.SubArchetypeId, StringComparer.OrdinalIgnoreCase)
                .Select(subtype =>
                {
                    var exampleDetails = BuildObservedExamples(manifestBySourceName, subtype.ExampleSlides, int.MaxValue);
                    return new ReferenceSubtypeInfo
                    {
                        SubArchetypeId = subtype.SubArchetypeId,
                        Description = subtype.Description,
                        HeuristicPhrases = subtype.HeuristicPhrases,
                        Count = subtype.Count,
                        ExampleSlides = [.. exampleDetails.Select(example => example.Id)],
                        ExampleDetails = exampleDetails
                    };
                })
        ];
    }

    private static List<ReferenceMisbucketedSampleInfo> GetMisbucketedSamples(
        ReferenceSubArchetypeCatalog catalog,
        string archetypeId)
    {
        return
        [
            .. catalog.MisbucketedSamples
                .Where(sample => string.Equals(sample.CurrentArchetypeId, archetypeId, StringComparison.OrdinalIgnoreCase))
                .Select(sample =>
                {
                    var referenceId = DesignCatalogProvider.TryGetPublicReferenceIdFromSourceName(sample.SourceName);
                    if (string.IsNullOrWhiteSpace(referenceId))
                    {
                        return null;
                    }

                    return new ReferenceMisbucketedSampleInfo
                    {
                        ReferenceId = referenceId,
                        CurrentArchetypeId = sample.CurrentArchetypeId,
                        SuggestedArchetypeId = sample.SuggestedArchetypeId,
                        Reason = sample.Reason
                    };
                })
                .OfType<ReferenceMisbucketedSampleInfo>()
        ];
    }

    private static List<string> GetExampleSlides(ReferenceTopLevelEntry topLevel, int maxCount)
    {
        return
        [
            .. topLevel.Subtypes
                .SelectMany(subtype => subtype.ExampleSlides)
                .Select(example => DesignCatalogProvider.TryGetPublicReferenceIdFromSourceName(example))
                .OfType<string>()
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(maxCount)
        ];
    }

    private static List<ReferenceSlideInfo> GetObservedExamples(
        List<ReferenceManifestEntry> manifest,
        ReferenceTopLevelEntry topLevel,
        int maxCount)
    {
        var manifestBySourceName = manifest.ToDictionary(entry => entry.SourceName, StringComparer.OrdinalIgnoreCase);
        return BuildObservedExamples(
            manifestBySourceName,
            topLevel.Subtypes.SelectMany(subtype => subtype.ExampleSlides),
            maxCount);
    }

    private static List<ReferenceSlideInfo> BuildObservedExamples(
        IReadOnlyDictionary<string, ReferenceManifestEntry> manifestBySourceName,
        IEnumerable<string> exampleSourceNames,
        int maxCount)
    {
        return
        [
            .. exampleSourceNames
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(sourceName => manifestBySourceName.GetValueOrDefault(sourceName))
                .OfType<ReferenceManifestEntry>()
                .Select(ToReferenceSlideInfo)
                .Take(maxCount)
        ];
    }

    private static bool TryLoadReferenceCatalog(
        out List<ReferenceManifestEntry> manifest,
        out ReferenceSubArchetypeCatalog catalog,
        out string? errorMessage)
    {
        manifest = [];
        catalog = new ReferenceSubArchetypeCatalog();

        if (!DesignCatalogProvider.TryGetReferenceCatalogAvailability(out errorMessage))
        {
            return false;
        }

        manifest = DesignCatalogProvider.GetReferenceManifest();
        catalog = DesignCatalogProvider.GetReferenceSubArchetypeCatalog();
        errorMessage = null;
        return true;
    }

    private static string HumanizeIdentifier(string identifier)
    {
        return string.Join(
            ' ',
            identifier
                .Split('-', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(segment => char.ToUpperInvariant(segment[0]) + segment[1..]));
    }

    private static string BuildUnifiedArchetypeDetail(string archetypeId, string? curatedDetail)
    {
        if (!string.IsNullOrWhiteSpace(curatedDetail))
        {
            return curatedDetail;
        }

        var learnedEntry = DesignCatalogProvider.GetNewArchetype(archetypeId);
        if (learnedEntry != null)
        {
            return $"{learnedEntry.Description}\n\nNo curated layout-coordinate guide is available yet for this archetype. Use the observed subtypes and observed examples to infer the composition pattern.";
        }

        return $"Detail not available for archetype '{archetypeId}'.";
    }
}
