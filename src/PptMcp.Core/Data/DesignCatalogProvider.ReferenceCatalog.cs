using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text;

namespace PptMcp.Core.Data;

internal static partial class DesignCatalogProvider
{
    private const string ReferenceCatalogRootEnvironmentVariable = "PPTMCP_REFERENCE_DATA_ROOT";
    private const string EvalAssetRepoRootEnvironmentVariable = "PPTMCP_EVAL_ASSET_REPO_ROOT";
    private static readonly string[] ReferenceCatalogFiles =
    [
        "manifest.json",
        "sub-archetypes.json",
        "new-archetypes.json"
    ];
    private static readonly Lock ReferenceCatalogLock = new();

    private static string? _referenceCatalogRootOverride;
    private static string? _resolvedReferenceCatalogRoot;
    private static List<ReferenceManifestEntry>? _referenceManifest;
    private static ReferenceSubArchetypeCatalog? _referenceSubArchetypeCatalog;
    private static List<NewArchetypeEntry>? _newArchetypes;
    private static Dictionary<string, ReferenceManifestEntry>? _referenceManifestByPublicId;
    private static Dictionary<string, string>? _publicIdBySourceName;
    private static readonly ConcurrentDictionary<string, NewArchetypeEntry?> _newArchetypeLookup = new(StringComparer.OrdinalIgnoreCase);

    public static bool TryGetReferenceCatalogAvailability(out string? errorMessage)
    {
        var explicitRoot = GetExplicitReferenceCatalogRoot();

        if (!string.IsNullOrWhiteSpace(explicitRoot))
        {
            var normalizedRoot = Path.GetFullPath(explicitRoot);
            if (HasReferenceCatalogFiles(normalizedRoot))
            {
                _resolvedReferenceCatalogRoot = normalizedRoot;
                errorMessage = null;
                return true;
            }

            errorMessage = $"Reference catalog not found at '{normalizedRoot}'.";
            return false;
        }

        var explicitEvalAssetRepoRoot = GetExplicitEvalAssetRepoRoot();
        if (!string.IsNullOrWhiteSpace(explicitEvalAssetRepoRoot))
        {
            var derivedRoot = GetReferenceCatalogRootFromEvalAssetRepoRoot(explicitEvalAssetRepoRoot);
            if (HasReferenceCatalogFiles(derivedRoot))
            {
                _resolvedReferenceCatalogRoot = derivedRoot;
                errorMessage = null;
                return true;
            }

            errorMessage = $"Reference catalog not found at '{derivedRoot}'.";
            return false;
        }

        if (ResolveReferenceCatalogRoot() is not null)
        {
            errorMessage = null;
            return true;
        }

        errorMessage =
            $"Local reference catalog not available. Store raw reference data under 'eval\\data\\archetype-references', set {EvalAssetRepoRootEnvironmentVariable} to a repo containing that path, or set {ReferenceCatalogRootEnvironmentVariable} directly.";
        return false;
    }

    public static List<ReferenceManifestEntry> GetReferenceManifest()
    {
        _referenceManifest ??= LoadReferenceJson<List<ReferenceManifestEntry>>("manifest.json");
        return _referenceManifest;
    }

    public static ReferenceSubArchetypeCatalog GetReferenceSubArchetypeCatalog()
    {
        _referenceSubArchetypeCatalog ??= LoadReferenceJson<ReferenceSubArchetypeCatalog>("sub-archetypes.json");
        return _referenceSubArchetypeCatalog;
    }

    public static List<NewArchetypeEntry> GetNewArchetypes()
    {
        _newArchetypes ??= LoadReferenceJson<List<NewArchetypeEntry>>("new-archetypes.json");
        return _newArchetypes;
    }

    public static NewArchetypeEntry? GetNewArchetype(string archetypeId)
    {
        return _newArchetypeLookup.GetOrAdd(
            archetypeId,
            id => GetNewArchetypes().Find(entry => string.Equals(entry.Id, id, StringComparison.OrdinalIgnoreCase)));
    }

    public static string GetPublicReferenceId(ReferenceManifestEntry entry) => GetPublicReferenceId(entry.Id);

    public static string GetPublicReferenceId(string rawReferenceId)
    {
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(rawReferenceId));
        return $"ref-{Convert.ToHexString(hash.AsSpan(0, 6)).ToLowerInvariant()}";
    }

    public static ReferenceManifestEntry? GetReferenceManifestEntryByPublicId(string referenceId)
    {
        _referenceManifestByPublicId ??= GetReferenceManifest()
            .ToDictionary(GetPublicReferenceId, entry => entry, StringComparer.OrdinalIgnoreCase);

        return _referenceManifestByPublicId.GetValueOrDefault(referenceId);
    }

    public static string? TryGetPublicReferenceIdFromSourceName(string sourceName)
    {
        _publicIdBySourceName ??= GetReferenceManifest()
            .ToDictionary(entry => entry.SourceName, GetPublicReferenceId, StringComparer.OrdinalIgnoreCase);

        return _publicIdBySourceName.GetValueOrDefault(sourceName);
    }

    internal static IDisposable UseReferenceCatalogRootForTesting(string rootPath)
    {
        lock (ReferenceCatalogLock)
        {
            var previousRoot = _referenceCatalogRootOverride;
            _referenceCatalogRootOverride = Path.GetFullPath(rootPath);
            ResetReferenceCatalogCache();
            return new ReferenceCatalogRootScope(previousRoot);
        }
    }

    private static T LoadReferenceJson<T>(string fileName) where T : class
    {
        var root = GetReferenceCatalogRootOrThrow();
        var path = Path.Combine(root, fileName);
        using var stream = File.OpenRead(path);
        return JsonSerializer.Deserialize<T>(stream, JsonOptions)
            ?? throw new InvalidOperationException($"Failed to deserialize reference catalog file: {path}");
    }

    private static string GetReferenceCatalogRootOrThrow()
    {
        if (!TryGetReferenceCatalogAvailability(out var errorMessage))
        {
            throw new InvalidOperationException(errorMessage);
        }

        return ResolveReferenceCatalogRoot()
            ?? throw new InvalidOperationException(
                $"Reference catalog root resolution failed. Set {EvalAssetRepoRootEnvironmentVariable}, set {ReferenceCatalogRootEnvironmentVariable}, or create eval\\data\\archetype-references.");
    }

    private static string? ResolveReferenceCatalogRoot()
    {
        var explicitRoot = GetExplicitReferenceCatalogRoot();
        if (!string.IsNullOrWhiteSpace(explicitRoot))
        {
            var normalizedRoot = Path.GetFullPath(explicitRoot);
            _resolvedReferenceCatalogRoot = HasReferenceCatalogFiles(normalizedRoot) ? normalizedRoot : null;
            return _resolvedReferenceCatalogRoot;
        }

        var explicitEvalAssetRepoRoot = GetExplicitEvalAssetRepoRoot();
        if (!string.IsNullOrWhiteSpace(explicitEvalAssetRepoRoot))
        {
            var derivedRoot = GetReferenceCatalogRootFromEvalAssetRepoRoot(explicitEvalAssetRepoRoot);
            _resolvedReferenceCatalogRoot = HasReferenceCatalogFiles(derivedRoot) ? derivedRoot : null;
            return _resolvedReferenceCatalogRoot;
        }

        if (!string.IsNullOrWhiteSpace(_resolvedReferenceCatalogRoot) && HasReferenceCatalogFiles(_resolvedReferenceCatalogRoot))
        {
            return _resolvedReferenceCatalogRoot;
        }

        foreach (var basePath in EnumerateReferenceCatalogSearchRoots())
        {
            foreach (var candidateRoot in GetReferenceCatalogCandidateRoots(basePath))
            {
                if (HasReferenceCatalogFiles(candidateRoot))
                {
                    _resolvedReferenceCatalogRoot = candidateRoot;
                    return candidateRoot;
                }
            }
        }

        _resolvedReferenceCatalogRoot = null;
        return null;
    }

    private static string? GetExplicitReferenceCatalogRoot()
    {
        return string.IsNullOrWhiteSpace(_referenceCatalogRootOverride)
            ? Environment.GetEnvironmentVariable(ReferenceCatalogRootEnvironmentVariable)
            : _referenceCatalogRootOverride;
    }

    private static string? GetExplicitEvalAssetRepoRoot() => Environment.GetEnvironmentVariable(EvalAssetRepoRootEnvironmentVariable);

    private static string GetReferenceCatalogRootFromEvalAssetRepoRoot(string repoRoot)
    {
        return Path.Combine(Path.GetFullPath(repoRoot), "eval", "data", "archetype-references");
    }

    private static IEnumerable<string> EnumerateReferenceCatalogSearchRoots()
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var start in new[] { AppContext.BaseDirectory, Directory.GetCurrentDirectory() })
        {
            var current = new DirectoryInfo(start);
            while (current is not null)
            {
                if (seen.Add(current.FullName))
                {
                    yield return current.FullName;
                }

                current = current.Parent;
            }
        }
    }

    private static IEnumerable<string> GetReferenceCatalogCandidateRoots(string basePath)
    {
        yield return Path.Combine(basePath, "eval", "data", "archetype-references");
        yield return Path.Combine(basePath, "local-data", "archetype-references");
    }

    private static bool HasReferenceCatalogFiles(string candidateRoot)
    {
        return Directory.Exists(candidateRoot)
            && ReferenceCatalogFiles.All(fileName => File.Exists(Path.Combine(candidateRoot, fileName)));
    }

    private static void ResetReferenceCatalogCache()
    {
        _resolvedReferenceCatalogRoot = null;
        _referenceManifest = null;
        _referenceSubArchetypeCatalog = null;
        _newArchetypes = null;
        _referenceManifestByPublicId = null;
        _publicIdBySourceName = null;
        _newArchetypeLookup.Clear();
    }

    private sealed class ReferenceCatalogRootScope(string? previousRoot) : IDisposable
    {
        private readonly string? _previousRoot = previousRoot;
        private bool _disposed;

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            lock (ReferenceCatalogLock)
            {
                _referenceCatalogRootOverride = _previousRoot;
                ResetReferenceCatalogCache();
            }

            _disposed = true;
        }
    }
}

internal sealed class ReferenceManifestEntry
{
    public string Id { get; set; } = string.Empty;
    public string SourceName { get; set; } = string.Empty;
    public string SourceImage { get; set; } = string.Empty;
    public string ArchetypeId { get; set; } = string.Empty;
    public string Rationale { get; set; } = string.Empty;
    public string BatchId { get; set; } = string.Empty;
    public string DeckKey { get; set; } = string.Empty;
    public int SlideNumber { get; set; }
    public string? SubArchetypeId { get; set; }
}

internal sealed class ReferenceSubArchetypeCatalog
{
    public List<ReferenceTopLevelEntry> TopLevels { get; set; } = [];
    public List<ReferenceMisbucketedSample> MisbucketedSamples { get; set; } = [];
    public List<string> UnsplitArchetypes { get; set; } = [];
    public int AnnotatedEntries { get; set; }
    public int DistinctSubArchetypes { get; set; }
}

internal sealed class ReferenceTopLevelEntry
{
    public string ArchetypeId { get; set; } = string.Empty;
    public List<ReferenceSubArchetypeEntry> Subtypes { get; set; } = [];
}

internal sealed class ReferenceSubArchetypeEntry
{
    public string SubArchetypeId { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public List<string> HeuristicPhrases { get; set; } = [];
    public int Count { get; set; }
    public List<string> ExampleSlides { get; set; } = [];
}

internal sealed class ReferenceMisbucketedSample
{
    public string SourceName { get; set; } = string.Empty;
    public string CurrentArchetypeId { get; set; } = string.Empty;
    public string SuggestedArchetypeId { get; set; } = string.Empty;
    public string Reason { get; set; } = string.Empty;
}

internal sealed class NewArchetypeEntry
{
    public string Id { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public int Count { get; set; }
    public List<string> ExampleSlides { get; set; } = [];
}
