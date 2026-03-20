using System.Security.Cryptography;
using System.Text;

namespace PptMcp.Core.Tests.Helpers;

internal static class ReferenceCatalogFixture
{
    public const string FrameworkMatrixRawId = "fixture-framework-001";
    public const string FrameworkAuditRawId = "fixture-framework-002";
    public const string FrameworkPillarsRawId = "fixture-framework-003";
    public const string TableCriteriaRawId1 = "fixture-table-001";
    public const string TableCriteriaRawId2 = "fixture-table-002";
    public const string OrgChartRawId = "fixture-org-001";

    public static string FrameworkMatrixReferenceId => GetPublicReferenceId(FrameworkMatrixRawId);
    public static string FrameworkAuditReferenceId => GetPublicReferenceId(FrameworkAuditRawId);
    public static string TableCriteriaReferenceId1 => GetPublicReferenceId(TableCriteriaRawId1);
    public static string TableCriteriaReferenceId2 => GetPublicReferenceId(TableCriteriaRawId2);
    public static string OrgChartReferenceId => GetPublicReferenceId(OrgChartRawId);

    public static string GetCatalogRoot()
    {
        var directOutputPath = Path.Combine(AppContext.BaseDirectory, "TestAssets", "ReferenceCatalog");
        if (HasCatalog(directOutputPath))
        {
            return directOutputPath;
        }

        var current = new DirectoryInfo(AppContext.BaseDirectory);
        while (current is not null)
        {
            var repoPath = Path.Combine(current.FullName, "tests", "PptMcp.Core.Tests", "TestAssets", "ReferenceCatalog");
            if (HasCatalog(repoPath))
            {
                return repoPath;
            }

            current = current.Parent;
        }

        throw new DirectoryNotFoundException("Reference catalog fixture files not found.");
    }

    public static Dictionary<string, string> GetEnvironmentVariables()
    {
        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["PPTMCP_REFERENCE_DATA_ROOT"] = GetCatalogRoot()
        };
    }

    public static string GetPublicReferenceId(string rawReferenceId)
    {
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(rawReferenceId));
        return $"ref-{Convert.ToHexString(hash.AsSpan(0, 6)).ToLowerInvariant()}";
    }

    private static bool HasCatalog(string candidateRoot)
    {
        return File.Exists(Path.Combine(candidateRoot, "manifest.json"))
            && File.Exists(Path.Combine(candidateRoot, "sub-archetypes.json"))
            && File.Exists(Path.Combine(candidateRoot, "new-archetypes.json"));
    }
}
