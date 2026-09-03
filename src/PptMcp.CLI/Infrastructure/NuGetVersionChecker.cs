using System.Net.Http.Json;
using System.Text.Json.Serialization;

namespace PptMcp.CLI.Infrastructure;

/// <summary>
/// Checks NuGet for the latest version of the CLI package.
/// </summary>
internal static class NuGetVersionChecker
{
    // The flat container API requires the package ID lowercased (LOWER_ID); any other
    // casing returns 404. Published ID is "PptMcp.CLI".
    private const string PackageId = "pptmcp.cli";
    private const string NuGetIndexUrl = $"https://api.nuget.org/v3-flatcontainer/{PackageId}/index.json";
    private static readonly TimeSpan Timeout = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Environment variable that disables the update check entirely.
    /// </summary>
    internal const string OptOutVariable = "PPTMCP_NO_UPDATE_CHECK";

    /// <summary>
    /// Returns true when the update check has been disabled via
    /// <c>PPTMCP_NO_UPDATE_CHECK</c>. Accepts 1, true, yes or on, case-insensitively.
    /// </summary>
    public static bool IsUpdateCheckDisabled()
    {
        // Some environments block public package registries as policy rather than as a
        // firewall accident. There the check can never succeed, so it only ever costs the
        // full HttpClient timeout - and opting out must not require a network change.
        var value = Environment.GetEnvironmentVariable(OptOutVariable)?.Trim();

        return !string.IsNullOrEmpty(value)
            && (value == "1"
                || value.Equals("true", StringComparison.OrdinalIgnoreCase)
                || value.Equals("yes", StringComparison.OrdinalIgnoreCase)
                || value.Equals("on", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Checks NuGet for the latest version.
    /// </summary>
    /// <returns>Latest version string, or null if check failed.</returns>
    public static async Task<string?> GetLatestVersionAsync(CancellationToken cancellationToken = default)
    {
        if (IsUpdateCheckDisabled())
        {
            return null;
        }

        try
        {
            using var httpClient = new HttpClient { Timeout = Timeout };
            var response = await httpClient.GetFromJsonAsync<NuGetVersionsResponse>(NuGetIndexUrl, cancellationToken);

            if (response?.Versions == null || response.Versions.Count == 0)
                return null;

            // Get highest non-prerelease version
            var latestVersion = response.Versions
                .Where(v => !v.Contains('-')) // Exclude prerelease versions
                .OrderByDescending(v => ParseVersion(v))
                .FirstOrDefault();

            return latestVersion ?? response.Versions.Last();
        }
        catch (Exception)
        {
            // Network error, timeout, etc. — return null to indicate check failed
            return null;
        }
    }

    private static Version ParseVersion(string versionString)
    {
        // Handle versions like "1.2.3" - strip any suffix after +
        var cleanVersion = versionString.Split('+')[0].Split('-')[0];
        return Version.TryParse(cleanVersion, out var version) ? version : new Version(0, 0, 0);
    }

    private sealed class NuGetVersionsResponse
    {
        [JsonPropertyName("versions")]
        public List<string> Versions { get; set; } = [];
    }
}
