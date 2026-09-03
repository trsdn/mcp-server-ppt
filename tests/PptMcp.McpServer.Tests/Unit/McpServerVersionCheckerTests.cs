using Xunit;

namespace PptMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "VersionCheck")]
[Trait("Speed", "Fast")]
public sealed class McpServerVersionCheckerTests
{
    private const string OptOutVariable = "PPTMCP_NO_UPDATE_CHECK";

    [Theory]
    [InlineData("1")]
    [InlineData("true")]
    [InlineData("TRUE")]
    [InlineData("yes")]
    public async Task CheckForUpdateAsync_WhenOptedOut_ReturnsNullWithoutNetworkCall(string value)
    {
        var previous = Environment.GetEnvironmentVariable(OptOutVariable);
        try
        {
            Environment.SetEnvironmentVariable(OptOutVariable, value);

            // A network call would take up to the 5s HttpClient timeout on a machine
            // where api.nuget.org is blocked. Opting out must short-circuit before that,
            // so the elapsed time is the assertion that no request was attempted.
            var started = System.Diagnostics.Stopwatch.StartNew();
            var latestVersion = await Infrastructure.McpServerVersionChecker.CheckForUpdateAsync();
            started.Stop();

            Assert.Null(latestVersion);
            Assert.True(
                started.ElapsedMilliseconds < 1000,
                $"Opted-out check took {started.ElapsedMilliseconds}ms, which means it still went to the network.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(OptOutVariable, previous);
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData("0")]
    [InlineData("false")]
    [InlineData("no")]
    public void IsUpdateCheckDisabled_WhenNotOptedOut_ReturnsFalse(string value)
    {
        var previous = Environment.GetEnvironmentVariable(OptOutVariable);
        try
        {
            Environment.SetEnvironmentVariable(OptOutVariable, value);

            Assert.False(Infrastructure.NuGetVersionChecker.IsUpdateCheckDisabled());
        }
        finally
        {
            Environment.SetEnvironmentVariable(OptOutVariable, previous);
        }
    }

    [Fact]
    public void IsUpdateCheckDisabled_WhenUnset_ReturnsFalse()
    {
        var previous = Environment.GetEnvironmentVariable(OptOutVariable);
        try
        {
            Environment.SetEnvironmentVariable(OptOutVariable, null);

            Assert.False(Infrastructure.NuGetVersionChecker.IsUpdateCheckDisabled());
        }
        finally
        {
            Environment.SetEnvironmentVariable(OptOutVariable, previous);
        }
    }

    [Fact]
    public async Task CheckForUpdateAsync_NetworkFailure_ReturnsNull()
    {
        using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(1));
        var latestVersion = await Infrastructure.McpServerVersionChecker.CheckForUpdateAsync(cts.Token);

        Assert.Null(latestVersion);
    }

    [Fact]
    public void GetCurrentVersion_ReturnsNonEmptyString()
    {
        var version = Infrastructure.McpServerVersionChecker.GetCurrentVersion();

        Assert.NotNull(version);
        Assert.NotEmpty(version);
        Assert.NotEqual("0.0.0", version);
    }
}

