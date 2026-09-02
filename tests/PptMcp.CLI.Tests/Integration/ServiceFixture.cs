using PptMcp.Service;
using Xunit;

namespace PptMcp.CLI.Tests.Integration;

/// <summary>
/// Fixture that starts an in-process PptMcp service for CLI integration tests.
/// Uses the CLI pipe name so CLI commands can connect to it.
/// </summary>
public sealed class ServiceFixture : IAsyncLifetime, IDisposable
{
    private PptMcpService? _service;

    public async Task InitializeAsync()
    {
        var pipeName = ServiceSecurity.GetCliPipeName();

        // A pptcli daemon left running by an earlier suite serves this same pipe name.
        // Windows allows several server instances per name, so starting ours would not
        // fail - the CLI under test would simply connect to whichever instance answers
        // first. That daemon has none of this process's in-process test hooks (the
        // reference-catalog override, for one), so tests would exercise the wrong
        // service and fail in ways that vanish when re-run in isolation. Refuse loudly.
        using (var probe = new ServiceClient(pipeName, connectTimeout: TimeSpan.FromSeconds(1)))
        {
            if (await probe.PingAsync())
            {
                throw new InvalidOperationException(
                    $"A PptMcp service is already serving the CLI pipe '{pipeName}' (most likely a leftover " +
                    "pptcli daemon). CLI integration tests must talk to the in-process service they configure, " +
                    "so this run would be meaningless. Stop it first: " +
                    "Get-Process pptcli | ForEach-Object { Stop-Process -Id $_.Id -Force }");
            }
        }

        _service = new PptMcpService();
        _ = Task.Run(() => _service.RunAsync(pipeName));

        // Wait for pipe server to be ready
        for (int i = 0; i < 20; i++)
        {
            await Task.Delay(100);
            using var client = new ServiceClient(pipeName, connectTimeout: TimeSpan.FromSeconds(1));
            if (await client.PingAsync())
            {
                return;
            }
        }

        throw new InvalidOperationException("PptMcp service did not start within timeout.");
    }

    public Task DisposeAsync()
    {
        Dispose();
        return Task.CompletedTask;
    }

    public void Dispose()
    {
        _service?.RequestShutdown();
        _service?.Dispose();
        _service = null;
    }
}

/// <summary>
/// Collection definition for tests that require the PptMcp service.
/// Apply [Collection("Service")] to test classes that call pptcli commands.
/// </summary>
[CollectionDefinition("Service")]
public sealed class ServiceTestGroup : ICollectionFixture<ServiceFixture>;
