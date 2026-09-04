using System.Diagnostics;
using PptMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace PptMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Verifies the responsibility rule <see cref="SuiteOwnedPowerPoint"/> applies before it is
/// trusted to clean up after <see cref="PptBatchTimeoutTests"/> (issue #172).
///
/// <para>Both directions are asserted deliberately. A guard that reclaims everything would
/// terminate the user's own PowerPoint, which is the hazard issue #160 exists to prevent; a
/// guard that reclaims nothing would leave the survivor that issue #172 is about. Neither
/// failure is visible from the passing direction alone.</para>
///
/// <para>Each test uses a genuine POWERPNT process started by a real session, and injects the
/// snapshot so the process sits on a known side of the boundary. That keeps the outcome
/// independent of whatever PowerPoint the machine happens to have open, and guarantees the
/// guard is never asked about a process this test did not create.</para>
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "PptBatch")]
[Trait("RequiresPowerPoint", "true")]
[Collection("Sequential")]
public class SuiteOwnedPowerPointTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _testFile;

    public SuiteOwnedPowerPointTests(ITestOutputHelper output)
    {
        _output = output;

        var template = Path.Combine(
            AppContext.BaseDirectory, "Integration", "Session", "TestFiles", "batch-test-static.pptx");

        _testFile = Path.Combine(Path.GetTempPath(), $"suite-owned-ppt-{Guid.NewGuid():N}.pptx");
        File.Copy(template, _testFile);
    }

    public void Dispose()
    {
        if (File.Exists(_testFile))
        {
            try
            {
                File.Delete(_testFile);
            }
            catch (IOException)
            {
                // Best effort - a stuck PowerPoint may still hold the handle.
            }
        }

        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// A PowerPoint that appeared after the snapshot is the suite's, and must be terminated.
    /// This is the survivor that would otherwise be inherited by every later session.
    /// </summary>
    [Fact]
    public void Reclaim_TerminatesPowerPointThatAppearedAfterTheSnapshot()
    {
        var batch = PptSession.BeginBatch(_testFile);
        int pid = RequireProcessId(batch);

        // Snapshot excluding this process, so it stands as one that appeared while the suite ran.
        var guard = new SuiteOwnedPowerPoint(RunningPowerPointPids().Where(p => p != pid));

        Assert.Contains(pid, guard.Survivors());

        var reclaimed = guard.Reclaim();
        _output.WriteLine($"Reclaimed: {string.Join(", ", reclaimed)}");

        Assert.Contains(pid, reclaimed);
        Assert.True(
            HasExited(pid),
            $"PowerPoint process {pid} appeared while the suite was running, so the suite is " +
            "responsible for it. Leaving it alive strands an instance that every later session " +
            "attaches to and then declines to clean up.");

        Assert.DoesNotContain(pid, guard.Survivors());

        batch.Dispose();
    }

    /// <summary>
    /// A PowerPoint that was already running when the snapshot was taken is not the suite's, and
    /// must survive. It may be the user's own, holding unsaved work in decks we never opened.
    /// </summary>
    [Fact]
    public void Reclaim_LeavesPowerPointThatWasRunningBeforeTheSnapshot()
    {
        var batch = PptSession.BeginBatch(_testFile);
        int pid = RequireProcessId(batch);

        try
        {
            // Snapshot including this process, so it stands as one the user already had open.
            var guard = new SuiteOwnedPowerPoint(RunningPowerPointPids());

            Assert.DoesNotContain(pid, guard.Survivors());

            var reclaimed = guard.Reclaim();
            _output.WriteLine($"Reclaimed: [{string.Join(", ", reclaimed)}]");

            Assert.DoesNotContain(pid, reclaimed);
            Assert.False(
                HasExited(pid),
                $"PowerPoint process {pid} was already running when the snapshot was taken, so it " +
                "is not the suite's to terminate. Killing it would discard the user's unsaved work.");
        }
        finally
        {
            batch.Dispose();
        }
    }

    private static int RequireProcessId(IPptBatch batch)
    {
        int? pid = batch.PowerPointProcessId;

        Assert.True(
            pid.HasValue,
            "The session did not capture a PowerPoint process ID, so this test cannot place a known " +
            "process on either side of the responsibility boundary.");

        return pid!.Value;
    }

    private static List<int> RunningPowerPointPids()
    {
        var pids = new List<int>();

        foreach (var process in Process.GetProcessesByName("POWERPNT"))
        {
            try
            {
                pids.Add(process.Id);
            }
            finally
            {
                process.Dispose();
            }
        }

        return pids;
    }

    private static bool HasExited(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return process.HasExited;
        }
        catch (ArgumentException)
        {
            return true; // Exited and was reaped.
        }
    }
}
