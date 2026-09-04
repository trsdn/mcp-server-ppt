using System.Diagnostics;

namespace PptMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Identifies the POWERPNT processes a suite is responsible for, and reclaims them.
///
/// <para><b>Why this exists.</b> A session that force-kills PowerPoint can leave a survivor,
/// and a survivor is not merely untidy - it is permanent. The next session attaches to it,
/// because PowerPoint is single-instance, and therefore reports <c>_ownsPowerPointProcess =
/// false</c>. If that session then times out, teardown declines to force-kill (correctly: the
/// process might be the user's, issue #160), so its STA thread stays wedged, so STA cleanup
/// never runs, so <c>_foreignPresentationCount</c> is never computed and stays null. The
/// terminate-on-disposal check is <c>_ownsPowerPointProcess || _foreignPresentationCount == 0</c>,
/// which is then false on both sides. Every later session inherits the same instance and makes
/// the same refusal (issue #172).</para>
///
/// <para><b>Why the fix is here and not in PptBatch.</b> That refusal is right in production.
/// An unknown foreign-presentation count means "the user may have unsaved work open", and the
/// safe answer to that is to leave the process alone. What is wrong is a test suite that
/// force-kills PowerPoint by design leaving a survivor behind for its successors to inherit.</para>
///
/// <para><b>The responsibility rule.</b> Snapshot the running POWERPNT process IDs before the
/// suite's first test. Anything running afterwards that is not in that snapshot appeared while
/// the suite was running and is the suite's to reclaim. Anything in the snapshot is left strictly
/// alone - it may be the user's own PowerPoint, and it is the same signal
/// <c>PptBatch</c> itself uses to decide ownership.</para>
///
/// <para>This asserts observable state rather than waiting out a delay, so it cannot decay into
/// a slower flake the way a sleep would.</para>
/// </summary>
internal sealed class SuiteOwnedPowerPoint
{
    private readonly HashSet<int> _preExisting;

    /// <summary>
    /// Snapshots the POWERPNT processes running right now. Every one of them is off limits.
    /// </summary>
    public SuiteOwnedPowerPoint()
        : this(RunningPowerPointProcessIds())
    {
    }

    /// <summary>
    /// Uses an explicit snapshot. Tests use this to place a process they created on either side
    /// of the responsibility boundary without depending on what was running on the machine.
    /// </summary>
    public SuiteOwnedPowerPoint(IEnumerable<int> preExisting)
    {
        _preExisting = new HashSet<int>(preExisting);
    }

    /// <summary>
    /// POWERPNT processes running right now. A caller that needs to reason about the snapshot
    /// itself - for example to decline to run at all while a foreign instance is open - takes it
    /// from here and passes it to the constructor, so the guard and the decision cannot disagree
    /// about which processes were pre-existing.
    /// </summary>
    public static IReadOnlyList<int> RunningProcessIds() => RunningPowerPointProcessIds();

    /// <summary>
    /// POWERPNT processes running now that were not running when the snapshot was taken.
    /// </summary>
    public IReadOnlyList<int> Survivors() =>
        RunningPowerPointProcessIds().Where(pid => !_preExisting.Contains(pid)).ToList();

    /// <summary>
    /// Terminates every survivor and returns the process IDs that were terminated, so a caller
    /// can report what it cleaned up rather than cleaning up silently. An empty result means the
    /// suite left nothing behind.
    /// </summary>
    public IReadOnlyList<int> Reclaim()
    {
        var reclaimed = new List<int>();

        foreach (var pid in Survivors())
        {
            try
            {
                using var process = Process.GetProcessById(pid);
                if (process.HasExited)
                {
                    continue;
                }

                process.Kill(entireProcessTree: true);
                process.WaitForExit(5000);
                reclaimed.Add(pid);
            }
            catch (ArgumentException)
            {
                // Exited and was reaped between the enumeration and the kill.
            }
        }

        return reclaimed;
    }

    private static List<int> RunningPowerPointProcessIds()
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
}
