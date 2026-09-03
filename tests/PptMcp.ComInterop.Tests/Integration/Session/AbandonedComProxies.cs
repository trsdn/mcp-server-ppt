using System.Diagnostics;

namespace PptMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Releases the COM proxies a force-killing suite abandons, before the next test inherits them.
///
/// <see cref="PptBatchTimeoutTests"/> force-kills POWERPNT while COM calls are still in
/// flight, which leaves RCWs in the finalizer queue pointing at dead RPC endpoints.
/// Releasing one of those blocks until COM's own call timeout expires - measured at
/// 15-20s worth per force-killing test.
///
/// That queue is process-wide, so the cost otherwise lands on whichever test next calls
/// <see cref="GC.WaitForPendingFinalizers"/> - and every session teardown does, inside
/// PptShutdownService.CloseAndQuit. That is how the timeout suite charged ~14s of teardown
/// to SessionShutdownCleanlinessTests, whose own application-proxy release measures 1ms
/// (issue #161). The cleanliness guard was failing on a cost it did not create.
///
/// Draining after every test, rather than once per class, is the measured answer rather
/// than the cheap one. The class-scoped variant looked attractive - it cost ~70s less,
/// on the theory that the finalizer thread releases these in the background anyway - and
/// it passed twice before failing on the third run, with the cleanliness test still
/// reporting a 10847ms drain. xUnit disposes a class fixture too late to help: by then
/// the next class is already running. Two green runs were luck, not evidence.
///
/// This is a settle condition, not a sleep: <see cref="GC.WaitForPendingFinalizers"/>
/// returns when the queue is empty, however long that takes, so it cannot degrade into a
/// slower flake the way a fixed delay would.
/// </summary>
internal static class AbandonedComProxies
{
    /// <summary>
    /// Blocks until every abandoned COM proxy has been finalised. Returns how long that took,
    /// so a suite can report the cleanup cost it is charging to itself rather than to its successor.
    /// </summary>
    public static TimeSpan Drain()
    {
        var sw = Stopwatch.StartNew();

        // Collect to queue the abandoned RCWs, wait for the finalizer thread to release
        // them, then collect again to reclaim what the finalizers freed.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        sw.Stop();
        return sw.Elapsed;
    }
}
