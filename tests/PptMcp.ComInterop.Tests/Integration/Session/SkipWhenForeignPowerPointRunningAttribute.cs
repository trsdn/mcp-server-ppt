using Xunit;

namespace PptMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// A <see cref="FactAttribute"/> that skips itself when a POWERPNT process the test suite did not
/// start is already running.
///
/// <para><b>Why a test would decline to run.</b> A suite that fabricates "a PowerPoint the user
/// already had open" cannot do so while the user really does have one open, because PowerPoint is
/// single-instance: the stand-in launch hands off to the existing process. Every assertion then
/// applies to the user's instance, and an assertion that a process <i>was terminated</i> becomes a
/// demand that the user's PowerPoint be killed, discarding unsaved work in every deck they had
/// open with no prompt and no AutoRecover entry (issues #160, #181).</para>
///
/// <para><b>Why not simply fail.</b> A failure is safe, but it would make the local integration
/// gate unrunnable whenever the maintainer has PowerPoint open - which, for a PowerPoint
/// automation project, is most of the time. A gate that is routinely overridden stops being a
/// gate.</para>
///
/// <para><b>Why this rather than a package.</b> xunit v2 has no runtime skip - <c>Assert.Skip</c>
/// is v3-only - and the established package for it, <c>Xunit.SkippableFact</c>, is MS-PL, which is
/// outside this repository's dependency-review licence allow-list. <see cref="FactAttribute.Skip"/>
/// is settable, and xunit reads it during discovery, so a constructor that sets it achieves the
/// same result with no dependency at all.</para>
///
/// <para><b>Timing.</b> Discovery runs before the test class is constructed, so this reads the
/// machine slightly earlier than <see cref="SuiteOwnedPowerPoint"/>'s own snapshot. If PowerPoint
/// is opened in the gap, the tests run and fail rather than skip - a visible failure, never data
/// loss, because reclamation is still bounded by the constructor snapshot.</para>
/// </summary>
public sealed class SkipWhenForeignPowerPointRunningAttribute : FactAttribute
{
    /// <summary>
    /// Reads the running POWERPNT processes at discovery time and sets
    /// <see cref="FactAttribute.Skip"/> if any are present, naming the process IDs so the reason
    /// is visible in the run output rather than silent.
    /// </summary>
    public SkipWhenForeignPowerPointRunningAttribute()
    {
        var foreign = SuiteOwnedPowerPoint.RunningProcessIds();
        if (foreign.Count > 0)
        {
            Skip =
                $"PowerPoint is already running (PID {string.Join(", ", foreign)}) and was not started by this " +
                "suite. These tests fabricate a pre-existing instance, which is impossible while a real one is " +
                "open, and asserting termination against the user's process would discard their unsaved work. " +
                "Close PowerPoint and re-run.";
        }
    }
}
