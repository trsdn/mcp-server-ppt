using Xunit;

namespace PptMcp.ComInterop.Tests.Unit;

/// <summary>
/// Unit tests for OleMessageFilter registration and revocation.
/// Tests verify that the message filter can be registered/revoked without errors.
///
/// NOTE: These tests verify the registration mechanism but don't test actual
/// COM retry behavior (that requires PowerPoint and would be OnDemand tests).
/// </summary>
[Trait("Category", "Unit")]
[Trait("Feature", "OleMessageFilter")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
public class OleMessageFilterTests
{
    [Fact]
    public void Register_OnStaThread_DoesNotThrow()
    {
        // Arrange & Act & Assert
        var thread = new Thread(() =>
        {
            try
            {
                OleMessageFilter.Register();
                OleMessageFilter.Revoke();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Filter registration failed: {ex.Message}", ex);
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
    }

    [Fact]
    public void RegisterAndRevoke_MultipleTimes_DoesNotThrow()
    {
        // Arrange & Act & Assert
        var thread = new Thread(() =>
        {
            // First registration
            OleMessageFilter.Register();
            OleMessageFilter.Revoke();

            // Second registration (simulates reuse)
            OleMessageFilter.Register();
            OleMessageFilter.Revoke();
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
    }

    [Fact]
    public void Revoke_WithoutRegister_DoesNotThrow()
    {
        // Revoke without prior Register should not crash
        // Arrange & Act & Assert - Should handle gracefully
        var thread = new Thread(OleMessageFilter.Revoke);

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
    }

    /// <summary>
    /// Regression guard for the STA deadlock: <c>MessagePending</c> must return
    /// <c>PENDINGMSG_WAITDEFPROCESS</c> (2), never <c>PENDINGMSG_WAITNOPROCESS</c> (1).
    ///
    /// WAITNOPROCESS blocks all inbound COM message processing while an outgoing call
    /// is in progress. When PowerPoint fires a re-entrant callback during a long
    /// operation, the callback is queued but never dispatched: PowerPoint waits for the
    /// callback, the STA thread waits for PowerPoint, and both stop.
    ///
    /// WAITDEFPROCESS lets COM dispatch the pending inbound call so the callback can
    /// complete and the outgoing call returns normally.
    ///
    /// This test was itself broken until issue #126's sweep. It declared the two
    /// constants with their values exchanged - WAITDEFPROCESS as 1 and WAITNOPROCESS
    /// as 2 - and its prose repeated the same inversion. The Win32 header is
    /// unambiguous: PENDINGMSG_CANCELCALL = 0, PENDINGMSG_WAITNOPROCESS = 1,
    /// PENDINGMSG_WAITDEFPROCESS = 2.
    ///
    /// The consequence was worse than a red test. The production code returns 2, which
    /// is correct, so the assertions were unsatisfiable and the test could never pass.
    /// Anyone who had "fixed" the code to satisfy it would have changed the return to 1
    /// - which is WAITNOPROCESS, the exact value that causes the deadlock this test
    /// claims to guard against.
    /// </summary>
    [Fact]
    public void MessagePending_ReturnValue_MustBe_WaitDefProcess()
    {
        // IOleMessageFilter is internal, so the filter is instantiated and invoked
        // through the interface by reflection, on a real STA thread.
        //
        // Win32 PENDINGMSG (objidl.h). These were previously declared with their values
        // exchanged, which made the two assertions below mutually unsatisfiable.
        const int PENDINGMSG_CANCELCALL = 0;
        const int PENDINGMSG_WAITNOPROCESS = 1;
        const int PENDINGMSG_WAITDEFPROCESS = 2;

        var returnValue = -1;
        Exception? threadException = null;

        var thread = new Thread(() =>
        {
            try
            {
                OleMessageFilter.Register();

                // The filter implements IOleMessageFilter which is internal.
                // We can verify via the public static IsRegistered and the logical behavior:
                // After Register(), the filter IS the active message filter for this thread.
                //
                // Verify that the filter is registered (prerequisite for the bug to manifest).
                Assert.True(OleMessageFilter.IsRegistered, "Filter must be registered to have any effect");

                // Use reflection to invoke MessagePending on the filter instance.
                // The filter class is internal, but we can get to it via the assembly.
                var filterType = typeof(OleMessageFilter);
                var iOleMsgFilterType = filterType.Assembly.GetType(
                    "PptMcp.ComInterop.IOleMessageFilter");
                Assert.NotNull(iOleMsgFilterType);

                // Create a filter instance and call MessagePending
                var filterInstance = Activator.CreateInstance(filterType);
                Assert.NotNull(filterInstance);
                var method = iOleMsgFilterType.GetMethod("MessagePending");
                Assert.NotNull(method);

                returnValue = (int)method.Invoke(filterInstance, [IntPtr.Zero, 1000, 1])!;
                OleMessageFilter.Revoke();
            }
            catch (Exception ex)
            {
                threadException = ex;
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (threadException != null) throw new InvalidOperationException($"Thread exception: {threadException.Message}", threadException);

        // WAITNOPROCESS queues the inbound callback without dispatching it, which
        // deadlocks the STA thread against PowerPoint. CANCELCALL abandons the
        // outgoing call outright. Only WAITDEFPROCESS is correct here.
        Assert.NotEqual(PENDINGMSG_CANCELCALL, returnValue);
        Assert.NotEqual(PENDINGMSG_WAITNOPROCESS, returnValue);
        Assert.Equal(PENDINGMSG_WAITDEFPROCESS, returnValue);
    }
}





