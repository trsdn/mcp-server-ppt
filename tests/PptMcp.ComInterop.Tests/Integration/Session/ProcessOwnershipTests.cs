using System.Diagnostics;
using Microsoft.Win32;
using PptMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace PptMcp.ComInterop.Tests.Integration;

/// <summary>
/// Verifies that session teardown only terminates a PowerPoint process the session itself started.
///
/// PowerPoint is a single-instance application: Activator.CreateInstance attaches to an instance
/// the user already had open rather than starting a second one. Terminating that process discards
/// every unsaved change in every deck the user had open, with no prompt and no recovery entry.
///
/// Both directions are asserted deliberately. Skipping termination unconditionally would leak a
/// PowerPoint process that every later session attaches to, so the owned-process path must keep
/// terminating as before (issue #148).
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresPowerPoint", "true")]
[Collection("Sequential")]
public class ProcessOwnershipTests : IAsyncLifetime
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly List<string> _testFiles = new();

    private static readonly string TemplateFilePath = Path.Combine(
        AppContext.BaseDirectory,
        "Integration", "Session", "TestFiles", "batch-test-static.pptx");

    public ProcessOwnershipTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"ProcessOwnershipTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public Task InitializeAsync()
    {
        KillAllPowerPoint();
        return Task.CompletedTask;
    }

    public Task DisposeAsync()
    {
        KillAllPowerPoint();

        foreach (var file in _testFiles)
        {
            try
            {
                if (File.Exists(file))
                {
                    File.Delete(file);
                }
            }
            catch (IOException)
            {
                // Best effort - a stuck PowerPoint may still hold the handle
            }
        }

        try
        {
            Directory.Delete(_tempDir, recursive: true);
        }
        catch (IOException)
        {
            // Best effort
        }

        return Task.CompletedTask;
    }

    /// <summary>
    /// A PowerPoint process the session started itself must still be terminated on disposal.
    /// This guards the issue #148 teardown win against being silently disabled by the
    /// ownership check.
    /// </summary>
    [Fact]
    public void Dispose_ProcessStartedBySession_IsTerminated()
    {
        Assert.Empty(Process.GetProcessesByName("POWERPNT"));

        var testFile = CreateTestFile(nameof(Dispose_ProcessStartedBySession_IsTerminated));

        var batch = new PptBatch(new[] { testFile });
        var owned = Process.GetProcessesByName("POWERPNT");
        Assert.NotEmpty(owned);
        var ownedPid = owned[0].Id;
        _output.WriteLine($"Session started PowerPoint process {ownedPid}");

        batch.Dispose();

        Assert.True(
            HasExited(ownedPid),
            $"PowerPoint process {ownedPid} was started by the session and must be terminated on disposal, " +
            "otherwise every later session attaches to the leaked instance.");
    }

    /// <summary>
    /// A PowerPoint process that was already running before the session started, and that still has
    /// the user's own presentation open, must survive disposal. Killing it discards unsaved work in
    /// every deck the user had open.
    /// </summary>
    [Fact]
    public void Dispose_PreExistingProcessHoldingUserWork_IsNotTerminated()
    {
        Assert.Empty(Process.GetProcessesByName("POWERPNT"));

        var userDeck = CreateTestFile("UserDeck");
        var preExistingPid = StartPowerPointOutsideSession(userDeck);
        _output.WriteLine($"Pre-existing PowerPoint process {preExistingPid} holding {Path.GetFileName(userDeck)}");

        var testFile = CreateTestFile(nameof(Dispose_PreExistingProcessHoldingUserWork_IsNotTerminated));

        var batch = new PptBatch(new[] { testFile });

        // PowerPoint is single-instance, so the session must have attached to the existing process
        // rather than starting its own. If that ever stops being true this test is not exercising
        // the hazard at all, so assert it rather than assume it.
        var running = Process.GetProcessesByName("POWERPNT");
        Assert.Single(running);
        Assert.Equal(preExistingPid, running[0].Id);

        batch.Dispose();

        Assert.False(
            HasExited(preExistingPid),
            $"PowerPoint process {preExistingPid} was already running with the user's own presentation open. " +
            "Terminating it discards their unsaved work.");
    }

    /// <summary>
    /// A PowerPoint the session attached to but which has nothing left open is terminated.
    ///
    /// Without this, a stranded instance would be inherited by every later session - each one
    /// declining to clean it up - while it holds presentation files locked. That cascade is a real
    /// observed failure mode, not a hypothetical: it broke unrelated tests when termination was
    /// withheld on the attach signal alone.
    /// </summary>
    [Fact]
    public void Dispose_AttachedProcessWithNoUserWork_IsTerminated()
    {
        Assert.Empty(Process.GetProcessesByName("POWERPNT"));

        var preExistingPid = StartPowerPointOutsideSession(documentPath: null);
        _output.WriteLine($"Pre-existing empty PowerPoint process {preExistingPid}");

        var testFile = CreateTestFile(nameof(Dispose_AttachedProcessWithNoUserWork_IsTerminated));

        var batch = new PptBatch(new[] { testFile });

        var running = Process.GetProcessesByName("POWERPNT");
        Assert.Single(running);
        Assert.Equal(preExistingPid, running[0].Id);

        batch.Dispose();

        Assert.True(
            HasExited(preExistingPid),
            $"PowerPoint process {preExistingPid} had no presentations left open, so no unsaved work was at risk. " +
            "Leaving it running strands an instance that later sessions inherit while it holds files locked.");
    }

    private string CreateTestFile(string testName)
    {
        var fileName = $"{testName}_{Guid.NewGuid():N}.pptx";
        var filePath = Path.Combine(_tempDir, fileName);
        File.Copy(TemplateFilePath, filePath);
        _testFiles.Add(filePath);
        return filePath;
    }

    /// <summary>
    /// Launches POWERPNT.EXE directly, standing in for an instance the user already had open.
    /// COM is deliberately not used here so the process is genuinely external to the session.
    /// </summary>
    /// <param name="documentPath">
    /// Presentation to open, standing in for the user's own work. Null launches an empty instance.
    /// </param>
    private static int StartPowerPointOutsideSession(string? documentPath)
    {
        var exePath = ResolvePowerPointExecutable();
        var startInfo = new ProcessStartInfo(exePath) { UseShellExecute = true };
        if (documentPath != null)
        {
            startInfo.ArgumentList.Add(documentPath);
        }

        var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException($"Failed to start {exePath}");

        // Wait for the window to exist, so the session attaches to a fully initialised instance.
        var deadline = DateTime.UtcNow.AddSeconds(30);
        while (DateTime.UtcNow < deadline)
        {
            process.Refresh();
            if (process.HasExited)
            {
                throw new InvalidOperationException("PowerPoint exited immediately after launch");
            }

            if (process.MainWindowHandle != IntPtr.Zero)
            {
                return process.Id;
            }

            Thread.Sleep(250);
        }

        throw new TimeoutException("PowerPoint did not present a window within 30 seconds");
    }

    private static string ResolvePowerPointExecutable()
    {
        using var key = Registry.LocalMachine.OpenSubKey(
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE");
        var path = key?.GetValue(null) as string;

        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            throw new InvalidOperationException(
                "POWERPNT.EXE could not be located via App Paths. PowerPoint is required for this test.");
        }

        return path;
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
            // No such process - it exited and was reaped
            return true;
        }
    }

    private void KillAllPowerPoint()
    {
        foreach (var process in Process.GetProcessesByName("POWERPNT"))
        {
            try
            {
                process.Kill(entireProcessTree: true);
                process.WaitForExit(5000);
            }
            catch (Exception ex)
            {
                _output.WriteLine($"Warning: failed to kill PowerPoint {process.Id}: {ex.Message}");
            }
            finally
            {
                process.Dispose();
            }
        }
    }
}
