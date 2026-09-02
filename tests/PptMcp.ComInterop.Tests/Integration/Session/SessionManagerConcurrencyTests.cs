using System.Collections.Concurrent;
using PptMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace PptMcp.ComInterop.Tests.Integration;

/// <summary>
/// Verifies that the single-instance invariant survives concurrent callers (issue #132).
///
/// SessionManager enforces "only one session at a time" with a check-then-act:
/// <c>if (!_activeSessions.IsEmpty) throw;</c> followed - much later - by a
/// <c>TryAdd</c>. Each dictionary operation is individually atomic, but the *sequence*
/// is not, and between the two sits file I/O plus a full COM bootstrap that takes
/// roughly 2.5 seconds. That is an enormous window: two callers arriving within it both
/// observe an empty dictionary and both proceed.
///
/// <c>TryAdd</c> cannot act as a backstop, because the key is a freshly generated GUID
/// and therefore never collides.
///
/// These tests fail against the unsynchronised implementation with two live sessions
/// against a single-instance COM server.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresPowerPoint", "true")]
[Collection("Sequential")]
public class SessionManagerConcurrencyTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;

    private static readonly string TemplateFilePath = Path.Combine(
        Path.GetDirectoryName(typeof(SessionManagerConcurrencyTests).Assembly.Location)!,
        "Integration", "Session", "TestFiles", "batch-test-static.pptx");

    public SessionManagerConcurrencyTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"SessionManagerConcurrency_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);

        if (Directory.Exists(_tempDir))
        {
            try { Directory.Delete(_tempDir, recursive: true); } catch (IOException) { }
        }
    }

    private string CreateTestFile(string testName)
    {
        var filePath = Path.Combine(_tempDir, $"{testName}_{Guid.NewGuid():N}.pptx");
        File.Copy(TemplateFilePath, filePath);
        return filePath;
    }

    /// <summary>
    /// Two threads call CreateSession at the same instant on two different files.
    /// Different files, so the file-path guard cannot mask the race - only the
    /// single-instance guard is under test.
    /// </summary>
    [Fact]
    public void CreateSession_TwoConcurrentCallers_OnlyOneSucceeds()
    {
        var fileA = CreateTestFile("concurrentA");
        var fileB = CreateTestFile("concurrentB");

        using var manager = new SessionManager();

        var (sessionIds, failures) = RaceTwo(
            () => manager.CreateSession(fileA),
            () => manager.CreateSession(fileB));

        _output.WriteLine($"succeeded: {sessionIds.Count}, failed: {failures.Count}");
        foreach (var f in failures)
        {
            _output.WriteLine($"  rejected with: {f.GetType().Name}: {f.Message}");
        }

        try
        {
            Assert.Single(sessionIds);
            Assert.Single(failures);
            Assert.IsType<InvalidOperationException>(failures[0]);
            Assert.Contains("single-instance", failures[0].Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, manager.ActiveSessionCount);
        }
        finally
        {
            foreach (var id in sessionIds)
            {
                try { manager.CloseSession(id); } catch (InvalidOperationException) { }
            }
        }
    }

    /// <summary>
    /// Same race on the create-new-file path, which carries its own copy of the
    /// check-then-act and its own slow bootstrap.
    /// </summary>
    [Fact]
    public void CreateSessionForNewFile_TwoConcurrentCallers_OnlyOneSucceeds()
    {
        var fileA = Path.Combine(_tempDir, $"newA_{Guid.NewGuid():N}.pptx");
        var fileB = Path.Combine(_tempDir, $"newB_{Guid.NewGuid():N}.pptx");

        using var manager = new SessionManager();

        var (sessionIds, failures) = RaceTwo(
            () => manager.CreateSessionForNewFile(fileA),
            () => manager.CreateSessionForNewFile(fileB));

        _output.WriteLine($"succeeded: {sessionIds.Count}, failed: {failures.Count}");
        foreach (var f in failures)
        {
            _output.WriteLine($"  rejected with: {f.GetType().Name}: {f.Message}");
        }

        try
        {
            Assert.Single(sessionIds);
            Assert.Single(failures);
            Assert.IsType<InvalidOperationException>(failures[0]);
            Assert.Contains("single-instance", failures[0].Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, manager.ActiveSessionCount);
        }
        finally
        {
            foreach (var id in sessionIds)
            {
                try { manager.CloseSession(id); } catch (InvalidOperationException) { }
            }
        }
    }

    /// <summary>
    /// The loser of the race must not leave debris behind: after closing the winner,
    /// the manager must accept a new session. A reservation that is taken and never
    /// released would wedge the manager permanently, which is a worse failure than the
    /// race it was introduced to fix.
    /// </summary>
    [Fact]
    public void CreateSession_AfterLosingCallerIsRejected_ManagerStillUsable()
    {
        var fileA = CreateTestFile("recoverA");
        var fileB = CreateTestFile("recoverB");

        using var manager = new SessionManager();

        var (sessionIds, failures) = RaceTwo(
            () => manager.CreateSession(fileA),
            () => manager.CreateSession(fileB));

        Assert.Single(sessionIds);
        Assert.Single(failures);

        manager.CloseSession(sessionIds[0]);
        Assert.Equal(0, manager.ActiveSessionCount);

        var reopened = manager.CreateSession(fileA);
        Assert.Equal(1, manager.ActiveSessionCount);
        manager.CloseSession(reopened);
    }

    /// <summary>
    /// Releases both callers from a single barrier so they enter the check-then-act
    /// window together, which is the only way to exercise it.
    /// </summary>
    private static (List<string> SessionIds, List<Exception> Failures) RaceTwo(
        Func<string> first, Func<string> second)
    {
        var sessionIds = new ConcurrentBag<string>();
        var failures = new ConcurrentBag<Exception>();

        using var startLine = new Barrier(2);

        void Run(Func<string> create)
        {
            startLine.SignalAndWait();
            try
            {
                sessionIds.Add(create());
            }
            catch (Exception ex)
            {
                failures.Add(ex);
            }
        }

        var t1 = new Thread(() => Run(first)) { IsBackground = true };
        var t2 = new Thread(() => Run(second)) { IsBackground = true };

        t1.Start();
        t2.Start();

        // Generous: a real bootstrap takes ~2.5s, and the losing caller may have to wait
        // for the winner to finish before it can be told no.
        Assert.True(t1.Join(TimeSpan.FromMinutes(2)), "first caller did not finish");
        Assert.True(t2.Join(TimeSpan.FromMinutes(2)), "second caller did not finish");

        return (sessionIds.ToList(), failures.ToList());
    }
}
