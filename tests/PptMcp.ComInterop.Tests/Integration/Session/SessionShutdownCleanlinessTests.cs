using Microsoft.Extensions.Logging;
using PptMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace PptMcp.ComInterop.Tests.Integration;

/// <summary>
/// Regression tests for issue #148: a normal session teardown must let PowerPoint
/// exit on its own. If any COM proxy created during session bootstrap is abandoned,
/// POWERPNT stays alive, Dispose burns its full grace period, and then force-kills
/// the process. That cost is paid by every session - tests, CLI and MCP alike.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "SessionShutdown")]
[Trait("RequiresPowerPoint", "true")]
[Collection("Sequential")]
public class SessionShutdownCleanlinessTests : IDisposable
{
    private readonly ITestOutputHelper _out;
    private readonly string _tempDir;

    public SessionShutdownCleanlinessTests(ITestOutputHelper output)
    {
        _out = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"SessionShutdownCleanliness_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);
        if (Directory.Exists(_tempDir))
        {
            try { Directory.Delete(_tempDir, recursive: true); }
            catch (IOException) { /* best effort */ }
            catch (UnauthorizedAccessException) { /* best effort */ }
        }
    }

    private static readonly string TemplateFilePath = Path.Combine(
        Path.GetDirectoryName(typeof(SessionShutdownCleanlinessTests).Assembly.Location)!,
        "Integration", "Session", "TestFiles", "batch-test-static.pptx");

    private string CreateTestFile(string name)
    {
        var path = Path.Combine(_tempDir, $"{name}_{Guid.NewGuid():N}.pptx");
        File.Copy(TemplateFilePath, path);
        return path;
    }

    /// <summary>
    /// Session teardown must be quick. Every multi-second wait in the shutdown path is a
    /// timeout that only elapses when something is wrong, so a clean Dispose that runs into
    /// one is a regression - it is paid by every session, in tests, CLI and MCP alike.
    /// </summary>
    [Fact]
    public void Dispose_CompletesWithoutRunningIntoATimeout()
    {
        var file = CreateTestFile("clean_shutdown");
        var logger = new CapturingLogger();

        var batch = new PptBatch([file], logger);
        int? pid = batch.PowerPointProcessId;

        batch.Execute((ctx, ct) =>
        {
            dynamic? slides = null;
            try
            {
                slides = ctx.Presentation.Slides;
                return (int)slides.Count;
            }
            finally
            {
                if (slides != null) ComUtilities.Release(ref slides!);
            }
        });

        var sw = System.Diagnostics.Stopwatch.StartNew();
        batch.Dispose();
        sw.Stop();

        foreach (var message in logger.Messages)
        {
            _out.WriteLine(message);
        }

        _out.WriteLine($"Dispose took {sw.Elapsed.TotalSeconds:F1}s");

        // The shutdown path contains several timeouts that must never be reached on a clean
        // teardown: the 45s STA join, the COM call timeout while releasing the application
        // proxy (40-60s), and the process termination grace period. Any of them would put
        // Dispose far above this ceiling, which is otherwise loose enough to absorb normal
        // machine variance.
        Assert.True(
            sw.Elapsed.TotalSeconds < 10,
            $"Dispose took {sw.Elapsed.TotalSeconds:F1}s, which means it ran into one of the " +
            "shutdown timeouts instead of completing cleanly - see issue #148. Log:\n" +
            string.Join("\n", logger.Messages));

        // The STA thread must shut down under its own steam. Reaching the join timeout means a
        // COM proxy created during bootstrap was abandoned and the thread could not exit.
        var stuck = logger.Messages
            .Where(m => m.Contains("did NOT exit within", StringComparison.OrdinalIgnoreCase))
            .ToList();

        Assert.True(
            stuck.Count == 0,
            "The STA thread could not exit, which means a COM proxy created during session " +
            "bootstrap was never released. Offending log lines:\n" + string.Join("\n", stuck));

        // The session owns the PowerPoint process it started, so nothing may be left behind.
        Assert.NotNull(pid);
        Assert.True(HasExited(pid!.Value), $"PowerPoint process {pid} was still running after Dispose.");
    }

    /// <summary>
    /// The create-new-file bootstrap path binds the Presentations collection just like the
    /// open-existing path does, and must release it just as cleanly. This path is reached
    /// through SessionManager rather than PptSession.BeginBatch, so it is easy to regress
    /// independently of the path above.
    /// </summary>
    [Fact]
    public void Dispose_IsCleanForNewlyCreatedPresentations()
    {
        var path = Path.Combine(_tempDir, $"created_{Guid.NewGuid():N}.pptx");
        var logger = new CapturingLogger();

        var batch = PptBatch.CreateNewPresentation(path, isMacroEnabled: false, logger: logger);
        int? pid = batch.PowerPointProcessId;

        var slideCount = batch.Execute((ctx, ct) =>
        {
            dynamic? slides = null;
            try
            {
                slides = ctx.Presentation.Slides;
                return (int)slides.Count;
            }
            finally
            {
                if (slides != null) ComUtilities.Release(ref slides!);
            }
        });

        var sw = System.Diagnostics.Stopwatch.StartNew();
        batch.Dispose();
        sw.Stop();

        foreach (var message in logger.Messages)
        {
            _out.WriteLine(message);
        }

        _out.WriteLine($"Slides: {slideCount}, Dispose took {sw.Elapsed.TotalSeconds:F1}s");

        Assert.True(File.Exists(path), $"The new presentation was not written to {path}.");

        Assert.True(
            sw.Elapsed.TotalSeconds < 10,
            $"Dispose took {sw.Elapsed.TotalSeconds:F1}s for a newly created presentation, which means " +
            "it ran into one of the shutdown timeouts - see issue #148. Log:\n" +
            string.Join("\n", logger.Messages));

        Assert.NotNull(pid);
        Assert.True(HasExited(pid!.Value), $"PowerPoint process {pid} was still running after Dispose.");
    }

    private static bool HasExited(int pid)
    {
        try
        {
            using var process = System.Diagnostics.Process.GetProcessById(pid);
            return process.HasExited;
        }
        catch (ArgumentException)
        {
            return true;
        }
    }

    private sealed class CapturingLogger : ILogger<PptBatch>
    {
        public List<string> Messages { get; } = [];

        public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;

        public bool IsEnabled(LogLevel logLevel) => true;

        public void Log<TState>(
            LogLevel logLevel,
            EventId eventId,
            TState state,
            Exception? exception,
            Func<TState, Exception?, string> formatter)
        {
            lock (Messages)
            {
                Messages.Add($"{logLevel}: {formatter(state, exception)}");
            }
        }
    }
}
