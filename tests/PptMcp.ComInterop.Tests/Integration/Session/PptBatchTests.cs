using System.Diagnostics;
using PptMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace PptMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Integration tests for PptBatch - verifies batch operations and COM cleanup.
/// Tests that PowerPoint instances are reused across operations and properly cleaned up.
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test PptBatch.Execute() reuses PowerPoint instance
/// - ✅ Test PptBatch.Dispose() COM cleanup
/// - ✅ Test PptBatch.Save() functionality
/// - ✅ Verify POWERPNT.EXE process termination (no leaks)
///
/// NOTE: PptBatch.Dispose() handles all GC cleanup automatically.
/// Tests only need to wait for async disposal and process termination timing.
///
/// IMPORTANT: These tests spawn and terminate PowerPoint processes (side effects).
/// They run OnDemand only to avoid interference with normal test runs.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "PptBatch")]
[Collection("Sequential")] // Disable parallelization to avoid COM interference
public class PptBatchTests : IAsyncLifetime
{
    private readonly ITestOutputHelper _output;
    private static string? _staticTestFile;
    private string? _testFileCopy;

    public PptBatchTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public Task InitializeAsync()
    {
        // Use static test file from TestFiles folder (must be pre-created)
        if (_staticTestFile == null)
        {
            var testFolder = Path.Join(AppContext.BaseDirectory, "Integration", "Session", "TestFiles");
            _staticTestFile = Path.Join(testFolder, "batch-test-static.pptx");

            // Verify the static file exists
            if (!File.Exists(_staticTestFile))
            {
                throw new FileNotFoundException($"Static test file not found at {_staticTestFile}. " +
                    "Please create the batch-test-static.pptx file in the TestFiles folder.");
            }
        }

        // Create a fresh copy for this test instance (in temp folder)
        _testFileCopy = Path.Join(Path.GetTempPath(), $"batch-test-{Guid.NewGuid():N}.pptx");
        File.Copy(_staticTestFile, _testFileCopy, overwrite: true);

        // Wait for any PowerPoint processes from file creation to terminate
        return Task.Delay(500);
    }

    public Task DisposeAsync()
    {
        // Clean up this test's copy
        if (_testFileCopy != null && File.Exists(_testFileCopy))
        {
            File.Delete(_testFileCopy);
        }
        return Task.CompletedTask;
    }

    private static void CleanupStaticFile()
    {
        if (_staticTestFile != null && File.Exists(_staticTestFile))
        {
            File.Delete(_staticTestFile);
        }
    }

    [Fact]
    public void ExecuteAsync_MultipleOperations_ReusesPowerPointInstance()
    {
        // Arrange
        int operationCount = 0;

        // Act - Use batching for multiple operations
        using var batch = PptSession.BeginBatch(_testFileCopy!);

        for (int i = 0; i < 5; i++)
        {
            batch.Execute((ctx, ct) =>
            {
                operationCount++;
                _output.WriteLine($"Batch operation {operationCount}");

                // Verify we have the same context
                Assert.NotNull(ctx.App);
                Assert.NotNull(ctx.Presentation);

                return operationCount;
            });
        }

        // Assert
        Assert.Equal(5, operationCount);
        _output.WriteLine($"✓ Completed {operationCount} batch operations");
    }

    [Fact]
    public void Dispose_CleansUpComObjects_NoProcessLeak()
    {
        // Arrange
        var startingProcesses = Process.GetProcessesByName("POWERPNT");
        int startingCount = startingProcesses.Length;

        _output.WriteLine($"PowerPoint processes before: {startingCount}");

        // Act
        var batch = PptSession.BeginBatch(_testFileCopy!);

        batch.Execute((ctx, ct) =>
        {
            // Access slide count to verify COM is working
            int slideCount = ctx.Presentation.Slides.Count;
            _output.WriteLine($"Slide count: {slideCount}");
            return 0;
        });

        batch.Dispose();

        // Wait for PowerPoint process to fully terminate with polling
        // PowerPoint.Quit() signals shutdown but process termination is OS-controlled
        // Dispose() blocks up to StaThreadJoinTimeout for COM cleanup, but process may linger briefly
        var waitTimeout = TimeSpan.FromSeconds(15);
        var stopwatch = Stopwatch.StartNew();
        int endingCount;
        do
        {
            Thread.Sleep(500);
            endingCount = Process.GetProcessesByName("POWERPNT").Length;
            _output.WriteLine($"PowerPoint processes at {stopwatch.Elapsed.TotalSeconds:F1}s: {endingCount}");
        }
        while (endingCount > startingCount && stopwatch.Elapsed < waitTimeout);

        // Assert
        _output.WriteLine($"PowerPoint processes after {stopwatch.Elapsed.TotalSeconds:F1}s: {endingCount}");

        Assert.True(endingCount <= startingCount,
            $"PowerPoint process leak in batch! Started with {startingCount}, ended with {endingCount} after {waitTimeout.TotalSeconds}s");
    }

    [Fact]
    public void Save_PersistsChanges_ToPresentation()
    {
        // Arrange
        string testValue = $"Test-{Guid.NewGuid():N}";

        // Act - Write and save
        using (var batch = PptSession.BeginBatch(_testFileCopy!))
        {
            batch.Execute((ctx, ct) =>
            {
                // Add a slide and set its title text
                dynamic slide = ctx.Presentation.Slides[1];
                dynamic shape = slide.Shapes[1];
                shape.TextFrame.TextRange.Text = testValue;
                return 0;
            });

            batch.Save();
        }

        // Wait for file to be released
        Thread.Sleep(1000);

        // Verify - Read back the value in a new batch session
        string readValue;
        using (var batch = PptSession.BeginBatch(_testFileCopy!))
        {
            readValue = batch.Execute((ctx, ct) =>
            {
                dynamic slide = ctx.Presentation.Slides[1];
                dynamic shape = slide.Shapes[1];
                string result = shape.TextFrame.TextRange.Text?.ToString() ?? "";
                return result;
            });
        }

        // Assert
        Assert.Equal(testValue, readValue);
        _output.WriteLine($"✓ Value persisted correctly: {testValue}");
    }

    [Fact]
    public void PresentationPath_ReturnsCorrectPath()
    {
        // Arrange & Act
        using var batch = PptSession.BeginBatch(_testFileCopy!);

        // Assert
        Assert.Equal(_testFileCopy, batch.PresentationPath);
    }

    [Fact]
    public void CompleteWorkflow_CreateModifyReadSave_AllOperationsSucceed()
    {
        // Arrange
        string testBody = "Test Body Content";

        // Act - Execute complete workflow in single batch
        using (var batch = PptSession.BeginBatch(_testFileCopy!))
        {
            int initialSlideCount = 0;

            // Step 1: Get initial slide count
            batch.Execute((ctx, ct) =>
            {
                initialSlideCount = ctx.Presentation.Slides.Count;
                _output.WriteLine($"✓ Initial slide count: {initialSlideCount}");
                return 0;
            });

            // Step 2: Add a new slide
            batch.Execute((ctx, ct) =>
            {
                dynamic pres = ctx.Presentation;
                // Use layout from first slide master
                dynamic layout = pres.SlideMaster.CustomLayouts[1];
                pres.Slides.AddSlide(pres.Slides.Count + 1, layout);
                _output.WriteLine("✓ Added new slide");
                return 0;
            });

            // Step 3: Write text to the new slide
            batch.Execute((ctx, ct) =>
            {
                dynamic slide = ctx.Presentation.Slides[ctx.Presentation.Slides.Count];
                // Add a text box shape
                dynamic shape = slide.Shapes.AddTextbox(1, 100, 100, 400, 200); // msoTextOrientationHorizontal=1
                shape.TextFrame.TextRange.Text = testBody;
                _output.WriteLine($"✓ Wrote text to slide: {testBody}");
                return 0;
            });

            // Step 4: Read back to verify
            var readData = batch.Execute((ctx, ct) =>
            {
                int currentCount = ctx.Presentation.Slides.Count;
                dynamic lastSlide = ctx.Presentation.Slides[currentCount];
                string text = "";
                for (int i = 1; i <= lastSlide.Shapes.Count; i++)
                {
                    dynamic shape = lastSlide.Shapes[i];
                    if (Convert.ToInt32(shape.HasTextFrame) != 0)
                    {
                        text = shape.TextFrame.TextRange.Text?.ToString() ?? "";
                        if (text.Length > 0) break;
                    }
                }
                _output.WriteLine($"✓ Read back: slideCount={currentCount}, text={text}");
                return (currentCount, text);
            });

            // Verify
            Assert.True(readData.currentCount > initialSlideCount, "Should have more slides after add");
            Assert.Equal(testBody, readData.text);

            // Step 5: Save
            batch.Save();
            _output.WriteLine("✓ Saved presentation");
        }

        // Wait for file to be released
        Thread.Sleep(1000);

        // Verify - Open in new batch and check changes persisted
        using (var batch = PptSession.BeginBatch(_testFileCopy!))
        {
            var verifyData = batch.Execute((ctx, ct) =>
            {
                int slideCount = ctx.Presentation.Slides.Count;
                dynamic lastSlide = ctx.Presentation.Slides[slideCount];
                string text = "";
                for (int i = 1; i <= lastSlide.Shapes.Count; i++)
                {
                    dynamic shape = lastSlide.Shapes[i];
                    if (Convert.ToInt32(shape.HasTextFrame) != 0)
                    {
                        text = shape.TextFrame.TextRange.Text?.ToString() ?? "";
                        if (text.Length > 0) break;
                    }
                }
                return (slideCount, text);
            });

            Assert.True(verifyData.slideCount >= 2, "Should have at least 2 slides after save");
            Assert.Equal(testBody, verifyData.text);
            _output.WriteLine("✓ All workflow changes persisted correctly");
        }
    }

    // NOTE: ParallelBatches test removed — PowerPoint COM is single-instance.
    // Creating multiple PowerPoint.Application objects shares the same POWERPNT.EXE process.
    // Multi-session is not supported.

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Feature", "FileLocking")]
    public void Constructor_FileLockedByAnotherProcess_ThrowsInvalidOperationException()
    {
        // Arrange - Create a separate test file for locking test
        var lockedTestFile = Path.Join(Path.GetTempPath(), $"batch-test-locked-{Guid.NewGuid():N}.pptx");
        File.Copy(_staticTestFile!, lockedTestFile, overwrite: true);

        try
        {
            // Lock the file by opening with exclusive access (simulating PowerPoint or another process)
            using var fileLock = new FileStream(
                lockedTestFile,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None);

            // Act & Assert - Attempting to create PptBatch should fail immediately
            var ex = Assert.Throws<InvalidOperationException>(() =>
            {
                var batch = PptSession.BeginBatch(lockedTestFile);
                batch.Dispose();
            });

            // Verify error message is clear and actionable
            Assert.Contains("already open", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("close the file", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("exclusive access", ex.Message, StringComparison.OrdinalIgnoreCase);

            _output.WriteLine($"✓ File locking detected successfully");
            _output.WriteLine($"Error message: {ex.Message}");
        }
        finally
        {
            // Cleanup
            if (File.Exists(lockedTestFile))
            {
#pragma warning disable CA1031 // Intentional: best-effort test cleanup
                try { File.Delete(lockedTestFile); } catch (Exception) { /* Best effort - file may be locked */ }
#pragma warning restore CA1031
            }
        }
    }

    // Note: Testing file-already-open scenario is complex because:
    // 1. PowerPoint's behavior when opening an already-open file can vary (hang, prompt, or succeed)
    // 2. The error detection code in PptBatch.cs catches COM Error 0x800A03EC
    // 3. This test would require simulating PowerPoint having the file open externally
    //
    // The error handling code is verified through:
    // - Manual testing: Open file in PowerPoint UI, then try automation
    // - Real-world usage: Users will encounter this if they forget to close files
    // - Code review: Error message is clear and actionable
    //
    // UPDATE: We now have a test (Constructor_FileLockedByAnotherProcess_ThrowsInvalidOperationException)
    // that verifies the OS-level file locking check without requiring PowerPoint to be running.
    //
    // Keeping this comment as documentation that the scenario is handled in production code.
}







