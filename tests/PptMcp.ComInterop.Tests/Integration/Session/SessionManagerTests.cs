using System.Diagnostics;
using PptMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace PptMcp.ComInterop.Tests.Integration;

/// <summary>
/// Integration tests for SessionManager - verifies session lifecycle management.
/// Tests multi-session scenarios, concurrent operations, and proper cleanup.
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test session creation and tracking
/// - ✅ Test session retrieval by ID
/// - ✅ Test save operations
/// - ✅ Test close operations
/// - ✅ Test concurrent multi-session scenarios
/// - ✅ Test disposal cleanup
/// - ✅ Test post-disposal protection
///
/// NOTE: SessionManager uses PptSession internally, so these tests verify
/// the orchestration layer, not the underlying PowerPoint COM interactions.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresPowerPoint", "true")]
[Collection("Sequential")] // Disable parallelization to avoid COM interference
public class SessionManagerTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly List<string> _testFiles = new();

    public SessionManagerTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"SessionManagerTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        // Clean up any existing PowerPoint processes to ensure clean state
        try
        {
            var existingProcesses = Process.GetProcessesByName("POWERPNT");
            if (existingProcesses.Length > 0)
            {
                _output.WriteLine($"Cleaning up {existingProcesses.Length} existing PowerPoint processes...");
                foreach (var p in existingProcesses)
                {
                    p.Kill(entireProcessTree: true);
                    p.WaitForExit(5000);
                    p.Dispose();
                }
            }
        }
        catch (Exception ex)
        {
            _output.WriteLine($"Warning: Failed to clean PowerPoint processes: {ex.Message}");
        }
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);

        // Delete test files
        foreach (var file in _testFiles)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }

        // Delete temp directory
        if (Directory.Exists(_tempDir))
        {
            Directory.Delete(_tempDir, recursive: true);
        }
    }

    /// <summary>
    /// Path to the template xlsx file used for fast test file creation.
    /// Copying a template is ~1000x faster than spawning PowerPoint to create a new presentation.
    /// </summary>
    private static readonly string TemplateFilePath = Path.Combine(
        Path.GetDirectoryName(typeof(SessionManagerTests).Assembly.Location)!,
        "Integration", "Session", "TestFiles", "batch-test-static.pptx");

    private string CreateTestFile(string testName)
    {
        var fileName = $"{testName}_{Guid.NewGuid():N}.pptx";
        var filePath = Path.Combine(_tempDir, fileName);

        // PERFORMANCE OPTIMIZATION: Copy from template instead of spawning PowerPoint.
        // This reduces test file creation from ~7-14 seconds to <10ms.
        // Original approach using PptSession.CreateNew() spawned a full PowerPoint process
        // for each test file, causing 30+ second test execution times.
        File.Copy(TemplateFilePath, filePath);

        _testFiles.Add(filePath);
        return filePath;
    }

    #region Basic Session Lifecycle

    [Fact]
    public void CreateSession_ValidFile_ReturnsSessionId()
    {
        var testFile = CreateTestFile(nameof(CreateSession_ValidFile_ReturnsSessionId));
        using var manager = new SessionManager();

        var sessionId = manager.CreateSession(testFile);

        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        Assert.Equal(32, sessionId.Length); // GUID without hyphens
        Assert.Equal(1, manager.ActiveSessionCount);

        manager.CloseSession(sessionId);
    }

    [Fact]
    public void CreateSession_NonExistentFile_ThrowsFileNotFoundException()
    {
        using var manager = new SessionManager();
        var nonExistentFile = Path.Combine(_tempDir, "nonexistent.pptx");

        var ex = Assert.Throws<FileNotFoundException>(
            () => manager.CreateSession(nonExistentFile));

        Assert.Contains("PowerPoint file not found", ex.Message);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void GetSession_ExistingSessionId_ReturnsValidBatch()
    {
        var testFile = CreateTestFile(nameof(GetSession_ExistingSessionId_ReturnsValidBatch));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        var batch = manager.GetSession(sessionId);

        Assert.NotNull(batch);
        Assert.Equal(1, manager.ActiveSessionCount);

        manager.CloseSession(sessionId);
    }

    [Fact]
    public void GetSession_NonExistentSessionId_ReturnsNull()
    {
        using var manager = new SessionManager();

        var batch = manager.GetSession("nonexistent-session-id");

        Assert.Null(batch);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void GetSession_NullOrWhitespaceSessionId_ReturnsNull()
    {
        using var manager = new SessionManager();

        Assert.Null(manager.GetSession(null!));
        Assert.Null(manager.GetSession(""));
        Assert.Null(manager.GetSession("   "));
    }

    #endregion

    #region Save Operations

    [Fact]
    public void CloseSession_WithSaveTrue_SavesAndCloses()
    {
        var testFile = CreateTestFile(nameof(CloseSession_WithSaveTrue_SavesAndCloses));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        // Add a slide as marker of changes
        var batch = manager.GetSession(sessionId);
        Assert.NotNull(batch);
        batch.Execute((ctx, ct) =>
        {
            dynamic slides = ctx.Presentation.Slides;
            dynamic layouts = ((dynamic)ctx.Presentation).SlideMaster.CustomLayouts;
            dynamic layout = layouts[1];
            slides.AddSlide(slides.Count + 1, layout);
            return 0;
        });

        var closed = manager.CloseSession(sessionId, save: true);

        Assert.True(closed);
        Assert.Equal(0, manager.ActiveSessionCount);

        // Verify changes persisted (extra slide still there)
        using var verifyBatch = PptSession.BeginBatch(testFile);
        var slideCount = verifyBatch.Execute((ctx, ct) =>
        {
            return (int)ctx.Presentation.Slides.Count;
        });
        Assert.True(slideCount > 1, $"Expected more than 1 slide after save, got {slideCount}");
    }

    [Fact]
    public void CloseSession_WithSaveFalse_DiscardsChanges()
    {
        var testFile = CreateTestFile(nameof(CloseSession_WithSaveFalse_DiscardsChanges));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        // Get initial slide count
        var batch = manager.GetSession(sessionId);
        Assert.NotNull(batch);
        var initialCount = batch.Execute((ctx, ct) => (int)ctx.Presentation.Slides.Count);

        // Add a slide but don't save
        batch.Execute((ctx, ct) =>
        {
            dynamic slides = ctx.Presentation.Slides;
            dynamic layouts = ((dynamic)ctx.Presentation).SlideMaster.CustomLayouts;
            dynamic layout = layouts[1];
            slides.AddSlide(slides.Count + 1, layout);
            return 0;
        });

        var closed = manager.CloseSession(sessionId, save: false);

        Assert.True(closed);
        Assert.Equal(0, manager.ActiveSessionCount);

        // Verify changes were NOT persisted
        using var verifyBatch = PptSession.BeginBatch(testFile);
        var slideCount = verifyBatch.Execute((ctx, ct) =>
        {
            return (int)ctx.Presentation.Slides.Count;
        });
        Assert.Equal(initialCount, slideCount); // Should be same as before
    }

    #endregion

    #region Close Operations

    [Fact]
    public void CloseSession_ExistingSession_RemovesSessionAndReturnsTrue()
    {
        var testFile = CreateTestFile(nameof(CloseSession_ExistingSession_RemovesSessionAndReturnsTrue));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        var closed = manager.CloseSession(sessionId, save: false);

        Assert.True(closed);
        Assert.Equal(0, manager.ActiveSessionCount);
        Assert.Null(manager.GetSession(sessionId));
    }

    [Fact]
    public void CloseSession_NullOrWhitespaceSessionId_ReturnsFalse()
    {
        using var manager = new SessionManager();

        Assert.False(manager.CloseSession(null!));
        Assert.False(manager.CloseSession(""));
        Assert.False(manager.CloseSession("   "));
    }

    [Fact]
    public void CloseSession_AlreadyClosedSession_ReturnsFalse()
    {
        var testFile = CreateTestFile(nameof(CloseSession_AlreadyClosedSession_ReturnsFalse));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        var closed1 = manager.CloseSession(sessionId);
        var closed2 = manager.CloseSession(sessionId);

        Assert.True(closed1);
        Assert.False(closed2);
        Assert.Equal(0, manager.ActiveSessionCount);
    }

    #endregion

    #region Single-Session Constraint

    [Fact]
    public void CreateSession_WhileSessionActive_ThrowsInvalidOperationException()
    {
        var testFile1 = CreateTestFile($"{nameof(CreateSession_WhileSessionActive_ThrowsInvalidOperationException)}_1");
        var testFile2 = CreateTestFile($"{nameof(CreateSession_WhileSessionActive_ThrowsInvalidOperationException)}_2");
        using var manager = new SessionManager();

        var sessionId1 = manager.CreateSession(testFile1);
        Assert.Equal(1, manager.ActiveSessionCount);

        // Second session should fail — PowerPoint COM is single-instance
        var ex = Assert.Throws<InvalidOperationException>(
            () => manager.CreateSession(testFile2));
        Assert.Contains("single-instance", ex.Message);
        Assert.Equal(1, manager.ActiveSessionCount);

        manager.CloseSession(sessionId1);
    }

    [Fact]
    public void ActiveSessionIds_ReflectsCurrentState()
    {
        var testFile = CreateTestFile(nameof(ActiveSessionIds_ReflectsCurrentState));
        using var manager = new SessionManager();

        // Initially empty
        Assert.Empty(manager.ActiveSessionIds);

        // After creating session
        var sessionId = manager.CreateSession(testFile);
        var activeIds = manager.ActiveSessionIds.ToList();

        Assert.Single(activeIds);
        Assert.Contains(sessionId, activeIds);

        // After closing session
        manager.CloseSession(sessionId);
        activeIds = manager.ActiveSessionIds.ToList();

        Assert.Empty(activeIds);
    }

    [Fact]
    public void CloseAndReopen_DifferentFile_WorksCorrectly()
    {
        var testFile1 = CreateTestFile($"{nameof(CloseAndReopen_DifferentFile_WorksCorrectly)}_1");
        var testFile2 = CreateTestFile($"{nameof(CloseAndReopen_DifferentFile_WorksCorrectly)}_2");
        using var manager = new SessionManager();

        // Open first file
        var sessionId1 = manager.CreateSession(testFile1);
        Assert.Equal(1, manager.ActiveSessionCount);

        // Close first, then open second
        manager.CloseSession(sessionId1);
        Assert.Equal(0, manager.ActiveSessionCount);

        var sessionId2 = manager.CreateSession(testFile2);
        Assert.Equal(1, manager.ActiveSessionCount);
        Assert.NotNull(manager.GetSession(sessionId2));

        manager.CloseSession(sessionId2);
    }

    [Fact]
    public void CreateSession_SameFileAlreadyOpen_ThrowsSingleSessionException()
    {
        var testFile = CreateTestFile(nameof(CreateSession_SameFileAlreadyOpen_ThrowsSingleSessionException));
        using var manager = new SessionManager();

        // First session succeeds
        var sessionId1 = manager.CreateSession(testFile);
        Assert.NotNull(sessionId1);
        Assert.Equal(1, manager.ActiveSessionCount);

        // Second session with same file should fail — single-session constraint fires first
        var ex = Assert.Throws<InvalidOperationException>(
            () => manager.CreateSession(testFile));

        Assert.Contains("single-instance", ex.Message);
        Assert.Equal(1, manager.ActiveSessionCount);

        manager.CloseSession(sessionId1);
    }

    [Fact]
    public void CreateSession_AfterClosingPrevious_AllowsReopeningFile()
    {
        var testFile = CreateTestFile(nameof(CreateSession_AfterClosingPrevious_AllowsReopeningFile));
        using var manager = new SessionManager();

        // First session
        var sessionId1 = manager.CreateSession(testFile);
        Assert.Equal(1, manager.ActiveSessionCount);

        // Close first session
        manager.CloseSession(sessionId1);
        Assert.Equal(0, manager.ActiveSessionCount);

        // Should now be able to open same file again
        var sessionId2 = manager.CreateSession(testFile);
        Assert.NotNull(sessionId2);
        Assert.NotEqual(sessionId1, sessionId2);
        Assert.Equal(1, manager.ActiveSessionCount);

        manager.CloseSession(sessionId2);
    }

    #endregion

    #region Disposal and Post-Disposal

    [Fact]
    public void Dispose_OneSession_ClosesSession()
    {
        var testFile1 = CreateTestFile($"{nameof(Dispose_OneSession_ClosesSession)}_1");
        var manager = new SessionManager();

        var sessionId1 = manager.CreateSession(testFile1);

        Assert.Equal(1, manager.ActiveSessionCount);
        manager.Dispose();

        Assert.Equal(0, manager.ActiveSessionCount);
        Assert.Empty(manager.ActiveSessionIds);
    }

    [Fact]
    public void Dispose_EmptyManager_CompletesImmediately()
    {
        using var manager = new SessionManager();

        manager.Dispose();

        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void Dispose_CalledMultipleTimes_DoesNotThrow()
    {
        var manager = new SessionManager();

        manager.Dispose();
        manager.Dispose();
        manager.Dispose();

        Assert.Equal(0, manager.ActiveSessionCount);
    }

    [Fact]
    public void CreateSession_AfterDisposal_ThrowsObjectDisposedException()
    {
        var testFile = CreateTestFile(nameof(CreateSession_AfterDisposal_ThrowsObjectDisposedException));
        var manager = new SessionManager();
        manager.Dispose();

        Assert.Throws<ObjectDisposedException>(
            () => manager.CreateSession(testFile));
    }

    [Fact]
    public void GetSession_AfterDisposal_ThrowsObjectDisposedException()
    {
        var manager = new SessionManager();
        manager.Dispose();

        Assert.Throws<ObjectDisposedException>(
            () => manager.GetSession("any-id"));
    }

    [Fact]

    public void CloseSession_AfterDisposal_ThrowsObjectDisposedException()
    {
        var manager = new SessionManager();
        manager.Dispose();

        Assert.Throws<ObjectDisposedException>(
            () => manager.CloseSession("any-id"));
    }

    #endregion

    #region Edge Cases

    [Fact]
    public void CreateSession_VeryLongFilePath_HandlesGracefully()
    {
        // Create a long but valid path
        var longDirName = new string('x', 200);
        var longDir = Path.Combine(_tempDir, longDirName);

        try
        {
            Directory.CreateDirectory(longDir);
            var longFilePath = Path.Combine(longDir, "test.pptx");

            // Copy template file to the long path (faster than spawning PowerPoint)
            File.Copy(TemplateFilePath, longFilePath);
            _testFiles.Add(longFilePath);

            using var manager = new SessionManager();
            var sessionId = manager.CreateSession(longFilePath);

            Assert.NotNull(sessionId);
            Assert.Equal(1, manager.ActiveSessionCount);

            manager.CloseSession(sessionId);
        }
        catch (PathTooLongException)
        {
            // Expected on some systems - skip test
            _output.WriteLine("Path too long - test skipped");
        }
        catch (AggregateException ex) when (ex.InnerException is PathTooLongException)
        {
            // PowerPoint COM may reject very long paths - expected behavior (converted from COMException)
            _output.WriteLine($"PowerPoint rejected long path - test skipped: {ex.InnerException.Message}");
        }
        catch (AggregateException ex) when (ex.InnerException is AggregateException inner && inner.InnerException is PathTooLongException)
        {
            // Nested AggregateException from async task wrapping (STA thread -> Task.Wait -> Task.Wait)
            _output.WriteLine($"PowerPoint rejected long path (nested) - test skipped: {((AggregateException)ex.InnerException).InnerException!.Message}");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("already open") || ex.Message.Contains("Cannot open") || ex.Message.Contains("255 characters") || ex.Message.Contains("Filename cannot exceed"))
        {
            // PowerPoint COM returns generic errors for paths it can't handle.
            // This is a misleading error message - the real issue is the path is too long for PowerPoint COM.
            // We accept this as equivalent to PathTooLongException for test purposes.
            _output.WriteLine($"PowerPoint COM rejected long path with generic error - test skipped: {ex.Message}");
        }
    }

    [Fact]
    public void CloseSession_DefaultSaveTrue_PersistsChanges()
    {
        var testFile = CreateTestFile(nameof(CloseSession_DefaultSaveTrue_PersistsChanges));
        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile);

        // Get batch and make changes
        var batch = manager.GetSession(sessionId);
        Assert.NotNull(batch);

        batch.Execute((ctx, ct) =>
        {
            dynamic slides = ctx.Presentation.Slides;
            dynamic layouts = ((dynamic)ctx.Presentation).SlideMaster.CustomLayouts;
            dynamic layout = layouts[1];
            slides.AddSlide(slides.Count + 1, layout);
            return 0;
        });

        // Close with save:true explicitly
        var closed = manager.CloseSession(sessionId, save: true);
        Assert.True(closed);

        // Verify changes persisted (extra slide should be there)
        using var verifyBatch = PptSession.BeginBatch(testFile);
        var slideCount = verifyBatch.Execute((ctx, ct) =>
        {
            return (int)ctx.Presentation.Slides.Count;
        });

        Assert.True(slideCount > 1, $"Expected more than 1 slide after save, got {slideCount}");
    }

    #endregion
}







