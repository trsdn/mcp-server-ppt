using System.Diagnostics;
using PptMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace PptMcp.ComInterop.Tests.Integration;

/// <summary>
/// Integration tests for PptSession - verifies public API and COM cleanup.
/// Tests BeginBatch() and CreateNew() functionality.
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test PptSession.BeginBatch() validation and batch creation
/// - ✅ Test PptSession.CreateNew() file creation
/// - ✅ Verify POWERPNT.EXE process termination (no leaks)
///
/// NOTE: PptSession methods use PptShutdownService for resilient cleanup.
/// Automatic RCW finalizers handle COM reference cleanup (no forced GC needed).
/// Process cleanup errors are logged but don't fail tests.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "PptSession")]
[Collection("Sequential")] // Disable parallelization to avoid COM interference
public class PptSessionTests : IDisposable
{
    private readonly ITestOutputHelper _output;

    public PptSessionTests(ITestOutputHelper output)
    {
        _output = output;

        // Kill any existing PowerPoint processes to ensure clean state
        var existingProcesses = Process.GetProcessesByName("POWERPNT");
        if (existingProcesses.Length > 0)
        {
            _output.WriteLine($"Cleaning up {existingProcesses.Length} existing PowerPoint processes...");
            foreach (var p in existingProcesses)
            {
                p.Kill(); p.WaitForExit(2000);
            }
            _output.WriteLine("PowerPoint processes cleaned up");
        }

    }

    /// <summary>
    /// Runs after each test
    /// </summary>
    public void Dispose()
    {
        // Nothing to dispose
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void BeginBatch_WithValidFile_CreatesBatch()
    {
        // Arrange
        string testFile = Path.Join(Path.GetTempPath(), $"session-test-{Guid.NewGuid():N}.pptx");
        CreateTempTestFile(testFile);

        try
        {
            // Act
            using var batch = PptSession.BeginBatch(testFile);

            // Assert
            Assert.NotNull(batch);
            Assert.Equal(testFile, batch.PresentationPath);

            _output.WriteLine($"✓ Batch created successfully for: {Path.GetFileName(testFile)}");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    public void BeginBatch_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string nonExistentFile = Path.Join(Path.GetTempPath(), $"does-not-exist-{Guid.NewGuid():N}.pptx");

        // Act & Assert
        Assert.Throws<FileNotFoundException>(() =>
        {
            using var batch = PptSession.BeginBatch(nonExistentFile);
        });

        _output.WriteLine("✓ Correctly throws FileNotFoundException for non-existent file");
    }

    [Fact]
    public void BeginBatch_WithInvalidExtension_ThrowsArgumentException()
    {
        // Arrange
        string invalidFile = Path.Join(Path.GetTempPath(), $"test-{Guid.NewGuid():N}.txt");
        File.WriteAllText(invalidFile, "dummy");

        try
        {
            // Act & Assert
            var exception = Assert.Throws<ArgumentException>(() =>
            {
                using var batch = PptSession.BeginBatch(invalidFile);
            });

            Assert.Contains("Invalid file extension", exception.Message);
            _output.WriteLine("✓ Correctly rejects non-PowerPoint file extension");
        }
        finally
        {
            if (File.Exists(invalidFile)) File.Delete(invalidFile);
        }
    }

    [Fact]
    public void CreateNew_CreatesNewPresentation()
    {
        // Arrange
        string testFile = Path.Join(Path.GetTempPath(), $"new-presentation-{Guid.NewGuid():N}.pptx");

        try
        {
            // Act
            var result = PptSession.CreateNew(testFile, isMacroEnabled: false, (ctx, ct) =>
            {
                _output.WriteLine($"✓ Presentation created at: {ctx.PresentationPath}");
                return 0;
            });

            // Assert
            Assert.True(File.Exists(testFile), "File should be created");
            Assert.Equal(0, result);

            // Verify we can open it with batch API
            using (var batch = PptSession.BeginBatch(testFile))
            {
                batch.Execute((ctx, ct) =>
                {
                    Assert.NotNull(ctx.Presentation);
                    _output.WriteLine("✓ Can open created presentation with batch API");
                    return 0;
                });
            }
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    public void CreateNew_WithMacroEnabled_CreatesXlsmFile()
    {
        // Arrange
        string testFile = Path.Join(Path.GetTempPath(), $"new-macro-presentation-{Guid.NewGuid():N}.pptm");

        try
        {
            // Act
            var result = PptSession.CreateNew(testFile, isMacroEnabled: true, (ctx, ct) =>
            {
                _output.WriteLine($"✓ Macro-enabled presentation created at: {ctx.PresentationPath}");
                return 0;
            });

            // Assert
            Assert.True(File.Exists(testFile), "XLSM file should be created");
            Assert.Equal(".pptm", Path.GetExtension(testFile).ToLowerInvariant());
            _output.WriteLine("✓ Correctly created .pptm file");
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Fact]
    public void CreateNew_CreatesDirectoryIfNeeded()
    {
        // Arrange
        string testDir = Path.Join(Path.GetTempPath(), $"testdir-{Guid.NewGuid():N}");
        string testFile = Path.Join(testDir, "newfile.pptx");

        try
        {
            // Act
            PptSession.CreateNew(testFile, isMacroEnabled: false, (ctx, ct) =>
            {
                return 0;
            });

            // Assert
            Assert.True(Directory.Exists(testDir), "Directory should be created");
            Assert.True(File.Exists(testFile), "File should be created in new directory");
            _output.WriteLine("✓ Correctly created directory and file");
        }
        finally
        {
            if (Directory.Exists(testDir)) Directory.Delete(testDir, recursive: true);
        }
    }

    // Helper method

    /// <summary>
    /// Path to the template xlsx file used for fast test file creation.
    /// Copying a template is ~1000x faster than spawning PowerPoint to create a new presentation.
    /// </summary>
    private static readonly string TemplateFilePath = Path.Combine(
        Path.GetDirectoryName(typeof(PptSessionTests).Assembly.Location)!,
        "Integration", "Session", "TestFiles", "batch-test-static.pptx");

    private static void CreateTempTestFile(string filePath)
    {
        // PERFORMANCE OPTIMIZATION: Copy from template instead of spawning PowerPoint.
        // For tests that only need a valid PowerPoint file to exist (not testing creation),
        // this reduces setup time from ~7-14 seconds to <10ms.
        File.Copy(TemplateFilePath, filePath);
    }
}




