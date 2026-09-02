// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Tests.Helpers;
using Xunit;

namespace PptMcp.Core.Tests.Integration;

/// <summary>
/// Integration coverage for the descriptive metadata returned by <c>slide list</c>
/// and <c>slide read</c> against real PowerPoint.
///
/// Regression guard for GitHub #126. Both operations populated their metadata with
/// inline COM member chains wrapped in swallowing catch blocks:
///
/// <code>
/// try { info.MasterName = slide.Design.SlideMaster.Name?.ToString() ?? ""; } catch { info.MasterName = ""; }
/// try { info.HasNotes = slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text?...; } catch { ... }
/// </code>
///
/// Two defects per line. The intermediate proxies were never bound to a local, so
/// <c>ComUtilities.Release</c> could not be called on them even in principle - the
/// six-hop notes chain abandoned six RCWs per slide, per call. And the catch
/// converted any genuine failure into an empty string, so a broken deck reported
/// the same result as a healthy one.
///
/// These tests pin the observable contract that survived the rewrite: layout and
/// master names are real values rather than the empty-string fallback, and the
/// optional notes/animation probes still return false rather than throwing when a
/// slide has no notes placeholder and no timeline.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Slide")]
[Trait("RequiresPowerPoint", "true")]
public sealed class SlideMetadataTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TempDirectoryFixture _fixture;
    private readonly SlideCommands _slides = new();

    public SlideMetadataTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void List_PopulatesLayoutAndMasterNames_NotEmptyFallback()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            var result = _slides.List(batch);

            Assert.True(result.Success);
            Assert.NotEmpty(result.Slides);

            var slide = result.Slides[^1];

            // The old code returned "" for both whenever the COM call threw, which
            // made a real fault indistinguishable from a slide with no layout.
            Assert.False(string.IsNullOrEmpty(slide.LayoutName));
            Assert.False(string.IsNullOrEmpty(slide.MasterName));
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void List_OnSlideWithoutNotesOrAnimations_ReportsFalseWithoutThrowing()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            var result = _slides.List(batch);

            Assert.True(result.Success);
            var slide = result.Slides[^1];

            // Genuinely optional (Rule 1b): a blank slide has no notes text and no
            // timeline entries. These must degrade to false, not throw.
            Assert.False(slide.HasNotes);
            Assert.False(slide.HasAnimations);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void Read_ReturnsSameMetadataAsList_ForTheSameSlide()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            var listed = _slides.List(batch);
            var index = listed.Slides.Count;
            var fromList = listed.Slides[^1];

            var detail = _slides.Read(batch, index);

            Assert.True(detail.Success);
            Assert.NotNull(detail.Slide);

            // List and Read now share PopulateSlideMetadata, so they cannot drift.
            Assert.Equal(fromList.LayoutName, detail.Slide!.LayoutName);
            Assert.Equal(fromList.MasterName, detail.Slide.MasterName);
            Assert.Equal(fromList.SlideId, detail.Slide.SlideId);
            Assert.Equal(fromList.SlideNumber, detail.Slide.SlideNumber);
            Assert.Equal(fromList.ShapeCount, detail.Slide.ShapeCount);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void Read_ShapeCountMatchesReturnedShapes()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");
            var index = _slides.List(batch).Slides.Count;

            var detail = _slides.Read(batch, index);

            Assert.True(detail.Success);
            Assert.NotNull(detail.Slide);
            Assert.Equal(detail.Slide!.ShapeCount, detail.Shapes.Count);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void List_RepeatedCalls_RemainStable()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            // Each call previously abandoned several RCWs per slide. Repeating the
            // call is the cheapest observable proxy for "the metadata read does not
            // corrupt or exhaust session state".
            var first = _slides.List(batch);
            var second = _slides.List(batch);
            var third = _slides.List(batch);

            Assert.True(first.Success);
            Assert.True(second.Success);
            Assert.True(third.Success);
            Assert.Equal(first.Slides.Count, third.Slides.Count);
            Assert.Equal(first.Slides[^1].MasterName, third.Slides[^1].MasterName);
            Assert.Equal(first.Slides[^1].LayoutName, third.Slides[^1].LayoutName);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    /// <summary>
    /// <c>slide read</c> populates the same metadata through a separate code path from
    /// <c>slide list</c>, so a fix applied to one does not imply the other (GitHub #133,
    /// item 5).
    /// </summary>
    [Fact]
    public void Read_PopulatesLayoutAndMasterNames_AndAgreesWithList()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            var listed = _slides.List(batch);
            Assert.True(listed.Success);
            var slideIndex = listed.Slides.Count;

            var read = _slides.Read(batch, slideIndex);

            Assert.True(read.Success, read.ErrorMessage);
            Assert.NotNull(read.Slide);
            Assert.False(string.IsNullOrEmpty(read.Slide.LayoutName));
            Assert.False(string.IsNullOrEmpty(read.Slide.MasterName));

            // The two paths describing the same slide differently would mean one of
            // them is reading the wrong object - invisible while only one is tested.
            Assert.Equal(listed.Slides[^1].LayoutName, read.Slide.LayoutName);
            Assert.Equal(listed.Slides[^1].MasterName, read.Slide.MasterName);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }
}
