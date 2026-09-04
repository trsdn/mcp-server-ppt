// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Background;
using PptMcp.Core.Commands.Media;
using PptMcp.Core.Commands.Shape;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Commands.Slideshow;
using PptMcp.Core.Commands.Text;
using PptMcp.Core.Tests.Helpers;
using Xunit;

namespace PptMcp.Core.Tests.Integration;

/// <summary>
/// Coverage for the last four swallowing catch blocks outside the text formatter
/// (GitHub #126). Each one produced a plausible wrong answer under a
/// <c>Success = true</c> result, and each had no behavioural test - only reflection
/// tests asserting that the action enums were wide enough.
///
/// <para><b>Two of these catches are genuine existence probes and are kept.</b></para>
///
/// <c>Presentation.SlideShowWindow</c> throws when no slideshow is running; COM offers
/// no <c>IsSlideShowRunning</c>, so the throw <i>is</i> the query. What was wrong was
/// the catch's <i>width</i>: it wrapped the probe together with the work that follows
/// it, so a failure during <c>View.Exit()</c> or while reading
/// <c>CurrentShowPosition</c> was absorbed into "no slideshow was running". The probes
/// are now narrowed to the acquisition alone, and these tests pin the not-running
/// answers that the probes must keep producing.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Background")]
[Trait("RequiresPowerPoint", "true")]
public sealed class SwallowedCatchBehaviourTests : IClassFixture<TempDirectoryFixture>
{
    private const int MsoShapeRectangle = 1;

    private readonly TempDirectoryFixture _fixture;
    private readonly SlideCommands _slides = new();
    private readonly ShapeCommands _shapes = new();
    private readonly BackgroundCommands _background = new();
    private readonly SlideshowCommands _slideshow = new();
    private readonly MediaCommands _media = new();
    private readonly TextCommands _text = new();

    public SwallowedCatchBehaviourTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void BackgroundGetInfo_ReportsTheColourThatWasSet_NotUnknown()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            var set = _background.SetColor(batch, slideIndex: 1, colorHex: "#0B3D91");
            Assert.True(set.Success, set.ErrorMessage);

            var info = _background.GetInfo(batch, slideIndex: 1);
            Assert.True(info.Success, info.ErrorMessage);

            // The removed catch answered FillType = "Unknown" on any read failure, while
            // still reporting Success = true. "Unknown" is a plausible-looking value, so
            // nothing downstream could tell it apart from a real answer.
            Assert.False(info.FollowMasterBackground);
            Assert.Equal("Solid", info.FillType);
            Assert.Equal("#0B3D91", info.Color);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void BackgroundGetInfo_ReportsMasterWhenTheSlideFollowsIt()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            var info = _background.GetInfo(batch, slideIndex: 1);
            Assert.True(info.Success, info.ErrorMessage);

            // The other direction: a fresh slide inherits from the master, and that path
            // never enters the guarded block at all. Asserting both keeps "Solid" from
            // being satisfied by a constant.
            Assert.True(info.FollowMasterBackground);
            Assert.Equal("Master", info.FillType);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void SlideshowGetStatus_ReportsNotRunningWhenNoShowIsActive()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            var status = _slideshow.GetStatus(batch);
            Assert.True(status.Success, status.ErrorMessage);

            // This is the answer the narrowed existence probe must keep producing.
            Assert.False(status.IsRunning);
            Assert.Equal(0, status.CurrentSlide);

            // TotalSlides is read outside the probe entirely, so it must be real. A probe
            // widened back over the whole method would leave this at 0.
            Assert.True(status.TotalSlides > 0, $"expected slides, got {status.TotalSlides}");
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void SlideshowEndShow_SucceedsWhenNoShowIsRunning()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            var stop = _slideshow.EndShow(batch);

            // Stopping a show that is not running is not an error - it is the requested
            // end state. The narrowed probe preserves that.
            Assert.True(stop.Success, stop.ErrorMessage);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void MediaGetInfo_FailsOnAShapeThatIsNotMedia()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            var added = _shapes.AddShape(batch, slideIndex: 1, autoShapeType: MsoShapeRectangle,
                left: 100f, top: 100f, width: 200f, height: 80f);
            Assert.True(added.Success, added.ErrorMessage);

            var shapeName = Assert.Single(_shapes.List(batch, slideIndex: 1).Shapes).Name;

            // shape.MediaType is only valid on a media shape. The removed catch turned
            // that failure into MediaType = "Unknown" under Success = true, so asking for
            // media details about a rectangle produced a confident, wrong answer instead
            // of an error.
            var ex = Assert.Throws<InvalidOperationException>(
                () => _media.GetInfo(batch, slideIndex: 1, shapeName: shapeName));

            // The message must name the shape, so the caller can act on it. The original
            // COM failure is preserved as the inner exception.
            Assert.Contains(shapeName, ex.Message);
            Assert.NotNull(ex.InnerException);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void GetText_ReportsRunColourInRgbOrderNotPowerPointsByteOrder()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            var added = _shapes.AddShape(batch, slideIndex: 1, autoShapeType: MsoShapeRectangle,
                left: 100f, top: 100f, width: 240f, height: 90f);
            Assert.True(added.Success, added.ErrorMessage);

            var shapeName = Assert.Single(_shapes.List(batch, slideIndex: 1).Shapes).Name;

            Assert.True(_text.SetText(batch, 1, shapeName, "Coloured run").Success);
            var format = _text.Format(batch, slideIndex: 1, shapeName: shapeName,
                fontName: null, fontSize: null, bold: null, italic: null,
                color: "#0B3D91", alignment: null, verticalAlignment: null);
            Assert.True(format.Success, format.ErrorMessage);

            var read = _text.GetText(batch, slideIndex: 1, shapeName: shapeName);
            Assert.True(read.Success, read.ErrorMessage);

            var run = read.Paragraphs.SelectMany(p => p.Runs).First();

            // A deliberately asymmetric colour: #0B3D91 byte-reversed is #913D0B, so a
            // read path that formats PowerPoint's raw 0x00BBGGRR value straight to hex
            // produces a well-formed but wrong answer rather than an error.
            Assert.Equal("#0B3D91", run.Color);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }
}
