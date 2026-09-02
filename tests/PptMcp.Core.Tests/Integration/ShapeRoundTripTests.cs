// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Shape;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Commands.Text;
using PptMcp.Core.Tests.Helpers;
using Xunit;

namespace PptMcp.Core.Tests.Integration;

/// <summary>
/// Round-trip integration coverage for the shape operations users actually run
/// (GitHub #133).
///
/// Before these tests, <c>shape add-shape</c> and <c>shape set-fill</c> were verified
/// only by hand. Nothing would have caught the two regressions that matter most:
/// mapping <c>autoShapeType</c> to the wrong <c>MsoAutoShapeType</c>, or writing the
/// wrong colour because PowerPoint stores RGB byte-reversed.
///
/// <para><b>Why two of these assert against raw COM rather than our own read path.</b></para>
///
/// A pure mutate-then-read-back test cannot detect a symmetric error. If
/// <c>SetFill</c> and <c>ReadFill</c> both inverted the byte order, the round trip
/// would agree with itself and stay green while every shape rendered the wrong
/// colour. So the colour and shape-type tests assert the value PowerPoint actually
/// holds, and a separate test covers the read path on top of it. The pair fails
/// distinguishably: a write bug breaks both, a read bug breaks only the second.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Shape")]
[Trait("RequiresPowerPoint", "true")]
public sealed class ShapeRoundTripTests : IClassFixture<TempDirectoryFixture>
{
    private const int MsoShapeOval = 9;
    private const int MsoShapeRectangle = 1;

    private readonly TempDirectoryFixture _fixture;
    private readonly SlideCommands _slides = new();
    private readonly ShapeCommands _shapes = new();
    private readonly TextCommands _text = new();

    public ShapeRoundTripTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void AddShape_StoresRequestedAutoShapeTypeAndBounds()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            var added = _shapes.AddShape(batch, slideIndex: 1, autoShapeType: MsoShapeOval,
                left: 100f, top: 120f, width: 200f, height: 80f);
            Assert.True(added.Success, added.ErrorMessage);

            var list = _shapes.List(batch, slideIndex: 1);
            Assert.True(list.Success, list.ErrorMessage);
            var shape = Assert.Single(list.Shapes);

            // Bounds are the part of the request most likely to be silently dropped:
            // AddShape takes four positional floats, so a transposed argument still
            // produces a valid shape.
            Assert.Equal(100f, shape.Left, 1);
            Assert.Equal(120f, shape.Top, 1);
            Assert.Equal(200f, shape.Width, 1);
            Assert.Equal(80f, shape.Height, 1);

            // ShapeInfo.ShapeType reports the MsoShapeType family ("AutoShape"), not
            // which auto shape it is, so the requested mapping is only observable on
            // the COM object itself.
            var actualAutoShapeType = batch.Execute((ctx, ct) =>
            {
                dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(1);
                dynamic comShape = slide.Shapes.Item(1);
                return Convert.ToInt32(comShape.AutoShapeType);
            });

            Assert.Equal(MsoShapeOval, actualAutoShapeType);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void SetFill_WritesRequestedColour_InPowerPointByteOrder()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");
            _shapes.AddShape(batch, 1, MsoShapeRectangle, 50f, 50f, 100f, 100f);

            var shapeName = _shapes.List(batch, 1).Shapes[0].Name;

            var filled = _shapes.SetFill(batch, 1, shapeName, "#112233");
            Assert.True(filled.Success, filled.ErrorMessage);

            // PowerPoint stores an OLE colour as 0x00BBGGRR, so #112233 must land as
            // 0x332211. Deliberately asymmetric: #FF0000 or any grey would pass even
            // if the byte order were reversed.
            var rgb = batch.Execute((ctx, ct) =>
            {
                dynamic slide = ((dynamic)ctx.Presentation).Slides.Item(1);
                dynamic comShape = slide.Shapes.Item(shapeName);
                return Convert.ToInt32(comShape.Fill.ForeColor.RGB);
            });

            Assert.Equal(0x332211, rgb);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void ReadFill_ReportsTheColourThatWasWritten()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");
            _shapes.AddShape(batch, 1, MsoShapeRectangle, 50f, 50f, 100f, 100f);

            var shapeName = _shapes.List(batch, 1).Shapes[0].Name;
            _shapes.SetFill(batch, 1, shapeName, "#112233");

            var read = _shapes.ReadFill(batch, 1, shapeName);

            Assert.True(read.Success, read.ErrorMessage);
            Assert.NotNull(read.Message);
            Assert.Contains("#112233", read.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Solid", read.Message, StringComparison.Ordinal);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void ShapeAndText_SurviveSaveAndReopen()
    {
        var testFile = _fixture.CreateTestFile();
        string shapeName;

        // Save is deliberate here: this is the persistence test required by #133, and
        // the only one in this file that pays for a save. An operation that succeeds
        // in memory and is lost on write looks identical to a working one until a
        // user reopens the deck.
        using (var manager = new SessionManager())
        {
            var sessionId = manager.CreateSession(testFile, show: false);
            try
            {
                var batch = manager.GetSession(sessionId)!;
                _slides.Create(batch, position: 0, layoutName: "Blank");
                _shapes.AddShape(batch, 1, MsoShapeRectangle, 60f, 70f, 150f, 90f);

                shapeName = _shapes.List(batch, 1).Shapes[0].Name;
                _shapes.SetFill(batch, 1, shapeName, "#112233");

                var written = _text.SetText(batch, 1, shapeName, "persisted content");
                Assert.True(written.Success, written.ErrorMessage);
            }
            finally
            {
                manager.CloseSession(sessionId, save: true);
            }
        }

        using (var manager = new SessionManager())
        {
            var sessionId = manager.CreateSession(testFile, show: false);
            try
            {
                var batch = manager.GetSession(sessionId)!;

                var list = _shapes.List(batch, 1);
                Assert.True(list.Success, list.ErrorMessage);
                var reopened = Assert.Single(list.Shapes, s => s.Name == shapeName);

                Assert.Equal(60f, reopened.Left, 1);
                Assert.Equal(70f, reopened.Top, 1);

                var text = _text.GetText(batch, 1, shapeName);
                Assert.True(text.Success, text.ErrorMessage);
                Assert.Equal("persisted content", text.Text.TrimEnd('\r', '\n'));

                var fill = _shapes.ReadFill(batch, 1, shapeName);
                Assert.True(fill.Success, fill.ErrorMessage);
                Assert.Contains("#112233", fill.Message!, StringComparison.OrdinalIgnoreCase);
            }
            finally
            {
                manager.CloseSession(sessionId, save: false);
            }
        }
    }
}
