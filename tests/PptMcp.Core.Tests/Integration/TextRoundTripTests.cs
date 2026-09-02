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
/// Round-trip integration coverage for <c>text set</c> and <c>text get</c> (GitHub #133).
///
/// The failure this guards against is a silent no-op: <c>SetText</c> returning
/// <c>Success = true</c> having written nothing, or having written to a different
/// shape than the one named. Neither is visible to a success-flag assertion, which
/// is why #133 calls those out as insufficient.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Text")]
[Trait("RequiresPowerPoint", "true")]
public sealed class TextRoundTripTests : IClassFixture<TempDirectoryFixture>
{
    private const int MsoShapeRectangle = 1;

    private readonly TempDirectoryFixture _fixture;
    private readonly SlideCommands _slides = new();
    private readonly ShapeCommands _shapes = new();
    private readonly TextCommands _text = new();

    public TextRoundTripTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void SetText_ThenGetText_ReturnsExactlyWhatWasWritten()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");
            _shapes.AddShape(batch, 1, MsoShapeRectangle, 40f, 40f, 300f, 120f);

            var shapeName = _shapes.List(batch, 1).Shapes[0].Name;

            var written = _text.SetText(batch, 1, shapeName, "Round trip 123");
            Assert.True(written.Success, written.ErrorMessage);

            var read = _text.GetText(batch, 1, shapeName);

            Assert.True(read.Success, read.ErrorMessage);
            Assert.Equal(shapeName, read.ShapeName);

            // PowerPoint appends a paragraph terminator to TextRange.Text; the content
            // either side of it must be byte-identical to what was requested.
            Assert.Equal("Round trip 123", read.Text.TrimEnd('\r', '\n'));
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void SetText_WritesOnlyToTheNamedShape()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");
            _shapes.AddShape(batch, 1, MsoShapeRectangle, 20f, 20f, 200f, 60f);
            _shapes.AddShape(batch, 1, MsoShapeRectangle, 20f, 200f, 200f, 60f);

            var shapes = _shapes.List(batch, 1).Shapes;
            Assert.Equal(2, shapes.Count);

            var target = shapes[1].Name;
            var other = shapes[0].Name;
            Assert.NotEqual(target, other);

            _text.SetText(batch, 1, target, "only here");

            Assert.Equal("only here", _text.GetText(batch, 1, target).Text.TrimEnd('\r', '\n'));

            // Writing to the wrong shape is indistinguishable from writing to the right
            // one unless the untouched shape is checked as well.
            Assert.Equal(string.Empty, _text.GetText(batch, 1, other).Text.TrimEnd('\r', '\n'));
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void SetText_ReplacesExistingContentRatherThanAppending()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");
            _shapes.AddShape(batch, 1, MsoShapeRectangle, 40f, 40f, 300f, 120f);

            var shapeName = _shapes.List(batch, 1).Shapes[0].Name;

            _text.SetText(batch, 1, shapeName, "first value");
            _text.SetText(batch, 1, shapeName, "second value");

            // The documented contract is "replaces all existing text".
            var read = _text.GetText(batch, 1, shapeName);
            Assert.Equal("second value", read.Text.TrimEnd('\r', '\n'));
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }
}
