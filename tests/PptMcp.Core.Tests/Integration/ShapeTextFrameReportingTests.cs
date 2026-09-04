// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Placeholder;
using PptMcp.Core.Commands.Shape;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Commands.SlideTable;
using PptMcp.Core.Commands.Text;
using PptMcp.Core.Tests.Helpers;
using Xunit;

namespace PptMcp.Core.Tests.Integration;

/// <summary>
/// Coverage for how <c>HasTextFrame</c> and <c>Text</c> are reported (GitHub #126).
///
/// Two places built a shape or placeholder description with the same guard:
///
/// <code>catch { info.HasTextFrame = false; }</code>
///
/// That does not report a failure — it reports a *fact*, and a false one. A caller that
/// checks <c>HasTextFrame</c> before writing text will skip a shape that does have a text
/// frame, and will do so silently, because the enclosing result still carries
/// <c>Success = true</c>. The read that failed and the shape that genuinely has no text
/// frame become indistinguishable.
///
/// <para><b>Both directions are asserted deliberately.</b></para>
///
/// Testing only that a textbox reports <c>true</c> would still pass if the code answered
/// <c>true</c> unconditionally; testing only that a line reports <c>false</c> would pass
/// against the very bug being removed, since the catch also answers <c>false</c>. Only the
/// pair pins the behaviour.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Shape")]
[Trait("RequiresPowerPoint", "true")]
public sealed class ShapeTextFrameReportingTests : IClassFixture<TempDirectoryFixture>
{
    private const int MsoShapeRectangle = 1;

    private readonly TempDirectoryFixture _fixture;
    private readonly SlideCommands _slides = new();
    private readonly ShapeCommands _shapes = new();
    private readonly TextCommands _text = new();
    private readonly SlideTableCommands _tables = new();
    private readonly PlaceholderCommands _placeholders = new();

    public ShapeTextFrameReportingTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void List_ReportsHasTextFrameAndTextForAShapeThatHasThem()
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

            var list = _shapes.List(batch, slideIndex: 1);
            Assert.True(list.Success, list.ErrorMessage);
            var shapeName = Assert.Single(list.Shapes).Name;

            var setText = _text.SetText(batch, slideIndex: 1, shapeName: shapeName, text: "Frame content");
            Assert.True(setText.Success, setText.ErrorMessage);

            var reread = _shapes.List(batch, slideIndex: 1);
            Assert.True(reread.Success, reread.ErrorMessage);
            var shape = Assert.Single(reread.Shapes);

            // The swallowed catch answered false here whenever the read failed, so a
            // caller would have skipped a shape it could perfectly well write to.
            Assert.True(shape.HasTextFrame);
            Assert.Equal("Frame content", shape.Text);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void List_ReportsNoTextFrameForAShapeThatGenuinelyHasNone()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            // A table shape is the clean negative case: PowerPoint reports msoFalse for
            // HasTextFrame and msoTrue for HasTable, so both neighbouring reads are
            // covered by one shape.
            var table = _tables.Create(batch, slideIndex: 1, rows: 2, columns: 2,
                left: 50f, top: 50f, width: 300f, height: 100f);
            Assert.True(table.Success, table.ErrorMessage);

            var list = _shapes.List(batch, slideIndex: 1);
            Assert.True(list.Success, list.ErrorMessage);
            var shape = Assert.Single(list.Shapes);

            // Removing the catch does not change this answer - which is the point. The
            // catch was never needed for legitimate shapes; it only masked genuine
            // read failures behind an identical-looking false.
            Assert.False(shape.HasTextFrame);
            Assert.True(shape.HasTable);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void PlaceholderList_ReportsHasTextFrameAndTextForATitlePlaceholder()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Title and Content");

            var list = _placeholders.List(batch, slideIndex: 1);
            Assert.True(list.Success, list.ErrorMessage);
            Assert.NotEmpty(list.Placeholders);

            var titlePlaceholder = list.Placeholders[0];
            var setText = _text.SetText(batch, slideIndex: 1,
                shapeName: titlePlaceholder.Name, text: "Placeholder content");
            Assert.True(setText.Success, setText.ErrorMessage);

            var reread = _placeholders.List(batch, slideIndex: 1);
            Assert.True(reread.Success, reread.ErrorMessage);
            var placeholder = reread.Placeholders[0];

            Assert.True(placeholder.HasTextFrame);
            Assert.Equal("Placeholder content", placeholder.Text);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void PlaceholderList_ReportsEveryPlaceholderOnTheSlide()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Title and Content");

            var list = _placeholders.List(batch, slideIndex: 1);
            Assert.True(list.Success, list.ErrorMessage);

            // Every placeholder must carry a usable name and a resolved type name. A
            // partially-populated entry is the shape of failure the catch allowed.
            Assert.All(list.Placeholders, p =>
            {
                Assert.False(string.IsNullOrWhiteSpace(p.Name));
                Assert.False(string.IsNullOrWhiteSpace(p.PlaceholderTypeName));
            });
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }
}
