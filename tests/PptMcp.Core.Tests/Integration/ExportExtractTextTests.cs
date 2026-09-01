// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Export;
using PptMcp.Core.Commands.Shape;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Commands.Text;
using PptMcp.Core.Tests.Helpers;
using Xunit;

namespace PptMcp.Core.Tests.Integration;

/// <summary>
/// Round-trip integration coverage for export operations against real PowerPoint.
///
/// Regression guard for GitHub #124: <c>ExtractText</c> read <c>Shape.HasTextFrame</c>
/// with a direct <c>(bool)</c> cast. That property is an <c>MsoTriState</c> (msoTrue = -1),
/// so unboxing it as a bool threw <see cref="InvalidCastException"/> and the operation
/// failed for every presentation containing at least one shape.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Export")]
[Trait("RequiresPowerPoint", "true")]
public sealed class ExportExtractTextTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TempDirectoryFixture _fixture;
    private readonly ExportCommands _export = new();
    private readonly ShapeCommands _shapes = new();
    private readonly SlideCommands _slides = new();
    private readonly TextCommands _text = new();

    public ExportExtractTextTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void ExtractText_WithShapeContainingText_WritesTextToFile()
    {
        var testFile = _fixture.CreateTestFile();
        var outputPath = Path.Combine(_fixture.TempDir, $"extract_{Guid.NewGuid():N}.txt");
        const string ExpectedText = "PptMcp extract-text regression";

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;

            _slides.Create(batch, position: 0, layoutName: "Blank");
            var slideIndex = _slides.List(batch).Slides.Count;

            // autoShapeType 1 = msoShapeRectangle
            var added = _shapes.AddShape(batch, slideIndex, 1, 100f, 100f, 300f, 100f);
            Assert.True(added.Success, added.ErrorMessage);

            var shapeName = _slides.Read(batch, slideIndex).Shapes[^1].Name;
            var setText = _text.SetText(batch, slideIndex, shapeName, ExpectedText);
            Assert.True(setText.Success, setText.ErrorMessage);

            // Before the fix this threw InvalidCastException on Shape.HasTextFrame.
            var result = _export.ExtractText(batch, outputPath);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.True(File.Exists(outputPath));
            Assert.Contains(ExpectedText, File.ReadAllText(outputPath), StringComparison.Ordinal);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false, force: true);
        }
    }
}
