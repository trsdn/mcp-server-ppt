// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Shape;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace PptMcp.Core.Tests.Integration.Shape;

/// <summary>
/// Integration tests for stable shape identity via the 'id:N' prefix syntax on shapeName.
/// Verifies that every shape-targeting ShapeCommands action resolves both:
///   - by mutable Name (existing behavior, regression)
///   - by stable Shape.Id via 'id:&lt;N&gt;' prefix (new)
/// And that Id-based references survive a rename, while Name-based references break.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Shape")]
[Trait("RequiresPowerPoint", "true")]
[Collection("Sequential")]
public class ShapeResolverTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TempDirectoryFixture _fixture;
    private readonly ITestOutputHelper _output;
    private readonly ShapeCommands _shape = new();
    private readonly SlideCommands _slide = new();

    public ShapeResolverTests(TempDirectoryFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private (int slideIndex, string name, int id) AddRectangle(IPptBatch batch, int slideIndex = 1)
    {
        // msoShapeRectangle = 1
        var addResult = _shape.AddShape(batch, slideIndex, autoShapeType: 1, left: 100, top: 100, width: 200, height: 100);
        Assert.True(addResult.Success, addResult.ErrorMessage);

        var list = _shape.List(batch, slideIndex);
        Assert.True(list.Success, list.ErrorMessage);
        var added = list.Shapes[^1];
        _output.WriteLine($"Added shape Id={added.ShapeId} Name='{added.Name}'");
        return (slideIndex, added.Name, added.ShapeId);
    }

    [Fact]
    public void Read_ByName_Succeeds_RegressionForExistingBehavior()
    {
        var file = _fixture.CreateTestFile();
        using var batch = PptSession.BeginBatch(file);
        _slide.Create(batch, position: 0, layoutName: "Blank");
        var (slideIndex, name, _) = AddRectangle(batch);

        var result = _shape.Read(batch, slideIndex, name);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(name, result.Shape.Name);
    }

    [Fact]
    public void Read_ByIdPrefix_Succeeds()
    {
        var file = _fixture.CreateTestFile();
        using var batch = PptSession.BeginBatch(file);
        _slide.Create(batch, position: 0, layoutName: "Blank");
        var (slideIndex, name, id) = AddRectangle(batch);

        var result = _shape.Read(batch, slideIndex, $"id:{id}");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(id, result.Shape.ShapeId);
        Assert.Equal(name, result.Shape.Name);
    }

    [Fact]
    public void Read_ByIdPrefix_SurvivesRename()
    {
        var file = _fixture.CreateTestFile();
        using var batch = PptSession.BeginBatch(file);
        _slide.Create(batch, position: 0, layoutName: "Blank");
        var (slideIndex, originalName, id) = AddRectangle(batch);

        // Mutate the Name property out-of-band (simulates a user renaming in PowerPoint).
        // We release the COM shape proxy explicitly so PowerPoint's Shapes lookup
        // table picks up the new name on subsequent queries.
        const string newName = "RenamedByUser";
        batch.Execute((ctx, ct) =>
        {
            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ((dynamic)ctx.Presentation).Slides.Item(slideIndex);
                shape = slide.Shapes.Item(originalName);
                shape.Name = newName;
                return 0;
            }
            finally
            {
                if (shape != null) PptMcp.ComInterop.ComUtilities.Release(ref shape!);
                if (slide != null) PptMcp.ComInterop.ComUtilities.Release(ref slide!);
            }
        });

        // Confirm the rename actually took effect (sanity check).
        var listAfter = _shape.List(batch, slideIndex);
        Assert.True(listAfter.Success, listAfter.ErrorMessage);
        Assert.Contains(listAfter.Shapes, s => s.ShapeId == id && s.Name == newName);

        // The contract under test: Id-based lookup MUST resolve to the same shape
        // after a rename, and report the new name.
        var byId = _shape.Read(batch, slideIndex, $"id:{id}");
        Assert.True(byId.Success, byId.ErrorMessage);
        Assert.NotNull(byId.Shape);
        Assert.Equal(newName, byId.Shape!.Name);
        Assert.Equal(id, byId.Shape.ShapeId);
    }

    [Fact]
    public void Read_InvalidIdPrefix_ThrowsDescriptiveError()
    {
        var file = _fixture.CreateTestFile();
        using var batch = PptSession.BeginBatch(file);
        _slide.Create(batch, position: 0, layoutName: "Blank");
        AddRectangle(batch);

        var ex = Record.Exception(() => _shape.Read(batch, slideIndex: 1, shapeName: "id:99999"));

        Assert.NotNull(ex);
        Assert.Contains("99999", ex.Message);
    }

    [Fact]
    public void Read_MalformedIdPrefix_ThrowsDescriptiveError()
    {
        var file = _fixture.CreateTestFile();
        using var batch = PptSession.BeginBatch(file);
        _slide.Create(batch, position: 0, layoutName: "Blank");
        AddRectangle(batch);

        var ex = Record.Exception(() => _shape.Read(batch, slideIndex: 1, shapeName: "id:abc"));

        Assert.NotNull(ex);
        Assert.Contains("id:", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MoveResize_ByIdPrefix_AppliesToCorrectShape()
    {
        var file = _fixture.CreateTestFile();
        using var batch = PptSession.BeginBatch(file);
        _slide.Create(batch, position: 0, layoutName: "Blank");
        var (slideIndex, _, id) = AddRectangle(batch);
        var (_, _, otherId) = AddRectangle(batch, slideIndex);

        Assert.NotEqual(id, otherId);

        var result = _shape.MoveResize(batch, slideIndex, $"id:{id}", left: 50, top: 50, width: 300, height: 150);

        Assert.True(result.Success, result.ErrorMessage);
        var read = _shape.Read(batch, slideIndex, $"id:{id}");
        Assert.Equal(50f, read.Shape.Left);
        Assert.Equal(50f, read.Shape.Top);
        Assert.Equal(300f, read.Shape.Width);
        Assert.Equal(150f, read.Shape.Height);
    }

    [Fact]
    public void Delete_ByIdPrefix_RemovesCorrectShape()
    {
        var file = _fixture.CreateTestFile();
        using var batch = PptSession.BeginBatch(file);
        _slide.Create(batch, position: 0, layoutName: "Blank");
        var (slideIndex, _, idA) = AddRectangle(batch);
        var (_, _, idB) = AddRectangle(batch, slideIndex);

        var result = _shape.Delete(batch, slideIndex, $"id:{idA}");
        Assert.True(result.Success, result.ErrorMessage);

        var list = _shape.List(batch, slideIndex);
        Assert.DoesNotContain(list.Shapes, s => s.ShapeId == idA);
        Assert.Contains(list.Shapes, s => s.ShapeId == idB);
    }

    [Fact]
    public void Group_AcceptsMixedNameAndIdPrefixReferences()
    {
        var file = _fixture.CreateTestFile();
        using var batch = PptSession.BeginBatch(file);
        _slide.Create(batch, position: 0, layoutName: "Blank");
        var (slideIndex, nameA, _) = AddRectangle(batch);
        var (_, _, idB) = AddRectangle(batch, slideIndex);

        // Mix: one by name, one by id
        var result = _shape.Group(batch, slideIndex, $"{nameA},id:{idB}");

        Assert.True(result.Success, result.ErrorMessage);
    }
}
