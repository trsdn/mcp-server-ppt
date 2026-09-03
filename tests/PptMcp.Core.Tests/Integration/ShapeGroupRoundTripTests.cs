// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Shape;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Tests.Helpers;
using Xunit;

namespace PptMcp.Core.Tests.Integration;

/// <summary>
/// Integration coverage for grouped-shape enumeration (GitHub #126).
///
/// <c>ShapeHelpers.ReadShapeInfo</c> walks <c>Shape.GroupItems</c> to populate
/// <c>ShapeInfo.GroupItems</c>. That walk used to sit inside a catch-all, so a failure
/// part-way through returned a group reporting fewer children than it has - or none -
/// while still reporting success. Nothing distinguished that from a genuinely small
/// group, which is why the catch could not simply be deleted without coverage first:
/// there was no test asserting what the correct child list even is.
///
/// These tests pin the correct answer, so the enumeration is now allowed to fail loudly.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Shape")]
[Trait("RequiresPowerPoint", "true")]
public sealed class ShapeGroupRoundTripTests : IClassFixture<TempDirectoryFixture>
{
    // MsoAutoShapeType.msoShapeRectangle
    private const int MsoShapeRectangle = 1;

    // MsoShapeType.msoGroup, the value ReadShapeInfo tests for.
    private const int MsoGroup = 6;

    private readonly TempDirectoryFixture _fixture;
    private readonly SlideCommands _slides = new();
    private readonly ShapeCommands _shapes = new();

    public ShapeGroupRoundTripTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void List_GroupedShape_ReportsEveryChild()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            // Three children rather than two: a truncated walk that stopped after the
            // first child would still satisfy a two-child assertion if the count were
            // read as 1, so the extra shape makes an off-by-one visible.
            var names = AddRectangles(batch, count: 3);

            var grouped = _shapes.Group(batch, slideIndex: 1, shapeNames: string.Join(",", names));
            Assert.True(grouped.Success, grouped.ErrorMessage);

            var list = _shapes.List(batch, slideIndex: 1);
            Assert.True(list.Success, list.ErrorMessage);

            var group = Assert.Single(list.Shapes, s => s.IsGroup);
            Assert.Equal(MsoGroup, ShapeTypeNameToMso(group.ShapeType));

            Assert.NotNull(group.GroupItems);
            Assert.Equal(3, group.GroupItems!.Count);

            // Every child must be present by name. Asserting the set rather than the
            // count catches a walk that enumerated the right number of slots but read
            // the same item repeatedly.
            var childNames = group.GroupItems.Select(c => c.Name).OrderBy(n => n).ToArray();
            Assert.Equal(names.OrderBy(n => n).ToArray(), childNames);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void List_UngroupedShape_HasNoGroupItems()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");
            AddRectangles(batch, count: 1);

            var list = _shapes.List(batch, slideIndex: 1);
            Assert.True(list.Success, list.ErrorMessage);

            var shape = Assert.Single(list.Shapes);
            Assert.False(shape.IsGroup);
            Assert.Null(shape.GroupItems);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    private string[] AddRectangles(PptMcp.ComInterop.Session.IPptBatch batch, int count)
    {
        for (int i = 0; i < count; i++)
        {
            var added = _shapes.AddShape(batch, slideIndex: 1, autoShapeType: MsoShapeRectangle,
                left: 40f + (i * 90f), top: 60f, width: 80f, height: 50f);
            Assert.True(added.Success, added.ErrorMessage);
        }

        // PowerPoint names the shapes itself, so read them back rather than assuming.
        var list = _shapes.List(batch, slideIndex: 1);
        Assert.True(list.Success, list.ErrorMessage);
        Assert.Equal(count, list.Shapes.Count);

        return list.Shapes.Select(s => s.Name).ToArray();
    }

    private static int ShapeTypeNameToMso(string? shapeTypeName) =>
        shapeTypeName == "Group" ? MsoGroup : -1;
}
