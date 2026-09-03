// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Chart;
using PptMcp.Core.Commands.Shape;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Tests.Helpers;
using Xunit;

namespace PptMcp.Core.Tests.Integration;

/// <summary>
/// Round-trip integration coverage for <c>chart create</c> (GitHub #133).
///
/// <c>Create</c> takes an <c>XlChartType</c> integer and four positional floats, so
/// the two things that can go wrong without any visible error are producing a chart
/// of the wrong type and placing it at the wrong bounds. Both are asserted here
/// against the created shape rather than against the success flag.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Chart")]
[Trait("RequiresPowerPoint", "true")]
public sealed class ChartRoundTripTests : IClassFixture<TempDirectoryFixture>
{
    // XlChartType.xlBarClustered. Chosen over the xlColumnClustered default so a
    // dropped parameter cannot pass by coincidence.
    private const int XlBarClustered = 57;

    private readonly TempDirectoryFixture _fixture;
    private readonly SlideCommands _slides = new();
    private readonly ShapeCommands _shapes = new();
    private readonly ChartCommands _charts = new();

    public ChartRoundTripTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void Create_ProducesChartOfRequestedTypeAtRequestedBounds()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: false);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            var created = _charts.Create(batch, slideIndex: 1, chartType: XlBarClustered,
                left: 80f, top: 60f, width: 400f, height: 300f);
            Assert.True(created.Success, created.ErrorMessage);

            var list = _shapes.List(batch, slideIndex: 1);
            Assert.True(list.Success, list.ErrorMessage);

            var chartShape = Assert.Single(list.Shapes, s => s.HasChart);
            Assert.Equal(80f, chartShape.Left, 1);
            Assert.Equal(60f, chartShape.Top, 1);
            Assert.Equal(400f, chartShape.Width, 1);
            Assert.Equal(300f, chartShape.Height, 1);

            var info = _charts.GetInfo(batch, slideIndex: 1, shapeName: chartShape.Name);

            Assert.True(info.Success, info.ErrorMessage);
            Assert.Equal(XlBarClustered, info.ChartType);
            Assert.False(string.IsNullOrEmpty(info.ChartTypeName));

            // SeriesCount used to be read inside a catch-all that left it at 0 on
            // failure, which is indistinguishable from a chart that genuinely has no
            // series - and GetInfo still reported Success. A new chart is created from
            // PowerPoint's default worksheet, so it always has at least one series.
            Assert.True(info.SeriesCount > 0,
                $"Expected the default chart to report at least one series, got {info.SeriesCount}.");
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }
}
