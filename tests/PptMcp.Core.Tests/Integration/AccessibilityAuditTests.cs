// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Accessibility;
using PptMcp.Core.Commands.Placeholder;
using PptMcp.Core.Commands.Shape;
using PptMcp.Core.Commands.Slide;
using PptMcp.Core.Tests.Helpers;
using Xunit;

namespace PptMcp.Core.Tests.Integration;

/// <summary>
/// Integration coverage for the accessibility audit (GitHub #126).
///
/// <c>AccessibilityCommands.AuditSlide</c> wrapped both of its enumeration loops - the
/// placeholder walk and the shape walk - in catch-alls. A failure part-way through either
/// loop produced a *plausible* wrong answer rather than an obvious one:
///
/// - The placeholder loop sets <c>hasTitle</c>. Aborting before it reaches the title
///   placeholder leaves it false, so the audit reports <c>MissingTitle</c> for a slide that
///   has a perfectly good title - a fabricated finding, reported as fact.
/// - The shape loop is what raises <c>MissingAltText</c>. Aborting mid-walk silently drops
///   the remaining findings, so the audit under-reports. That is the worse direction: a user
///   acts on "no accessibility issues found" and ships an inaccessible deck.
///
/// In both cases the result still carried <c>Success = true</c>, so nothing distinguished a
/// truncated audit from a clean one. That is why the catches could not simply be deleted:
/// there was no test asserting what a correct audit even returns. This file establishes it.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Accessibility")]
[Trait("RequiresPowerPoint", "true")]
public sealed class AccessibilityAuditTests : IClassFixture<TempDirectoryFixture>
{
    // MsoAutoShapeType.msoShapeRectangle
    private const int MsoShapeRectangle = 1;

    private readonly TempDirectoryFixture _fixture;
    private readonly SlideCommands _slides = new();
    private readonly ShapeCommands _shapes = new();
    private readonly AccessibilityCommands _accessibility = new();
    private readonly PlaceholderCommands _placeholders = new();

    public AccessibilityAuditTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void Audit_ShapesWithoutAltText_ReportsEveryOne()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            // Three rather than one: a walk that aborted after the first shape would still
            // satisfy an "at least one issue" assertion, which is exactly the truncation
            // the old catch-all hid.
            var names = AddRectangles(batch, count: 3);

            var audit = _accessibility.Audit(batch);
            Assert.True(audit.Success, audit.ErrorMessage);

            var altTextIssues = audit.Issues
                .Where(i => i.IssueType == "MissingAltText")
                .Select(i => i.ShapeName)
                .OrderBy(n => n)
                .ToArray();

            Assert.Equal(names.OrderBy(n => n).ToArray(), altTextIssues);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void Audit_AfterSettingAltText_StopsReportingTheShape()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");
            var names = AddRectangles(batch, count: 2);

            var set = _shapes.SetAltText(batch, slideIndex: 1, shapeName: names[0],
                altText: "A rectangle used as a decorative divider.");
            Assert.True(set.Success, set.ErrorMessage);

            var audit = _accessibility.Audit(batch);
            Assert.True(audit.Success, audit.ErrorMessage);

            var flagged = audit.Issues
                .Where(i => i.IssueType == "MissingAltText")
                .Select(i => i.ShapeName)
                .ToArray();

            // The remaining shape must still be flagged. Asserting both directions keeps
            // this honest: a walk that silently returned nothing would pass a test that
            // only checked the fixed shape was absent.
            Assert.DoesNotContain(names[0], flagged);
            Assert.Contains(names[1], flagged);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void Audit_SlideWithTitleText_DoesNotReportMissingTitle()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;

            // "Title Only" carries a title placeholder, which is what hasTitle detects.
            _slides.Create(batch, position: 0, layoutName: "Title Only");

            var title = _placeholders.SetText(batch, slideIndex: 1, placeholderIndex: 1,
                text: "Quarterly results");
            Assert.True(title.Success, title.ErrorMessage);

            var audit = _accessibility.Audit(batch);
            Assert.True(audit.Success, audit.ErrorMessage);

            // Both of these are findings the truncated placeholder walk used to fabricate:
            // aborting before reaching the title placeholder leaves hasTitle false, and a
            // failed text read leaves hasText false.
            Assert.DoesNotContain(audit.Issues, i => i.IssueType == "MissingTitle");
            Assert.DoesNotContain(audit.Issues, i => i.IssueType == "EmptyTitlePlaceholder");
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void Audit_BlankSlide_ReportsMissingTitle()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");

            var audit = _accessibility.Audit(batch);
            Assert.True(audit.Success, audit.ErrorMessage);

            // The counterpart to the test above: MissingTitle must still be raised when it
            // is genuinely true, otherwise that assertion could be satisfied by an audit
            // that never reports anything at all.
            Assert.Contains(audit.Issues, i => i.IssueType == "MissingTitle" && i.SlideIndex == 1);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void Audit_CountsMatchTheIssuesReturned()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            _slides.Create(batch, position: 0, layoutName: "Blank");
            AddRectangles(batch, count: 2);

            var audit = _accessibility.Audit(batch);
            Assert.True(audit.Success, audit.ErrorMessage);

            Assert.Equal(audit.Issues.Count, audit.IssueCount);
            Assert.Equal(1, audit.TotalSlides);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    private string[] AddRectangles(IPptBatch batch, int count)
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
}
