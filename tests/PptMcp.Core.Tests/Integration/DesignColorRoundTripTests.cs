// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using System.Text.RegularExpressions;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Design;
using PptMcp.Core.Tests.Helpers;
using Xunit;

namespace PptMcp.Core.Tests.Integration;

/// <summary>
/// Integration coverage for theme and colour-scheme reads (GitHub #126, #174).
///
/// <c>DesignCommands.GetColors</c> built a dictionary of named colours one entry at a time,
/// with each read wrapped in a catch-all. A failed read simply omitted its key, so the caller
/// received a *shorter map* rather than an error - and the result still carried
/// <c>Success = true</c>. Nothing distinguished "this theme has nine colours" from "three reads
/// failed", which matters because callers index these maps by name: a missing "Accent3" reads
/// as a theme that has no Accent3.
///
/// The catch also concealed a COM leak, because the release sat inside the guarded block and
/// was skipped whenever the read threw.
///
/// These tests pin the complete key set, so a truncated read now fails loudly.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "Design")]
[Trait("RequiresPowerPoint", "true")]
public sealed class DesignColorRoundTripTests : IClassFixture<TempDirectoryFixture>
{
    /// <summary>The twelve MsoThemeColorSchemeIndex roles GetColors enumerates.</summary>
    private static readonly string[] ThemeColorRoles =
    [
        "Dark1", "Light1", "Dark2", "Light2",
        "Accent1", "Accent2", "Accent3", "Accent4",
        "Accent5", "Accent6", "Hyperlink", "FollowedHyperlink"
    ];

    private static readonly Regex HexColor = new("^#[0-9A-F]{6}$", RegexOptions.Compiled);

    private readonly TempDirectoryFixture _fixture;
    private readonly DesignCommands _designs = new();

    public DesignColorRoundTripTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void GetColors_ReturnsEveryThemeColorRole()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;

            var colors = _designs.GetColors(batch, designIndex: 1);
            Assert.True(colors.Success, colors.ErrorMessage);

            // Asserting the exact key set, not merely a count: a truncated walk that
            // omitted Accent3 would still satisfy "at least eight colours".
            Assert.Equal(
                ThemeColorRoles.OrderBy(n => n).ToArray(),
                colors.Colors.Keys.OrderBy(n => n).ToArray());

            // Every value must be a real colour. An entry present but malformed would
            // otherwise pass a key-only assertion.
            Assert.All(colors.Colors, kvp =>
                Assert.True(HexColor.IsMatch(kvp.Value), $"{kvp.Key} = '{kvp.Value}'"));

            Assert.False(string.IsNullOrWhiteSpace(colors.DesignName));
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    /// <summary>
    /// Every design carries its own theme colour scheme, and <c>list-color-schemes</c> reports
    /// one entry per design (GitHub #174).
    ///
    /// <para>This replaces a characterisation test that pinned the old behaviour: the action read
    /// <c>Presentation.ColorSchemes</c>, the pre-2007 API that themes replaced, which is empty for
    /// every modern .pptx. It returned <c>Success = true</c> with an empty list, and success plus
    /// nothing is indistinguishable from a genuine answer - an LLM reads "this deck has no colour
    /// schemes" while <c>get-colors</c> returns a full twelve-role palette for the same file.</para>
    ///
    /// <para>The cross-check against <c>GetColors</c> is the point of the test, not decoration.
    /// Asserting only that the list is non-empty would be satisfied by entries reporting the wrong
    /// design's palette, since every value is a well-formed colour either way.</para>
    /// </summary>
    [Fact]
    public void ListColorSchemes_ReportsTheThemeSchemeOfEveryDesign()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            new PptMcp.Core.Commands.Slide.SlideCommands()
                .Create(batch, position: 0, layoutName: "Title Only");

            var designs = _designs.List(batch);
            Assert.True(designs.Success, designs.ErrorMessage);
            Assert.NotEmpty(designs.Designs);

            var schemes = _designs.ListColorSchemes(batch);
            Assert.True(schemes.Success, schemes.ErrorMessage);

            Assert.Equal(designs.Designs.Count, schemes.ColorSchemes.Count);

            foreach (var scheme in schemes.ColorSchemes)
            {
                // The exact key set, not a count: a truncated walk that omitted Accent3 would
                // still satisfy "at least eight colours".
                Assert.Equal(
                    ThemeColorRoles.OrderBy(n => n).ToArray(),
                    scheme.Colors.Keys.OrderBy(n => n).ToArray());

                Assert.All(scheme.Colors, kvp =>
                    Assert.True(HexColor.IsMatch(kvp.Value), $"[{scheme.Index}] {kvp.Key} = '{kvp.Value}'"));

                Assert.False(string.IsNullOrWhiteSpace(scheme.DesignName));

                // Independent check that Index addresses the design it claims to. Without this,
                // every entry could report design 1's palette and the test would still pass.
                var direct = _designs.GetColors(batch, scheme.Index);
                Assert.True(direct.Success, direct.ErrorMessage);
                Assert.Equal(direct.DesignName, scheme.DesignName);
                Assert.Equal(direct.Colors, scheme.Colors);
            }
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }
}
