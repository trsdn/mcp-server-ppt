// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using System.Text.RegularExpressions;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Design;
using PptMcp.Core.Tests.Helpers;
using Xunit;

namespace PptMcp.Core.Tests.Integration;

/// <summary>
/// Integration coverage for theme and colour-scheme reads (GitHub #126).
///
/// <c>DesignCommands.GetColors</c> and <c>ListColorSchemes</c> both build a dictionary of
/// named colours one entry at a time, with each read wrapped in a catch-all. A failed read
/// simply omitted its key, so the caller received a *shorter map* rather than an error -
/// and the result still carried <c>Success = true</c>. Nothing distinguished "this theme
/// has nine colours" from "three reads failed", which matters because callers index these
/// maps by name: a missing "Accent3" reads as a theme that has no Accent3.
///
/// Both catches also concealed a COM leak, because the release sat inside the guarded
/// block and was skipped whenever the read threw.
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
    /// Documents a surprising truth rather than an aspiration: <c>Presentation.ColorSchemes</c>
    /// is the pre-2007 API that themes replaced, and it is **empty** for a modern .pptx -
    /// verified here both with and without slides. So <c>design list-color-schemes</c> can
    /// never return data (GitHub #174), and its per-colour catch-all at
    /// <c>DesignCommands.ListColorSchemes</c> is unreachable, which is why that one catch is
    /// left in the #126 baseline: there is no way to prove removing it is safe.
    ///
    /// This is a characterisation test. If it ever fails, the operation started returning
    /// something - which is what the follow-up issue asks for - and this test should be
    /// replaced by real assertions on the roles, not deleted.
    /// </summary>
    [Fact]
    public void ListColorSchemes_OnAModernPresentation_ReturnsNothing()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;
            new PptMcp.Core.Commands.Slide.SlideCommands()
                .Create(batch, position: 0, layoutName: "Title Only");

            var schemes = _designs.ListColorSchemes(batch);

            // Success with nothing in it. That combination is the problem: a caller cannot
            // tell "this deck has no colour schemes" from "this API no longer reports them".
            Assert.True(schemes.Success, schemes.ErrorMessage);
            Assert.Empty(schemes.ColorSchemes);

            // The theme colours are the live equivalent, and they are populated - which is
            // what makes the empty result above a reporting gap rather than an empty deck.
            var colors = _designs.GetColors(batch, designIndex: 1);
            Assert.True(colors.Success, colors.ErrorMessage);
            Assert.NotEmpty(colors.Colors);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }
}
