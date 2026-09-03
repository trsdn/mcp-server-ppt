// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.DocumentProperty;
using PptMcp.Core.Tests.Helpers;
using Xunit;

namespace PptMcp.Core.Tests.Integration;

/// <summary>
/// Integration coverage for built-in document properties (GitHub #126).
///
/// Both halves of this round trip used to swallow their failures:
///
/// - <c>SetProp</c> caught everything on the grounds that "some props may be read-only",
///   so a write that never happened was still reported by <c>SetAll</c> as
///   <c>Success = true, "Updated document properties"</c>. A caller had no way to learn
///   that the presentation was unchanged. The seven properties <c>SetAll</c> writes are
///   all writable built-ins, so the tolerance was not buying anything.
/// - <c>GetProp</c> returned <c>""</c> on failure, which is indistinguishable from a
///   property that is genuinely empty.
///
/// Together those two produced the worst possible pairing: a write that silently did
/// nothing, followed by a read that silently reported nothing, with both operations
/// claiming success. A round trip is the only thing that catches it, and there was no
/// round-trip test - this feature had no integration coverage at all.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "DocumentProperty")]
[Trait("RequiresPowerPoint", "true")]
public sealed class DocumentPropertyRoundTripTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TempDirectoryFixture _fixture;
    private readonly DocumentPropertyCommands _properties = new();

    public DocumentPropertyRoundTripTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void SetAll_ThenGetAll_ReturnsEveryPropertyThatWasWritten()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;

            // Distinct values per field: identical values would let a set that wrote to
            // the wrong index still satisfy the assertions.
            var set = _properties.SetAll(batch,
                title: "Round trip title",
                subject: "Round trip subject",
                author: "Round trip author",
                keywords: "alpha, beta, gamma",
                comments: "Round trip comments",
                company: "Round trip company",
                category: "Round trip category");

            Assert.True(set.Success, set.ErrorMessage);

            var read = _properties.GetAll(batch);
            Assert.True(read.Success, read.ErrorMessage);

            Assert.Equal("Round trip title", read.Title);
            Assert.Equal("Round trip subject", read.Subject);
            Assert.Equal("Round trip author", read.Author);
            Assert.Equal("alpha, beta, gamma", read.Keywords);
            Assert.Equal("Round trip comments", read.Comments);
            Assert.Equal("Round trip company", read.Company);
            Assert.Equal("Round trip category", read.Category);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void SetAll_Twice_OverwritesRatherThanKeepingTheFirstValue()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;

            _properties.SetAll(batch, "First", "First", "First", "First", "First", "First", "First");
            var second = _properties.SetAll(batch,
                "Second", "Second", "Second", "Second", "Second", "Second", "Second");
            Assert.True(second.Success, second.ErrorMessage);

            var read = _properties.GetAll(batch);
            Assert.True(read.Success, read.ErrorMessage);

            // A silently-failing second write would leave "First" in place while still
            // reporting success - exactly what the swallowed catch allowed.
            Assert.Equal("Second", read.Title);
            Assert.Equal("Second", read.Author);
            Assert.Equal("Second", read.Company);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    [Fact]
    public void SetCustom_ThenGetCustom_RoundTrips()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;

            var set = _properties.SetCustom(batch, "ReviewStage", "Draft");
            Assert.True(set.Success, set.ErrorMessage);

            var read = _properties.GetCustom(batch, "ReviewStage");
            Assert.True(read.Success, read.ErrorMessage);
            Assert.Contains("Draft", read.Message);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }



    [Fact]
    public void SetCustom_OnAnExistingProperty_UpdatesItInPlace()
    {
        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;

            _properties.SetCustom(batch, "ReviewStage", "Draft");
            var update = _properties.SetCustom(batch, "ReviewStage", "Final");
            Assert.True(update.Success, update.ErrorMessage);

            // Exercises the update branch rather than the add branch, which is the one
            // guarded by the existence probe in SetCustom.
            var read = _properties.GetCustom(batch, "ReviewStage");
            Assert.True(read.Success, read.ErrorMessage);
            Assert.Contains("Final", read.Message);
            Assert.DoesNotContain("Draft", read.Message);
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }

    /// <summary>
    /// Pins the index-to-name mapping <c>DocumentPropertyCommands</c> relies on.
    /// The constants are raw <c>BuiltInDocumentProperties</c> indices, and a wrong index
    /// is invisible to a round trip: writing and reading the *same* wrong index still
    /// agrees with itself. Only the property's own <c>Name</c> reveals the mistake.
    /// All seven are checked in one session so a mismatch reports every offender at once.
    /// </summary>
    [Fact]
    public void BuiltInPropertyIndices_MapToTheExpectedNames()
    {
        var expected = new (int Index, string Name)[]
        {
            (1, "Title"),
            (2, "Subject"),
            (3, "Author"),
            (4, "Keywords"),
            (5, "Comments"),
            (18, "Category"),
            (21, "Company")
        };

        var testFile = _fixture.CreateTestFile();

        using var manager = new SessionManager();
        var sessionId = manager.CreateSession(testFile, show: true);

        try
        {
            var batch = manager.GetSession(sessionId)!;

            var actual = batch.Execute((ctx, ct) =>
            {
                var names = new List<string>();
                dynamic pres = ctx.Presentation;
                dynamic? builtIn = null;
                try
                {
                    builtIn = pres.BuiltInDocumentProperties;
                    foreach (var (index, _) in expected)
                    {
                        dynamic? prop = null;
                        try
                        {
                            prop = builtIn.Item(index);
                            names.Add((string)(prop.Name?.ToString() ?? ""));
                        }
                        finally
                        {
                            if (prop != null) ComUtilities.Release(ref prop!);
                        }
                    }

                    return names;
                }
                finally
                {
                    if (builtIn != null) ComUtilities.Release(ref builtIn!);
                }
            });

            Assert.Equal(expected.Select(e => e.Name).ToArray(), actual.ToArray());
        }
        finally
        {
            manager.CloseSession(sessionId, save: false);
        }
    }
}
