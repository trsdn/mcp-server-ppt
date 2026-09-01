// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using PptMcp.Generated;
using PptMcp.Service;
using Xunit;

namespace PptMcp.McpServer.Tests.Integration;

/// <summary>
/// Guards the invariant that every generated service category is reachable through
/// <see cref="PptMcpService.ProcessAsync"/>.
///
/// Tools and their action dispatch are generated from the Core interfaces, but the
/// top-level category switch in PptMcpService was hand-written. The two drifted:
/// 11 categories (background, comment, customshow, headerfooter, pagesetup,
/// placeholder, printoptions, shapealign, slideimport, smartart, tag) were advertised
/// through MCP discovery while every invocation fell through to
/// "Unknown command category" (GitHub #124).
///
/// These tests need no PowerPoint: dispatch parses the action before it touches a
/// session, so an unknown action proves routing exists without opening a presentation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Service")]
[Trait("Feature", "ServiceRouting")]
public sealed class ServiceRoutingTests
{
    /// <summary>
    /// An action string that is deliberately not a valid action in any category.
    /// Routing is proven by the service rejecting the *action* rather than the *category*.
    /// </summary>
    private const string UnroutableAction = "zzz-not-a-real-action";

    public static TheoryData<string> AllCategories()
    {
        var data = new TheoryData<string>();
        foreach (var (cliCommandName, _, _) in _CliCategoryMetadata.Categories)
        {
            data.Add(cliCommandName);
        }

        return data;
    }

    [Theory]
    [MemberData(nameof(AllCategories))]
    public async Task EveryGeneratedCategory_IsRoutable(string category)
    {
        using var service = new PptMcpService();

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = $"{category}.{UnroutableAction}",
        });

        Assert.False(response.Success);
        Assert.NotNull(response.ErrorMessage);

        Assert.DoesNotContain(
            "Unknown command category",
            response.ErrorMessage,
            StringComparison.Ordinal);

        // Routing reached the category's generated dispatch, which rejected the action.
        Assert.Contains(
            "Unknown action",
            response.ErrorMessage,
            StringComparison.Ordinal);
    }

    /// <summary>
    /// The generated category list is the single source of truth. If a category is added
    /// to Core, this count changes and the service must pick it up automatically.
    /// </summary>
    [Fact]
    public void GeneratedCategoryList_IsNotEmpty()
    {
        Assert.NotEmpty(_CliCategoryMetadata.Categories);
    }
}
