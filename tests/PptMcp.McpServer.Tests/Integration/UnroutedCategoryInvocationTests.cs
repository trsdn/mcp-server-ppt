// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json;
using PptMcp.Service;
using Xunit;

namespace PptMcp.McpServer.Tests.Integration;

/// <summary>
/// End-to-end coverage for categories that were advertised through MCP discovery but
/// unroutable in the service (GitHub #124).
///
/// <see cref="ServiceRoutingTests"/> proves every category is wired without needing
/// PowerPoint. This exercises the reported user path for real: create a session, then
/// invoke a previously-unroutable category and get a genuine result back.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Service")]
[Trait("Feature", "ServiceRouting")]
[Trait("RequiresPowerPoint", "true")]
public sealed class UnroutedCategoryInvocationTests : IDisposable
{
    private readonly string _tempDir;

    public UnroutedCategoryInvocationTests()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"PptMcp_Routing_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    [Fact]
    public async Task CommentList_ReturnsResult_InsteadOfUnknownCommandCategory()
    {
        var filePath = Path.Combine(_tempDir, $"comment_{Guid.NewGuid():N}.pptx");

        using var service = new PptMcpService();

        var created = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.create",
            Args = JsonSerializer.Serialize(new { filePath, show = false }),
        });

        Assert.True(created.Success, created.ErrorMessage);
        var sessionId = JsonDocument.Parse(created.Result!).RootElement.GetProperty("sessionId").GetString();

        try
        {
            // A new presentation starts with no slides.
            var slideCreated = await service.ProcessAsync(new ServiceRequest
            {
                Command = "slide.create",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new { position = 0, layoutName = "Blank" }),
            });

            Assert.True(slideCreated.Success, slideCreated.ErrorMessage);

            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "comment.list",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new { slideIndex = 1 }),
            });

            Assert.True(response.Success, response.ErrorMessage);
            Assert.NotNull(response.Result);

            // The reported symptom before the fix.
            Assert.DoesNotContain("Unknown command category", response.Result, StringComparison.Ordinal);
        }
        finally
        {
            await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new { save = false, force = true }),
            });
        }
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
                Directory.Delete(_tempDir, recursive: true);
        }
        catch
        {
            // Cleanup failures are non-critical.
        }
    }
}
