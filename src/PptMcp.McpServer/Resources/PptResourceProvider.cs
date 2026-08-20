using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace PptMcp.McpServer.Resources;

/// <summary>
/// MCP resources for documenting available PowerPoint presentation URIs.
/// Resources help LLMs understand what can be inspected in PowerPoint presentations.
/// 
/// NOTE: MCP SDK 0.4.0-preview.2 does NOT support McpServerResourceTemplate yet.
/// Dynamic URI patterns (ppt://{path}/slides/{name}) will be added when SDK supports it.
/// For now, use tools (slide list, etc.) for actual data retrieval.
/// </summary>
[McpServerResourceType]
public static class PptResourceProvider
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    /// <summary>
    /// Documents available PowerPoint presentation resource URIs.
    /// </summary>
    [McpServerResource(UriTemplate = "ppt://help/resources")]
    [Description("Guide to available PowerPoint presentation resources")]
    public static Task<string> GetResourceGuide()
    {
        var guide = new
        {
            title = "PowerPoint Presentation Resources",
            description = "URI patterns for inspecting PowerPoint presentations",
            note = "Use tools to retrieve actual data (MCP SDK resource templates not yet supported)",
            resourceTypes = new[]
            {
                new
                {
                    type = "Slides",
                    toolAction = "Use slide tool with action='list' to see all slides",
                    example = "slide(action: 'list', session_id: '1')"
                },
                new
                {
                    type = "Shapes",
                    toolAction = "Use shape tool with action='list' to see all shapes on a slide",
                    example = "shape(action: 'list', session_id: '1', slide_index: 1)"
                },
                new
                {
                    type = "Text",
                    toolAction = "Use text tool with action='get' to read shape text",
                    example = "text(action: 'get', session_id: '1', slide_index: 1, shape_name: 'Title 1')"
                },
                new
                {
                    type = "Speaker Notes",
                    toolAction = "Use notes tool with action='get'",
                    example = "notes(action: 'get', session_id: '1', slide_index: 1)"
                },
                new
                {
                    type = "Slide Tables",
                    toolAction = "Use slidetable tool with action='list'",
                    example = "slidetable(action: 'list', session_id: '1', slide_index: 1)"
                },
                new
                {
                    type = "Comments",
                    toolAction = "Use comment tool with action='list'",
                    example = "comment(action: 'list', session_id: '1', slide_index: 0)"
                },
                new
                {
                    type = "Sections",
                    toolAction = "Use section tool with action='list'",
                    example = "section(action: 'list', session_id: '1')"
                },
                new
                {
                    type = "Themes and Design",
                    toolAction = "Use design tool with action='list'",
                    example = "design(action: 'list', session_id: '1')"
                },
                new
                {
                    type = "VBA Modules",
                    toolAction = "Use vba tool with action='list' (.pptm files only)",
                    example = "vba(action: 'list', session_id: '1')"
                }
            },
            usage = new
            {
                discovery = "Use tool 'list' actions to discover presentation contents",
                inspection = "Use tool 'view' actions to examine specific items",
                modification = "Use other tool actions to create/update/delete items"
            },
            futureEnhancements = "Dynamic resource templates (ppt://{path}/slides/{name}) will be added when MCP SDK supports McpServerResourceTemplate"
        };

        return Task.FromResult(JsonSerializer.Serialize(guide, JsonOptions));
    }

    /// <summary>
    /// Quick reference for common PowerPoint operations.
    /// </summary>
    [McpServerResource(UriTemplate = "ppt://help/quickref")]
    [Description("Quick reference for common PowerPoint MCP operations")]
    public static Task<string> GetQuickReference()
    {
        var quickRef = new
        {
            title = "PowerPoint MCP Quick Reference",
            commonOperations = new[]
            {
                new
                {
                    task = "Open a presentation and start a session",
                    tool = "file",
                    action = "open",
                    example = "file(action: 'open', path: 'C:\\\\Decks\\\\presentation.pptx')"
                },
                new
                {
                    task = "List all slides",
                    tool = "slide",
                    action = "list",
                    example = "slide(action: 'list', session_id: '1')"
                },
                new
                {
                    task = "Add a slide",
                    tool = "slide",
                    action = "create",
                    example = "slide(action: 'create', session_id: '1', slide_index: 1)"
                },
                new
                {
                    task = "List shapes on a slide",
                    tool = "shape",
                    action = "list",
                    example = "shape(action: 'list', session_id: '1', slide_index: 1)"
                },
                new
                {
                    task = "Set shape text",
                    tool = "text",
                    action = "set",
                    example = "text(action: 'set', session_id: '1', slide_index: 1, shape_name: 'Title 1', text: 'Q4 Results')"
                },
                new
                {
                    task = "Render a slide as an image",
                    tool = "slide",
                    action = "get-thumbnail",
                    example = "slide(action: 'get-thumbnail', session_id: '1', slide_index: 1)"
                },
                new
                {
                    task = "Export the deck",
                    tool = "export",
                    action = "pdf",
                    example = "export(action: 'pdf', session_id: '1', destination_path: 'C:\\\\Decks\\\\out.pdf')"
                },
                new
                {
                    task = "Work with sessions",
                    tool = "file",
                    action = "open/close",
                    example = "file(action: 'open', path: '...') → operations with session_id → file(action: 'close', session_id: '1', save: true)"
                }
            },
            sessionWorkflow = new[]
            {
                "Open session: file(action: 'open', path: 'C:\\\\Decks\\\\deck.pptx')",
                "Use session_id with all subsequent operations",
                "Close session: file(action: 'close', session_id: '1', save: true)"
            }
        };

        return Task.FromResult(JsonSerializer.Serialize(quickRef, JsonOptions));
    }
}


