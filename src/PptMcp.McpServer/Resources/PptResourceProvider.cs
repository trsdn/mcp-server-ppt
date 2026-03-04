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
                    type = "Power Queries",
                    toolAction = "Use powerquery tool with action='list' to see all queries",
                    example = "powerquery(action: 'list', presentationPath: 'presentation.pptx')"
                },
                new
                {
                    type = "Slides",
                    toolAction = "Use slide tool with action='list' to see all slides",
                    example = "slide(action: 'list', presentationPath: 'presentation.pptx')"
                },
                new
                {
                    type = "Parameters (Named Ranges)",
                    toolAction = "Use namedrange tool with action='list' to see all parameters",
                    example = "namedrange(action: 'list', presentationPath: 'presentation.pptx')"
                },
                new
                {
                    type = "Data Model Tables",
                    toolAction = "Use datamodel tool with action='list-tables'",
                    example = "datamodel(action: 'list-tables', presentationPath: 'presentation.pptx')"
                },
                new
                {
                    type = "DAX Measures",
                    toolAction = "Use datamodel tool with action='list-measures'",
                    example = "datamodel(action: 'list-measures', presentationPath: 'presentation.pptx')"
                },
                new
                {
                    type = "VBA Modules",
                    toolAction = "Use vba tool with action='list'",
                    example = "vba(action: 'list', presentationPath: 'presentation.pptm')"
                },
                new
                {
                    type = "Slide Tables",
                    toolAction = "Use table tool with action='list'",
                    example = "table(action: 'list', presentationPath: 'presentation.pptx')"
                },
                new
                {
                    type = "Connections",
                    toolAction = "Use connection tool with action='list'",
                    example = "connection(action: 'list', presentationPath: 'presentation.pptx')"
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
                    task = "List all Power Queries",
                    tool = "powerquery",
                    action = "list",
                    example = "powerquery(action: 'list', presentationPath: 'presentation.pptx')"
                },
                new
                {
                    task = "View Power Query M code",
                    tool = "powerquery",
                    action = "view",
                    example = "powerquery(action: 'view', presentationPath: 'presentation.pptx', queryName: 'SalesData')"
                },
                new
                {
                    task = "Import query to Data Model",
                    tool = "powerquery",
                    action = "import",
                    example = "powerquery(action: 'import', presentationPath: 'presentation.pptx', queryName: 'Sales', sourcePath: 'sales.pq', loadDestination: 'data-model')"
                },
                new
                {
                    task = "List all slides",
                    tool = "slide",
                    action = "list",
                    example = "slide(action: 'list', presentationPath: 'presentation.pptx')"
                },
                new
                {
                    task = "List all DAX measures",
                    tool = "datamodel",
                    action = "list-measures",
                    example = "datamodel(action: 'list-measures', presentationPath: 'presentation.pptx')"
                },
                new
                {
                    task = "Get cell values",
                    tool = "range",
                    action = "get-values",
                    example = "range(action: 'get-values', presentationPath: 'presentation.pptx', sheetName: 'Data', rangeAddress: 'A1:D10')"
                },
                new
                {
                    task = "Work with sessions",
                    tool = "file",
                    action = "open/close",
                    example = "file(action: 'open') → operations with sessionId → file(action: 'close', save: true)"
                }
            },
            sessionWorkflow = new[]
            {
                "Open session: file(action: 'open', presentationPath: '...')",
                "Use sessionId with all subsequent operations",
                "Close session: file(action: 'close', sessionId: '...', save: true)"
            }
        };

        return Task.FromResult(JsonSerializer.Serialize(quickRef, JsonOptions));
    }
}


