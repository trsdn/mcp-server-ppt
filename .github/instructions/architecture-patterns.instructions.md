---
applyTo: "src/**/*.cs"
---

# Architecture Patterns

> **Core patterns for PptMcp development**

## .NET Class Design (MANDATORY)

**Official Docs:** [Framework Design Guidelines](https://learn.microsoft.com/en-us/dotnet/standard/design-guidelines/), [Partial Classes](https://learn.microsoft.com/en-us/dotnet/csharp/programming-guide/classes-and-structs/partial-classes-and-methods)

### Key Rules

1. **One Public Class Per File** - Standard .NET practice (System.Text.Json, ASP.NET Core, EF Core)
2. **File Name = Class Name** - `RangeCommands.cs` contains `RangeCommands`
3. **Partial Classes for Large Implementations** - Split 15+ method classes by feature domain
4. **Descriptive Names** - No over-optimization (`RangeCommands` ✅, `Commands` ❌)
5. **Folder = Organization, Not Identity** - `Commands/Range/RangeCommands.cs`

### Partial Class Pattern

**When:** Class has 15+ methods, multiple feature domains, team collaboration

**Structure:**
```
Commands/Range/
    IRangeCommands.cs           # Interface
    RangeCommands.cs            # Partial (constructor, DI)
    RangeCommands.Values.cs     # Partial (Get/Set values)
    RangeCommands.Formulas.cs   # Partial (formulas)
    RangeHelpers.cs             # Separate helper class
```

**Benefits:** Git-friendly, team-friendly, ~100-200 lines per file, mirrors .NET Framework patterns

---

## TWO EQUAL ENTRY POINTS (CRITICAL)

**PptMcp has TWO first-class entry points: MCP Server AND CLI.** Both must have:
- **Feature parity**: Every action in MCP must exist in CLI and vice versa
- **Parameter parity**: Same parameters, same defaults, same validation
- **Behavior parity**: Same Core command, same result format

When adding or changing ANY feature, ALWAYS update BOTH entry points. See Rule 24 (Post-Change Sync).

```
MCP Server (MCP tools, JSON-RPC) ──► In-process PptMcpService ──► Core Commands ──► PowerPoint COM
CLI (command-line args, console)  ──► CLI Daemon (named pipe) ─────► Core Commands ──► PowerPoint COM
```

---

## Command Pattern

### Structure
```
Commands/
├── ISlideCommands.cs    # Interface
├── SlideCommands.cs     # Implementation
```

### Routing (Program.cs)
```csharp
return args[0] switch
{
    "slide-list" => slide.List(args),
    "shape-read" => shape.Read(args),
    _ => ShowHelp()
};
```

---

## Resource Management Pattern

**See excel-com-interop.instructions.md** for complete WithPowerPoint() pattern and COM object lifecycle management.

---

## Exception Propagation Pattern (CRITICAL)

**Core Commands: Let exceptions propagate naturally** - Do NOT suppress with catch blocks that return error results.

```csharp
// ❌ WRONG: Suppressing exception with catch block
public async Task<OperationResult> SomeAsync(IPptBatch batch, string param)
{
    try
    {
        return await batch.Execute((ctx, ct) => {
            // ... operation ...
            return ValueTask.FromResult(new OperationResult { Success = true });
        });
    }
    catch (Exception ex)
    {
        // ❌ WRONG: Catches exception and returns error result
        return new OperationResult 
        { 
            Success = false, 
            ErrorMessage = ex.Message 
        };
    }
}

// ✅ CORRECT: Let exception propagate through batch.Execute()
public async Task<OperationResult> SomeAsync(IPptBatch batch, string param)
{
    return await batch.Execute((ctx, ct) => {
        // ... operation ...
        return ValueTask.FromResult(new OperationResult { Success = true });
    });
    // Exception automatically caught by batch.Execute() via TaskCompletionSource
    // Returns OperationResult { Success = false, ErrorMessage } from batch layer
}

// ✅ CORRECT: Finally blocks still allowed for COM resource cleanup
public async Task<OperationResult> ComplexAsync(IPptBatch batch, string param)
{
    dynamic? shapeRef = null;
    try
    {
        return await batch.Execute((ctx, ct) => {
            shapeRef = ctx.Presentation.Slides[1].Shapes.AddShape(...);
            // ... operation ...
            return ValueTask.FromResult(new OperationResult { Success = true });
        });
    }
    finally
    {
        if (shapeRef != null)
        {
            ComUtilities.Release(ref shapeRef!);  // ✅ Cleanup in finally
        }
    }
}
```

**Why This Pattern:**
- `batch.Execute()` already captures exceptions via `TaskCompletionSource`
- Exceptions in lambda automatically become `OperationResult { Success = false }`
- Double-wrapping (try-catch returning error result) loses stack context and originates from wrong layer
- Finally blocks are the correct place for resource cleanup, NOT catch blocks for error suppression

**See:** CRITICAL-RULES.md Rule 1 for Success flag requirements

---

## MCP Server Resource-Based Tools

**In-Process Architecture**: MCP Server hosts PptMcpService fully in-process with direct method calls (no pipe).
ServiceBridge holds the service reference and calls ProcessAsync() directly.

**19 Focused Tools:**
1. `file` - Session lifecycle (open, close, create, list)
2. `slide` - Slide operations
3. `slide_style` - Slide layout and background
4. `shape` - Shape operations (add, modify, delete)
5. `text` - Text and TextFrame operations
6. `table` - Table operations on slides
7. `image` - Image and picture operations
8. `chart` - Chart lifecycle
9. `chart_config` - Chart configuration
10. `animation` - Animation effects
11. `transition` - Slide transitions
12. `slide_master` - Slide master and layout management
13. `notes` - Speaker notes
14. `section` - Presentation sections
15. `media` - Audio and video operations
16. `vba` - VBA macros
17. `comment` - Slide comments
18. `export` - Export slides (images, PDF)
19. `hyperlink` - Hyperlink operations

### Action-Based Routing with ForwardToService
```csharp
[McpServerTool]
public static string PptSlide(string action, string sessionId, ...)
{
    return action.ToLowerInvariant() switch
    {
        "list" => ForwardList(sessionId),
        "get" => ForwardGet(sessionId, slideIndex),
        _ => throw new McpException($"Unknown action: {action}")
    };
}

private static string ForwardList(string sessionId)
{
    return PptToolsBase.ForwardToService("slide.list", sessionId);
}
```

---

## DRY Shared Utilities

**PptHelper Methods:** `FindSlide()`, `FindShape()`, `GetShapeTypeName()`, `GetSlideLayout()`

**Why:** Prevents 60+ lines of duplicate code per feature

---

## Security-First Patterns

```csharp
// Defaults
SavePassword = false  // Never export credentials by default
```

---

## Performance Patterns

**Minimize presentation opens** - Use single session for multiple operations
**Bulk operations** - Minimize COM round-trips by batching shape/slide operations

---

## Key Principles

1. **WithPowerPoint() for everything** - See excel-com-interop.instructions.md
2. **Release intermediate objects** - Prevents PowerPoint hanging
3. **Batch/Session for MCP** - Multiple operations in single session
4. **Resource-based tools** - 19 tools, not 33+ operations
5. **DRY utilities** - Share common patterns
6. **Security defaults** - Never expose credentials
7. **Bulk operations** - Minimize COM round-trips
