---
applyTo: "src/PptMcp.Core/**/*.cs"
---

# PowerPoint COM Interop Patterns

> **Essential patterns for PowerPoint COM automation**

## Core Principles

1. **Use Late Binding** - `dynamic` types with `Type.GetTypeFromProgID()`
2. **1-Based Indexing** - PowerPoint collections start at 1, not 0
3. **Exception Propagation** - Never wrap in try-catch, let batch.Execute() handle exceptions (see Exception Propagation section)

## Reference Resources

**NetOffice Library** - THE BEST source for ALL PowerPoint COM Interop patterns:
- GitHub: https://github.com/NetOfficeFw/NetOffice
- **Use for ALL COM Interop work** - slides, shapes, presentations, charts, tables, VBA, everything
- NetOffice wraps Office COM APIs in strongly-typed C# - study their patterns for dynamic interop conversion
- Search NetOffice repository BEFORE implementing any PowerPoint COM automation
- Particularly valuable for: Shapes, Slide layouts, Masters, Animations, complex COM scenarios

## Exception Propagation Pattern (CRITICAL)

**Core Commands: NEVER wrap operations in try-catch blocks that return error results. Let exceptions propagate naturally.**

```csharp
// ❌ WRONG: Catching and wrapping exceptions
public async Task<OperationResult> CreateAsync(IPptBatch batch, string name)
{
    try
    {
        return await batch.Execute((ctx, ct) => {
            var item = ctx.Create(name);
            return ValueTask.FromResult(new OperationResult { Success = true });
        });
    }
    catch (Exception ex)
    {
        // ❌ WRONG: Double-wrapping suppresses the exception
        return new OperationResult { Success = false, ErrorMessage = ex.Message };
    }
}

// ✅ CORRECT: Let batch.Execute() handle exceptions via TaskCompletionSource
public async Task<OperationResult> CreateAsync(IPptBatch batch, string name)
{
    return await batch.Execute((ctx, ct) => {
        var item = ctx.Create(name);
        return ValueTask.FromResult(new OperationResult { Success = true });
    });
    // Exception flows to batch.Execute() → caught via TaskCompletionSource
    // → Returns OperationResult { Success = false, ErrorMessage }
}

// ✅ CORRECT: Finally blocks are the right place for COM resource cleanup
public async Task<OperationResult> ComplexAsync(IPptBatch batch, string name)
{
    dynamic? temp = null;
    try
    {
        return await batch.Execute((ctx, ct) => {
            temp = ctx.CreateTemp(name);
            // ... operation ...
            return ValueTask.FromResult(new OperationResult { Success = true });
        });
    }
    finally
    {
        // ✅ Finally for resource cleanup, NOT catch for error handling
        if (temp != null)
        {
            ComUtilities.Release(ref temp!);
        }
    }
}
```

**Why This Pattern:**
- `batch.Execute()` ALREADY captures exceptions via `TaskCompletionSource` 
- Inner try-catch suppresses exceptions, causing double-wrapping and lost stack context
- Finally blocks work perfectly for COM resource cleanup (which must happen regardless of exception)
- Exception occurs at correct layer (batch), not suppressed at method level

**Safe Exception Handling (Keep these):**
- ✅ Loop continuations: `catch { continue; }` (safe, recovers loop)
- ✅ Optional property access: `catch { value = null; }` (safe, uses fallback)
- ✅ Specific error routing: `catch (COMException ex) when (ex.HResult == code) { ... }` (specific, not general)
- ✅ Finally blocks: Resource cleanup for COM objects (always needed)

**Pattern to Remove:**
- ❌ `catch (Exception ex) { return new Result { Success = false, ErrorMessage = ex.Message }; }`

**Architecture:**
```
Core Command (NO try-catch wrapping)
  └─> await batch.Execute()
      └─> TaskCompletionSource captures exception
          └─> Returns OperationResult { Success = false, ErrorMessage }
```

---

## Resource Management

### ✅ Unified Shutdown Pattern (Current Standard)

**All presentation close and PowerPoint quit operations use `PptShutdownService` with resilient retry:**

```csharp
// In PptBatch, PptSession, FileCommands:
PptShutdownService.CloseAndQuit(presentation, powerpoint, save: false, filePath, logger);
```

**Shutdown Order:**
1. **Optional Save** - If `save=true`, calls `presentation.Save()` explicitly before close
2. **Close Presentation** - Calls `presentation.Close()` (save param controls PowerPoint's prompt behavior)
3. **Release Presentation** - Releases COM reference via `ComUtilities.Release()`
4. **Quit PowerPoint** - Calls `powerpoint.Quit()` with exponential backoff retry (6 attempts, 200ms base delay)
5. **Release PowerPoint** - Releases COM reference via `ComUtilities.Release()`
6. **Automatic GC** - RCW finalizers handle final cleanup automatically (no forced GC needed per Microsoft guidance)

**Resilience Features:**
- Uses `Microsoft.Extensions.Resilience` retry pipeline
- **Outer timeout (30s)**: Overall cancellation for PowerPoint.Quit() - catches hung PowerPoint (modal dialogs, deadlocks)
- **Inner retry**: Exponential backoff (200ms base, 2x factor, 6 attempts) for transient COM busy errors
- Retries on: `RPC_E_SERVERCALL_RETRYLATER` (-2147417851), `RPC_E_CALL_REJECTED` (-2147418111)
- Structured logging for diagnostics (attempt number, HResult, elapsed time)
- Continues with COM cleanup even if Quit fails/times out
- **STA thread join (45s)**: Must be >= PowerPointQuitTimeout + margin (currently 30s + 15s) to ensure Dispose() waits for full cleanup

**Save Semantics:**
```csharp
// Discard changes (default for disposal paths)
PptShutdownService.CloseAndQuit(presentation, powerpoint, save: false, filePath, logger);

// Save before close (for explicit save operations)
PptShutdownService.CloseAndQuit(presentation, powerpoint, save: true, filePath, logger);
```

**Why Unified Service:**
- Eliminates duplicated try/catch blocks across `PptBatch`, `PptSession`, `FileCommands`
- Consistent retry behavior for all PowerPoint quit operations
- Centralized logging and diagnostics
- Handles edge cases: disconnected COM proxies, hung PowerPoint, modal dialogs

**Timeout Architecture (Proper Layering):**
```
Overall Quit Timeout: 30 seconds (outer)
  └─> Resilient Retry: 6 attempts with exponential backoff (inner, ~6s max)
      └─> Individual Quit() calls
  └─> STA Thread Join: 45 seconds (PowerPointQuitTimeout + 15s margin)
```
- **30s quit timeout**: Catches truly hung PowerPoint (modal dialogs, deadlocks) via CancellationToken
- **6-attempt retry**: Handles transient COM busy states within the 30s window
- **45s thread join**: Must be >= PowerPointQuitTimeout + margin to ensure Dispose() waits for full cleanup

## COM Object Cleanup Pattern (CRITICAL)

**ALWAYS use try-finally for COM object cleanup. NEVER use catch blocks to swallow exceptions.**

### ❌ WRONG Patterns

```csharp
// WRONG #1: COM cleanup in try block (won't execute if exception occurs)
try
{
    dynamic pivotLayout = chart.PivotLayout;
    dynamic pivotTable = pivotLayout.PivotTable;
    name = pivotTable.Name?.ToString() ?? string.Empty;
    ComUtilities.Release(ref pivotTable!);  // ❌ Won't execute if exception above!
    ComUtilities.Release(ref pivotLayout!);
}
catch
{
    name = "(unknown)";  // ❌ Swallows exception, causes COM leak
}

// WRONG #2: Empty catch block (swallows exceptions silently)
try
{
    dynamic item = GetItem();
    // ... operations ...
    ComUtilities.Release(ref item!);
}
catch
{
    // ❌ Empty catch - swallows exception, no cleanup
}
```

### ✅ CORRECT Pattern

```csharp
// CORRECT: Finally block ensures cleanup regardless of exceptions
dynamic? pivotLayout = null;
dynamic? pivotTable = null;
try
{
    pivotLayout = chart.PivotLayout;
    pivotTable = pivotLayout.PivotTable;
    name = pivotTable.Name?.ToString() ?? string.Empty;
}
finally
{
    // ✅ ALWAYS executes - exception or no exception
    if (pivotTable != null) ComUtilities.Release(ref pivotTable!);
    if (pivotLayout != null) ComUtilities.Release(ref pivotLayout!);
}
// ✅ Exception propagates naturally to batch.Execute()
```

**Pattern Requirements:**
1. **Declare COM objects as `dynamic?` nullable** before try block
2. **Initialize to `null`**
3. **Acquire COM objects in try block**
4. **Release in finally block** with null checks
5. **NO catch blocks** unless specific exception handling required
6. **NEVER catch to set fallback values** - let exceptions propagate

**Why This Matters:**
- Finally blocks execute **regardless** of exceptions (try succeeds or fails)
- COM objects leak if Release() not reached before exception
- Swallowing exceptions with catch blocks hides real problems from batch.Execute()
- Empty catch blocks are code smell - remove them entirely
- Let exceptions propagate naturally to batch.Execute() for proper error handling

**See Also:**
- CRITICAL-RULES.md Rule 22 for complete requirements
- CRITICAL-RULES.md Rule 1b for exception propagation pattern

## Critical COM Issues

### 1. PowerPoint Collections Are 1-Based
```csharp
// ❌ WRONG: collection.Item(0)  
// ✅ CORRECT: collection.Item(1)
for (int i = 1; i <= collection.Count; i++) { var item = collection.Item(i); }
```

### 2. Numeric Property Type Conversions

**⚠️ ALL PowerPoint COM numeric properties return `double`, NOT `int`!**

```csharp
// ❌ WRONG: Implicit conversion fails at runtime
int slideIndex = slide.SlideIndex;    // Runtime error: Cannot convert double to int
int shapeCount = slide.Shapes.Count;  // Runtime error: Cannot convert double to int

// ✅ CORRECT: Explicit conversion required
int slideIndex = Convert.ToInt32(slide.SlideIndex);
int shapeCount = Convert.ToInt32(slide.Shapes.Count);
```

**Common Properties Affected:**
- `Slide.SlideIndex` → `double` (not `int`)
- `Shape.Left`, `Shape.Top`, `Shape.Width`, `Shape.Height` → `double` (not `float`)
- Any numeric property from PowerPoint COM → assume `double`

**Why:** PowerPoint COM uses `VARIANT` types internally, which represent numbers as `double`. C# `dynamic` binding preserves this type.

### 3. PowerPoint Busy Handling
```csharp
catch (COMException ex) when (ex.HResult == -2147417851)
{
    // RPC_E_SERVERCALL_RETRYLATER - PowerPoint is busy
}
```

## Common Patterns

### Read Slide Content
```csharp
dynamic slide = presentation.Slides[1];
dynamic shapes = slide.Shapes;
for (int i = 1; i <= shapes.Count; i++) { var shape = shapes.Item(i); }
```

### Add Shape
```csharp
dynamic shape = slide.Shapes.AddShape(msoShapeType, left, top, width, height);
shape.TextFrame.TextRange.Text = "Hello";
```

---

## Common Mistakes

| Mistake | Fix |
|---------|-----|
| 0-based indexing | PowerPoint is 1-based |
| Not releasing objects | `try/finally` + `ReleaseComObject()` |
| `int x = shape.Property` | Use `Convert.ToInt32()` for ALL numeric properties |
| Assuming enum types | Numeric properties return `double`, convert to enum |

**📚 Reference:** [PowerPoint Object Model](https://docs.microsoft.com/en-us/office/vba/api/overview/powerpoint)
