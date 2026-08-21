---
applyTo: "src/PptMcp.Core/Commands/**/*.cs,src/PptMcp.McpServer/**/*.cs"
---

# Exposing Core Commands

> Adding a Core method is enough. The CLI command, the MCP tool and the action
> enum are generated from the interface — do not hand-write them.

## How the tool surface is produced

Three attributes on a Core interface drive everything:

| Attribute | Where | Produces |
|---|---|---|
| `[ServiceCategory("section")]` | interface | The service registry entry and the CLI command name |
| `[McpTool("section", Title = ..., Description = ...)]` | interface | The MCP tool, its title and the description sent to LLMs |
| `[ServiceAction("add")]` | method | One action on both entry points, plus the enum member |

The generators emit:

- `ServiceRegistry.<Tool>.g.cs` in `PptMcp.Core` — the action enum and dispatch
- `McpTool.<Tool>.g.cs` in `PptMcp.McpServer` — the MCP tool

Because the CLI and the MCP server are generated from the *same* interface, they
cannot drift apart. There is no `ToolActions.cs` and no `ActionExtensions.cs` to
keep in sync; both were removed with the hand-maintained architecture.

## Adding an operation

```csharp
// 1. Declare it on the interface, with an action name and XML documentation.
//    The XML docs become the descriptions an LLM sees, so write them for a
//    reader who cannot see the implementation.
[ServiceAction("duplicate")]
OperationResult Duplicate(IPptBatch batch, int slideIndex);

// 2. Implement it in the Commands class. Let exceptions propagate - batch.Execute()
//    converts them into OperationResult { Success = false } (Rule 1b).
public OperationResult Duplicate(IPptBatch batch, int slideIndex) =>
    batch.Execute((ctx, ct) => { /* ... */ });
```

```powershell
# 3. Build, so the generators run.
dotnet build

# 4. Write an integration test (Rule 29: the test comes first and must fail
#    before the implementation exists; Rule 30: never a unit test).
dotnet test --filter "Feature=Slide&RunType!=OnDemand"
```

Then update the counts and the feature reference (Rule 24). `FEATURES.md` is
generated, not hand-edited.

## Parameter naming

The MCP parameter name is derived from the C# parameter name via
`StringHelper.ToSnakeCase()`, so **never use underscores in C# parameter names**.
Choose a camelCase name that produces the wanted snake_case result
(`sourceRangeAddress` → `source_range_address`), or use
`[FromString("desiredName")]` when it cannot.

Names are judged as they appear in a *flat* tool schema, without the surrounding
class name (Rule 28): `rotation` is self-describing, but a bare `name` or `index`
is not — use `shapeName` and `slideIndex`.

## What is actually verified before commit

`scripts/check-documented-counts.ps1` derives the tool surface from the generated
registry and fails when a documented count disagrees with it. It also fails when
it detects nothing, because a coverage check that reports success without
inspecting anything is worse than no check at all — the gate it replaced reported
"100% coverage maintained" while detecting zero methods.

```powershell
# Requires a build first: the generated registry is the source of truth.
.\scripts\check-documented-counts.ps1
```
