# PowerPoint MCP Server - Quick Reference

> **When user asks about PowerPoint files, presentations, or slides in .pptx/.pptm files - USE the PowerPoint MCP tools.**

## When to Use PowerPoint MCP

USE these tools when user wants to:
- Read/write PowerPoint data, shapes, or formatting
- Create slides, charts, or tables
- Import data via Power Query
- Run VBA macros
- Any .pptx or .pptm file operations

DO NOT use for: CSV files (use standard file tools), Google Slides, or non-PowerPoint formats.

---

## Prerequisites

- **Windows OS** - PowerPoint COM automation requires Windows
- **Microsoft PowerPoint 2016+** - Must be installed
- **File CLOSED in PowerPoint** - COM requires exclusive access

---

## Tool Selection (Which Tool for Which Task?)

| Task | Use | NOT |
|------|-----|-----|
| Import external data (CSV, SQL, APIs) | `powerquery` | `table` |
| DAX measures / calculated fields | `datamodel` | `range` |
| Slide formulas (SUM, VLOOKUP) | `range` | `datamodel` |
| Structured data with filtering | `table` | `range` |
| Interactive summarization | `pivottable` | `table` |
| VBA automation | `vba` | Requires .pptm |

**Data Model prerequisite**: Before using `datamodel`, data must be loaded with `loadDestination: 'data-model'` or `'both'` via `powerquery`.

---

## Cross-Tool Workflow Patterns

### Import Data -> Analyze -> Visualize
```
1. file(action: 'open')
2. powerquery(action: 'create', loadDestination: 'slide')
3. pivottable(action: 'create-from-table')
4. chart(action: 'create-from-pivottable')
5. file(action: 'close', save: true)
```

### Build DAX Analytics
```
1. file(action: 'open')
2. powerquery(action: 'create', loadDestination: 'data-model')
3. datamodel(action: 'create-measure', formula: 'SUM(...)')
4. pivottable(action: 'create-from-datamodel')
5. file(action: 'close', save: true)
```

### Batch Updates (Multiple Items)
```
# Use bulk data operations - set entire ranges at once:
range(action: 'set-values', values: [[1,2,3], [4,5,6]])  # NOT cell-by-cell
range(action: 'set-formulas', formulas: [['=A1', '=B1']])  # Multiple formulas at once
```

---

## Common Cross-Tool Mistakes

| Mistake | Fix |
|---------|-----|
| Using `table` to import CSV | Use `powerquery` (handles encoding, transforms) |
| Using `range` for DAX | Use `datamodel` (DAX != slide formulas) |
| Multiple single-item calls | Use bulk actions when available |
| Closing session between operations | Keep session open until workflow complete |
| Working on file open in PowerPoint | Ask user to close file first |

---

## Session Lifecycle Reminder

```
file(action: 'open') -> [all operations with sessionId] -> file(action: 'close')
```

- **DEFAULT: `showPowerPoint: false`** - Use hidden mode for faster background automation
- Only use `showPowerPoint: true` if user explicitly requests to watch changes
- If `showPowerPoint: true` was used, **ask before closing** (user may want to inspect)
- Use `file(action: 'list')` to check session state if uncertain

---

## For Per-Tool Details

Each tool has detailed documentation in its schema description. For server-specific quirks, the MCP server exposes prompt resources you can request.
