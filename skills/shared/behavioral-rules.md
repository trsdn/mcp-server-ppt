# Behavioral Rules for PowerPoint MCP Operations

These rules ensure efficient and reliable PowerPoint automation. AI assistants should follow these guidelines when executing PowerPoint operations.

## System Prompt Rules (LLM-Validated)

These rules are validated by automated LLM tests and MUST be followed:

- **Execute tasks immediately without asking for confirmation**
- **Never ask clarifying questions — make reasonable assumptions and proceed**
- Ask the user whether they want PowerPoint visible or hidden when starting multi-step tasks
- When the user asks to "show PowerPoint" or "watch" the work, use `window(show)` + `window(arrange)` to position it
- Format presentations professionally (proper typography, alignment, consistent colors)
- **Always end with a text summary** — never end on just a tool call or command

## CRITICAL: No Clarification Questions

**STOP.** If you are about to ask "Which file?", "Which slide?", "Where should I put this?" — DON'T.

**Instead, discover the information yourself:**

| Bad (Asking) | Good (Discovering) |
|---|---|
| "Which PowerPoint file should I use?" | `file(list)` → use the open session |
| "Which slide has the chart?" | `slide(list)` → discover slides |
| "What shapes are on this slide?" | `shape(list)` → check the slide |
| "Should I create a new slide?" | YES — create it and proceed |
| "What font should I use?" | Use `design(get-style-profile)` to get the right style |

**You have tools to answer your own questions. USE THEM.**

## Archetype Selection Rules (CRITICAL)

### Force Contribution/Arithmetic Pattern for Build-Up Language

**HARD RULE: When prompt contains "waterfall," "bridge," "build-up," "show each source," "prove the math," or "demonstrate how" → MUST use contribution/arithmetic archetype, NOT big-number archetype.**

Required elements:
- **Each source labeled explicitly**: "Platform savings: $3.1M", "Revenue uplift: $1.8M", "License consolidation: $0.4M"
- **Each value tied to named source**: No generic placeholders
- **Visual sum structure**: Waterfall chart or build-up table showing additive logic
- **Mathematical flow**: Each component shows how it contributes to headline metric

**NEVER use big-number archetype when build-up language is present** — the user is explicitly requesting component breakdown logic, not hero display.

### Business Update Title Slides — Preserve Prompt Fidelity

**MANDATORY for business update presentations**: Title slides must preserve the core message from user prompts and include headline results, NOT generic meeting labels.

**Required title content for business updates:**
- **Include headline metric**: Specific result that answers "what happened?"
- **Add comparison + implication**: Context showing performance vs. baseline
- **Preserve prompt subject**: Don't drift from user's core message

**WRONG — Generic meeting title:**
- "Quarterly Business Review"
- "Financial Update - Q4 2024"

**CORRECT — Action title with headline result:**
- "Q4 revenue beat forecast by 8%, driving record profitability"
- "Cost reduction program delivered $5.2M savings, exceeding target"

## Core Execution Rules

### Execute Immediately

Do NOT ask clarifying questions for standard operations. Proceed with reasonable defaults:

- **Slide creation**: Create the slide and report what was built
- **Shape operations**: Execute and report results
- **Formatting**: Apply formatting and confirm completion

**When to ask**: Only when the request is genuinely ambiguous (e.g., "make it better" without specifying what).

### Ask About PowerPoint Visibility

When starting a multi-step task, **ask the user** whether they want PowerPoint visible or hidden:

> **Watch me work** — Show PowerPoint side-by-side so you see every change live. Operations run slightly slower because PowerPoint renders each update on screen.
>
> **Work in background** — Keep PowerPoint hidden for maximum speed. You won't see changes until the task is done, but operations complete faster.

**Skip asking** when the user has already stated a preference:
- User says "show me PowerPoint", "let me watch" → Show immediately
- User says "just do it", "work in background" → Keep hidden
- Simple one-shot operations (e.g., "how many slides?") → Keep hidden, no need to ask

**How to show PowerPoint:**
```
1. window(action: 'show')                         → Make visible
2. window(action: 'arrange', preset: 'left-half') → Position for side-by-side
```

### Design-First Workflow

Before building slides, query the design catalog for the right configuration:

```
1. design(get-context-model)                    → Determine density from meeting type
2. design(get-style-profile, profileId='...')    → Get fonts, sizes, spacing
3. design(get-palette, paletteId='...')          → Get color hex values
4. design(get-archetype, archetypeId='...')      → Get the unified archetype view: layout rules, observed subtypes, and concrete sanitized example details when local reference data exists
5. design(get-layout-grid, gridId='...')         → Get exact x/y/w/h positions
```

This replaces reading long reference documents — query only what you need.
Reference examples come back embedded in `get-archetype` as sanitized ids/details; raw filenames and provenance metadata stay in local gitignored reference data.

### External Controller Workflow

If a host or custom client is orchestrating a larger deck build, keep the orchestration OUTSIDE the MCP server:

- Planning, execution, verification, and improvement are **client-controlled phases**
- The MCP server remains a normal request/response tool surface
- Do **not** assume MCP batch execution or subagents are available
- Use one controlling client that issues ordinary sequential MCP calls

Recommended external flow:

1. **Plan locally** — create a structured slide plan from the user task
2. **Execute via MCP** — open/create one presentation and build slides in order
3. **Verify via MCP** — re-read slides, inspect shapes/text, and export slide images when helpful
4. **Apply targeted fixes** — adjust only the specific issues found

### Plan Extraction Fallback

External controllers should be tolerant if the model returns a plan as text instead of a perfect structured object.

Accept these patterns:

- JSON object with `{"slides":[...]}`
- JSON object with `{"plan":{"slides":[...]}}`
- Bare JSON array of slide objects
- Markdown blocks like:

```
### Slide 1: Title
- Archetype: executive-summary
- Intent: ...
- Content: ...
```

If a valid structured plan can be recovered from one of these formats, continue with execution rather than failing immediately.

### Format Professionally

When creating presentations:

- Use action titles on every content slide (see slide-design-principles)
- Apply consistent typography hierarchy (title, body, footnote)
- Use colors from the chosen palette only
- Maintain 36pt margins on all sides
- Add source bars to every data slide

### Session Lifecycle

Always close sessions when done:

```
1. file(action: 'open', path: '...')  → sessionId
2. All operations use sessionId
3. file(action: 'close', sessionId: '...', save: true)
```

**Why**: Unclosed sessions leave PowerPoint processes running, consuming memory and locking files.

### CRITICAL: Always End With a Text Response

**NEVER end your turn with only a tool call.** After all operations, provide a text summary.

| Bad (Silent completion) | Good (Text summary) |
|---|---|
| *(tool call with no text)* | "Created 3-card KPI dashboard on slide 2 with Corporate Blue palette." |
| *(just runs a command)* | "Added title slide with dark hero background and accent bar." |

### Format Results as Tables

When presenting data to users, format as Markdown tables, not raw JSON.

## Slide Building Rules

### Shape and Text Operations

When building slides programmatically:

- Create shapes with explicit positions (x, y, w, h in points)
- Use `text(set)` immediately after creating text-containing shapes
- Apply formatting (`text(format)`) after setting text content
- Use `--color` parameter for text color (NOT `--font-color`)
- Use `--alignment` parameter (NOT `--horizontal-alignment`)

### Multi-Line Text

- **MCP Server**: `\n` in JSON strings works correctly for line breaks
- **CLI**: `\n` literal does NOT work in `--text` arguments — use separate textboxes stacked vertically instead

### Table Operations on Slides

When adding tables to slides:
- Use `slidetable(create)` to add tables directly to slides
- Position tables using exact coordinates from layout grids
- Format header rows with bold text and accent color fill
- Right-align numeric columns, left-align text columns

### Chart Operations

When creating charts:
- Use `chart(create)` with appropriate chart type for the data shape
- Position charts using layout grid coordinates
- Always add chart titles and axis labels
- Apply palette colors to chart series
- Add source bars below charts

## Data Modification Rules

### Verify Before Delete

Before deleting slides, shapes, or sections:

1. List existing items first
2. Confirm the exact name/index exists
3. Delete the specified item

### Save Explicitly

Call `file(action: 'close', save: true)` to persist changes:

- Operations modify the in-memory presentation
- Changes are NOT automatically saved to disk
- Session termination WITHOUT save loses all changes

## Error Handling Rules

### Interpret Error Messages

PowerPoint MCP errors include actionable context:

```json
{
  "success": false,
  "errorMessage": "Shape 'Title' not found on slide 1",
  "suggestedNextActions": ["shape(action: 'list', slideIndex: 1)"]
}
```

Follow `suggestedNextActions` when provided.

### Retry with Corrections

If an operation fails:

1. Read the error message carefully
2. Check prerequisites (session open, slide exists, shape exists)
3. Retry with corrected parameters

Do NOT immediately re-run the same failing command.

### Report Failures Clearly

When operations fail:

- State what was attempted
- Explain what went wrong
- Suggest the corrective action

**Good**: "Failed to set text: Shape 'Title 1' not found on slide 2. Use `shape(list, slideIndex=2)` to see available shapes."

**Bad**: "An error occurred."
