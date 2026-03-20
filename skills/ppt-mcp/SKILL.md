---
name: ppt-mcp
description: >
  Automate Microsoft PowerPoint on Windows via COM interop. Use when creating, reading,
  or modifying PowerPoint presentations. Supports Slides, Shapes, Text, Charts, Tables,
  Animations, Transitions, Notes, and Media.
  Triggers: PowerPoint, presentation, pptx, slides, shapes, animations.
---

# PowerPoint MCP Server Skill

Provides 218 PowerPoint operations via Model Context Protocol. The MCP Server forwards all requests to the shared PptMcp Service, enabling session sharing with CLI. Tools are auto-discovered - this documents quirks, workflows, and gotchas.

## Workflow Checklist

| Step | Tool | Action | When |
|------|------|--------|------|
| 1. Open file | `file` | `open` or `create` | Always first |
| 2. Create slides | `slide` | `create`, `duplicate` | If needed |
| 3. Add shapes | `shape` | `create`, `set-text` | Add visual elements |
| 4. Add charts | `chart` | `create` | Visualize data |
| 5. Format | `shape` | `format`, `position` | After adding content |
| 6. Save & close | `file` | `close` with `save: true` | Always last |

## Preconditions

- Windows host with Microsoft PowerPoint installed (2016+)
- Use full Windows paths: `C:\Users\Name\Documents\Report.pptx`
- PowerPoint files must not be open in another PowerPoint instance

## Large Deck Pattern

For larger builds, use a **single controlling client** with these logical phases:

1. Plan
2. Execute
3. Verify
4. Improve

Important:

- Keep this orchestration in the client, not in the MCP server
- Use normal sequential MCP calls
- Do **not** assume MCP batch execution or subagents are available

## Slide Transitions Workflow

Use `transition` for **slide transition effects**. Apply transitions to individual slides or across the entire presentation:

```
1. transition(action: 'set', slideIndex: 1, effect: 'push')  → Set transition
2. Add content to slides (shapes, text, charts)
3. transition(action: 'set', slideIndex: 2, effect: 'fade')  → Set next slide transition
4. animation(action: 'add', slideIndex: 1, shapeName: 'Title', effect: 'fly-in')  → Animate shapes
```

**Note:** Transitions apply between slides. Animations apply to individual shapes within a slide.

## CRITICAL: Execution Rules (MUST FOLLOW)

### Rule 1: NEVER Ask Clarifying Questions

**STOP.** If you're about to ask "Which file?", "What table?", "Where should I put this?" - DON'T.

| Bad (Asking) | Good (Discovering) |
|--------------|-------------------|
| "Which PowerPoint file should I use?" | `file(list)` → use the open session |
| "What shapes are on this slide?" | `shape(list)` → discover shapes |
| "Which slide has the content?" | `slide(list)` → check all slides |
| "Should I add an animation?" | YES - add it and apply appropriate timing |

**You have tools to answer your own questions. USE THEM.**

### Rule 2: Always End With a Text Summary

**NEVER end your turn with only a tool call.** After completing all operations, always provide a brief text message confirming what was done. Silent tool-call-only responses are incomplete.

### Rule 3: Design Slides Professionally

Apply consistent formatting across slides:

| Element | Property | Example |
|---------|----------|---------|
| Title | Font size | 28pt+ |
| Body text | Font size | 18-24pt |
| Shape fill | Color | Brand colors |
| Transitions | Duration | 0.5-1.5s |

**Workflow:**
```
1. shape create (add visual elements)
2. shape format (apply colors, borders)
3. text set (add content)
```

### Rule 4: Use Slide Layouts

Always use appropriate slide layouts for consistent design:

```
1. slide(action: 'create', layout: 'Title and Content')  → Structured slide
2. shape(action: 'set-text', shapeName: 'Title')  → Set title text
3. shape(action: 'set-text', shapeName: 'Content')  → Set body content
```

**Why:** Slide layouts provide consistent positioning, fonts, and structure.

### Rule 5: Session Lifecycle

```
1. file(action: 'open', path: '...')  → sessionId
2. All operations use sessionId
3. file(action: 'close', save: true)  → saves and closes
```

**Unclosed sessions leave PowerPoint processes running, locking files.**

### Rule 6: Slide Masters and Layouts

Slides inherit formatting from slide masters:

```
Step 1: Choose layout → slide(action: 'create', layout: 'Title Slide')
Step 2: Set content → shape(action: 'set-text', shapeName: 'Title', text: '...')
Step 3: Customize → shape(action: 'format', fillColor: '#0078D4')
```

### Rule 7: Animation Sequence

**BEST PRACTICE: Build animations in logical order**

```
1. shape(action: 'create', ...) → Add shapes first
2. animation(action: 'add', effect: 'fade-in', order: 1) → Entrance effect
3. animation(action: 'add', effect: 'emphasis', order: 2) → Emphasis effect
4. transition(action: 'set', effect: 'push') → Slide transition
```

**Why add shapes first:**
- Shapes must exist before animations can be applied
- Animation order determines playback sequence
- Transitions are separate from shape animations

### Rule 8: Targeted Updates Over Delete-Rebuild

- **Prefer**: Modifying shape properties directly (text, color, position)
- **Avoid**: Deleting and recreating entire slides or shapes

**Why:** Preserves animations, transitions, and layout relationships.

### Rule 9: Follow suggestedNextActions

Error responses include actionable hints:
```json
{
  "success": false,
  "errorMessage": "Shape 'Title' not found on slide 1",
  "suggestedNextActions": ["shape(action: 'list', slideIndex: 1)"]
}
```

### Rule 10: Use Consistent Styling Across Slides

When building multi-slide presentations, maintain consistent colors, fonts, and positioning:

```
1. Use slide layouts for structure consistency
2. Apply matching colors to shapes across slides
3. Keep title positions and sizes uniform
```

**When NOT needed:** Single-slide modifications or quick text updates.

## Tool Selection Quick Reference

| Task | Tool | Key Action |
|------|------|------------|
| Create/open/save presentations | `file` | open, create, close |
| Create/manage slides | `slide` | create, duplicate, delete |
| Add/modify shapes | `shape` | create, format, position |
| Set text content | `text` | set, get |
| Create charts | `chart` | create, update |
| Add tables to slides | `table` | create |
| Set animations | `animation` | add, remove, reorder |
| Set slide transitions | `transition` | set, remove |
| Add speaker notes | `notes` | set, get |
| Visual verification | `export` | slide-to-image |

## Reference Documentation

See `references/` for detailed guidance:

- [Core execution rules and LLM guidelines](./references/behavioral-rules.md)
- [Slide design principles](./references/slide-design-principles.md)
- [Slide design review checklist](./references/slide-design-review.md)
- [Generation pipeline](./references/generation-pipeline.md)
- [Agent mode patterns](./references/ppt_agent_mode.md)

For design catalog data (archetypes, palettes, grids, styles, density), use the `design` tool:
- `design(list-archetypes)` / `design(get-archetype)` — Primary unified archetype surface with curated layout guidance plus learned subtypes and concrete sanitized example details when local reference data is available
- `design(list-palettes)` / `design(get-palette)` — Color palettes with hex values
- `design(list-layout-grids)` / `design(get-layout-grid)` — Exact positioning coordinates
- `design(list-style-profiles)` / `design(get-style-profile)` — Style configurations
- `design(get-context-model)` — Meeting type → density mapping
- `design(get-deck-sequence)` — Canonical slide sequences
- `design(get-slide-patterns)` — Advanced consulting patterns
- `design(get-icon-shapes)` — Native shape icon catalog

Reference examples are exposed only as sanitized ids/details embedded in `design(get-archetype)`. Raw filenames and source provenance remain in local gitignored reference data and never appear in MCP responses.
