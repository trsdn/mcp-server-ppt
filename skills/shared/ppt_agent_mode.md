# Agent Mode in PowerPoint — Watch AI Work

PowerPoint MCP's Agent Mode lets users watch AI operations happen in real-time. Instead of hidden automation, users see PowerPoint respond to commands live.

> **MCP Server feature only.** Agent Mode uses conversational UI to ask about visibility preferences.

## When to Offer Agent Mode

**Always ask the user** at session start whether they want PowerPoint visible or hidden:

> **Watch me work** — Show PowerPoint side-by-side so you see every change live. Operations run slightly slower because PowerPoint renders each update on screen.
>
> **Work in background** — Keep PowerPoint hidden for maximum speed. You won't see changes until the task is done, but operations complete faster.

**Skip asking** only when the user has already stated a preference:
- User says "show me", "let me watch", "I want to see" → Show immediately
- User says "just do it", "work in background" → Keep hidden
- Simple one-shot operations (e.g., "how many slides?") → Keep hidden, no need to ask

## Three Workflows

### 1. Agent Mode — Interactive Side-by-Side

User watches AI build a presentation in real-time, side-by-side with the AI assistant.

```
1. file(open, path='report.pptx')
2. window(show)                                    → Make visible
3. slide(create, layoutName='Title and Content')   → User sees slide appear
4. text(set, slideIndex=1, shapeName='Title 1', text='Q1 Report')
5. shape(add-textbox, slideIndex=2, ...)           → User sees textbox appear
6. chart(create, slideIndex=3, ...)                → User sees chart render
7. animation(add, slideIndex=3, shapeName='Chart') → User sees animation added
8. ASK: "I've finished the presentation. Would you like me to save and close?"
```

**Key behaviors:**
- Ask before closing — user may want to inspect or make manual changes
- Narrate what you're building alongside visual changes

### 2. Presentation Mode — Guided Walkthrough

AI navigates through a completed presentation, explaining content.

```
1. file(open, path='analysis.pptx')
2. window(show)
3. window(maximize)                                → Full screen for visibility
4. slide(list)                                     → Get slide overview
5. "Slide 1 is the title slide with 'Annual Report 2026'..."
6. slide(read, slideIndex=2)                       → Read slide details
7. "Slide 2 shows the revenue chart with 4 data series..."
8. text(get, slideIndex=3, shapeName='Content')    → Read text content
9. "Here's what I found: [summary with insights]"
```

### 3. Debug Mode — Step-by-Step Inspection

AI performs operations one at a time for troubleshooting.

```
1. file(open, path='broken.pptx')
2. window(show)
3. slide(list)
4. "Found 12 slides. Let me check each one..."
5. shape(list, slideIndex=1)
6. "Slide 1 has 5 shapes. Checking for issues..."
7. placeholder(list, slideIndex=1)
8. "Title placeholder is empty — should be filled."
9. text(find, searchText='TODO', slideIndex=0)
10. "Found 3 TODO markers across the presentation."
```

## Asking About Visibility

When starting a session, present the visibility choice:

> **Watch me work** — Show PowerPoint side-by-side so you see every change live.
>
> **Work in background** — Keep PowerPoint hidden for maximum speed.

If the user picks "Watch me work":
1. `window(restore)` → Make PowerPoint visible
2. Use narration throughout the workflow for live progress

If the user picks "Work in background", keep PowerPoint hidden.
