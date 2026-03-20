# Evaluation Scoring Criteria

Score each simulated presentation output on these 10 dimensions (0-2 each, max total 20).

## Scoring Scale

| Score | Meaning |
|---|---|
| 0 | Missing or fundamentally wrong |
| 1 | Present but weak or inconsistent |
| 2 | Strong, professional quality |

## Dimensions

### 1. Action Titles (0-2)
- **0**: Topic labels ("Revenue Overview", "Next Steps")
- **1**: Some action titles but inconsistent, or titles too vague
- **2**: Every slide has a complete-sentence action title stating the takeaway

### 2. One Message Per Slide (0-2)
- **0**: Multiple competing messages crammed on slides
- **1**: Most slides focused but some overloaded
- **2**: Every slide communicates exactly one clear idea

### 3. Archetype Selection (0-2)
- **0**: Wrong slide types for the content (e.g., bullets where a chart belongs)
- **1**: Mostly appropriate but some missed opportunities
- **2**: Every slide uses the optimal archetype for its content

### 4. Whitespace & Margins (0-2)
- **0**: Crowded, no breathing room, elements touching edges
- **1**: Acceptable spacing but inconsistent or tight in places
- **2**: Generous margins (36pt+), consistent gaps, slides breathe

### 5. Typography Hierarchy (0-2)
- **0**: All text same size/weight, no visual hierarchy
- **1**: Some hierarchy but inconsistent sizes or too many fonts
- **2**: Clear 3-level hierarchy (title/body/footnote), single font family, consistent

### 6. Color Consistency (0-2)
- **0**: Random colors, no palette, clashing combinations
- **1**: Generally consistent but some off-palette colors or too many colors
- **2**: Strict adherence to one palette, colors serve clear roles

### 7. Content Density (0-2)
- **0**: Walls of text, >6 bullets, >15 words per bullet
- **1**: Slightly overloaded but readable
- **2**: 3-5 bullets, concise text, appropriate chart complexity

### 8. Layout & Alignment (0-2)
- **0**: Elements randomly placed, misaligned, no grid
- **1**: Mostly aligned but some inconsistencies between slides
- **2**: All elements on grid, consistent positions, professional alignment

### 9. Source Citations (0-2)
- **0**: Data claims with no sources
- **1**: Some sources but inconsistent placement
- **2**: All data slides have source bar, consistent formatting

### 10. Overall Professionalism (0-2)
- **0**: Would not present to external audience
- **1**: Acceptable for internal use
- **2**: Consulting/executive quality, would present to a board or client

## Interpreting Scores

| Total Score | Quality Level | Action |
|---|---|---|
| 0-8 | Poor | Major skill gaps, fundamental redesign needed |
| 9-12 | Acceptable | Specific areas need improvement |
| 13-16 | Good | Minor polish needed |
| 17-20 | Excellent | Consulting-grade output |

## Gap Categories

When scoring reveals patterns, classify the gap:

| Gap Type | Example | Fix Location |
|---|---|---|
| Missing guidance | LLM doesn't know to use action titles | `slide-design-principles.md` |
| Wrong archetype | Uses bullets when chart needed | `design(get-archetype)` tool |
| Poor colors | Random colors, no palette selection | `design(get-palette)` tool |
| Bad layout | Elements poorly positioned | `design(get-layout-grid)` tool |
| No self-review | Obvious issues not caught | `slide-design-review.md` |
| Profile mismatch | Sales deck looks like academic paper | `design(get-style-profile)` tool |
