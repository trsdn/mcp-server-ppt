# Slide Design Self-Review

After creating slides, run this checklist to judge and improve your output. Fix any failures before presenting the result to the user.

## Quick Check (Every Slide)

Run these 10 checks on EVERY slide you create:

| # | Check | Pass Criteria | Fix |
|---|---|---|---|
| 1 | Action title? | Title is a complete sentence stating a conclusion, not a topic label | Rewrite: "Revenue Overview" → "Revenue grew 18% in Q1, exceeding targets" |
| 2 | One message? | Slide communicates exactly one idea | Split into multiple slides |
| 3 | Content density? | ≤5 bullets, ≤15 words each, ≤5 chart series | Remove or split content |
| 4 | Typography hierarchy? | Max 3 font sizes, consistent weight usage | Standardize sizes |
| 5 | Whitespace? | ≥36pt margins, no crowded areas, elements breathe | Add spacing or split slide |
| 6 | Color consistency? | All colors from chosen palette, ≤4 colors per slide | Replace off-palette colors |
| 7 | Alignment? | All elements on grid, no visual misalignment | Snap to grid positions from `design(get-layout-grid)` |
| 8 | Right archetype? | Slide type matches the content being presented | Switch to appropriate archetype |
| 9 | Source cited? | Data slides have source footnotes | Add source bar |
| 10 | Readable at distance? | Minimum 14pt body text, sufficient contrast | Increase font size |

## Deck-Level Check (Full Presentation)

| # | Check | Pass Criteria | Fix |
|---|---|---|---|
| 11 | Consistent style? | Same fonts, colors, margins across ALL slides | Standardize to chosen profile |
| 12 | Logical flow? | Slides follow a narrative arc (context → analysis → recommendation) | Reorder slides |
| 13 | Title story? | Reading ONLY the titles tells the complete story | Rewrite titles to form narrative |
| 14 | Section breaks? | Sections divided with divider slides for decks >8 slides | Add section dividers |
| 15 | Opening strong? | First slide clearly states purpose/topic | Revise title slide |
| 16 | Closing actionable? | Last content slide has clear next steps or recommendations | Add recommendations slide |
| 17 | Slide count? | Not too many (audience fatigue) or too few (incomplete) | Target: 1 slide per 2-3 minutes of talk time |
| 18 | Variety? | Not all slides are the same archetype | Mix archetypes |
| 19 | Data-to-text ratio? | Balance of visual (charts, diagrams) and text slides | Aim for 50/50 or more visual |
| 20 | Appendix? | Supporting detail in appendix, not cluttering main deck | Move backup data to appendix |

## The Title Story Test

The most important check. Read ONLY the slide titles in sequence. They should form a coherent argument:

**GOOD title story:**
1. "Customer churn increased 15% in Q4, driven by three root causes"
2. "Competitor pricing undercuts our mid-tier by 20%"
3. "Onboarding friction causes 40% of first-month cancellations"
4. "Support response times doubled after the September restructuring"
5. "Three initiatives can reduce churn to pre-Q3 levels within 6 months"
6. "Priority 1: Match competitor pricing on mid-tier plans (impact: -5pp churn)"

**BAD title story:**
1. "Introduction"
2. "Background"
3. "Analysis"
4. "Data"
5. "Recommendations"
6. "Next Steps"

If your titles read like the BAD example, rewrite every title as an action title.

## Common Failures and Fixes

### Failure: Wall of Text
**Symptom:** Slide has >6 bullets or >100 words of body text
**Fix:** 
1. Identify the ONE key message
2. Write an action title that states it
3. Keep 3 most important supporting points
4. Move everything else to appendix or next slide

### Failure: Chart Overload
**Symptom:** Multiple charts crammed on one slide, all small
**Fix:**
1. One chart per slide (or 2 if comparing related data)
2. Use 2×2 dashboard grid only when all 4 metrics are equally important
3. Add insight callout next to chart explaining the "so what"

### Failure: Inconsistent Colors
**Symptom:** Different blues, random accent colors, no cohesion
**Fix:**
1. Pick ONE palette via `design(get-palette)`
2. Use Primary for dominant elements, Secondary for supporting, Accent for highlights
3. Never use a color not in the palette

### Failure: Missing Hierarchy
**Symptom:** All text same size and weight, nothing stands out
**Fix:**
1. Title: 20-24pt bold
2. Key number/message: 28-48pt bold, accent color
3. Body: 14-18pt regular
4. Source: 9-10pt grey
5. Most important element should be largest

### Failure: Poor Layout
**Symptom:** Elements feel randomly placed, uneven spacing
**Fix:**
1. Choose a grid via `design(get-layout-grid)`
2. Snap ALL elements to grid positions
3. Verify consistent margins (36pt on all sides)
4. Check alignment: left edges, top edges, spacing between elements

### Failure: Topic Labels Instead of Action Titles
**Symptom:** Titles like "Overview", "Summary", "Analysis", "Q3 Results"
**Fix:** Add the "so what":
- "Overview" → "Three market shifts require immediate strategic response"
- "Summary" → "Consolidation strategy will deliver $12M in annual savings"
- "Q3 Results" → "Q3 revenue exceeded forecast by 8% despite supply constraints"

## Self-Improvement Loop

After completing a presentation:
1. Run Quick Check on each slide (fix immediately)
2. Run Deck-Level Check (fix structure issues)
3. Run Title Story Test (rewrite titles if needed)
4. Take a screenshot of each slide to visually verify
5. Ask: "Would an executive spend more than 3 seconds understanding any slide?" If yes, simplify.

## External Controller Verify-Fix Loop

When a custom client is orchestrating the build, run verification as an explicit follow-up phase:

1. Re-open the generated presentation
2. Use `slide(list)` and `slide(read)` to inspect slide order, titles, and structure
3. Use `shape(list)` and text reads where needed to confirm expected content exists
4. Export slide images with `export(slide-to-image)` when human review artifacts are helpful
5. Apply only **targeted fixes** — do not rebuild the whole deck unless the structure is fundamentally wrong
6. Save and close

Constraints:

- No dependence on MCP batch execution
- No dependence on subagents
- The controller owns the loop; MCP provides the primitive operations

## Quality Scorecard (7-Dimension Assessment)

For rigorous quality evaluation, score each slide on seven dimensions (0-5 scale):

| Dimension | 0 (Fail) | 3 (Acceptable) | 5 (Excellent) |
|-----------|---------|----------------|---------------|
| Claim clarity | No discernible point; topic label only | Point present but could be sharper | Immediately clear from headline alone |
| One-message discipline | Multiple competing messages | One main message with some clutter | Every element supports a single message |
| Evidence fit | Visual contradicts or is irrelevant | Visual supports but does not strongly prove | Optimal format, directly proves the claim |
| Visual economy | Cluttered, no clear focal point | Generally clean with minor noise | Precisely what is needed, nothing more |
| Audience fit | Wrong density for the context | Approximately right density | Perfectly calibrated to audience and mode |
| Traceability | No source on a data slide | Source present but incomplete | Full provenance appropriate to density |
| Narrative fit | Slide is a non-sequitur in deck flow | Slide fits but transition is not smooth | Logically follows predecessor, sets up successor |

**Thresholds:**
- **Pass:** Total >= 24/35 AND no individual dimension below 3
- **Review:** Total 20-23 OR one dimension at 2
- **Reject:** Total below 20 OR any dimension at 0-1

## Auto-Reject Triggers

A slide must be regenerated if ANY of these conditions are true:

**Title failures:**
- Title is a noun phrase / topic label (e.g., "Revenue overview")
- Title exceeds 25 words
- Title contradicts the visual evidence

**Visual failures:**
- More than one chart competes for attention without comparative structure
- Chart type does not match data shape (pie chart for time series)
- More than 2 accent colors used without semantic meaning
- Legend used when direct labels were feasible
- Y-axis bar chart does not start at zero (unless noted)

**Evidence failures:**
- No source or date on a data slide
- Waterfall bars do not sum correctly
- 2x2 axes are unlabeled or vaguely defined
- "Other" is the largest category in a breakdown

**Structural failures:**
- Roadmap has no sequencing logic (just a task list)
- Recommendation slide has no explicit ask
- Executive summary has no recommendation
- Appendix slide has no headline

**Density failures:**
- D1 slide has more than 40 words
- D2 slide has more than 80 words
- Appendix-level detail on a live presentation slide
- No footer on a D3+ data slide

## Deck-Level Validation

Beyond individual slides, validate the deck as a whole:

| Check | Rule |
|-------|------|
| Headline flow | Extract all headlines. Read sequentially. Must form coherent argument. |
| SCQA presence | First 2-4 slides must establish Situation, Complication, and Answer. |
| MECE grouping | Support sections must be mutually exclusive and collectively exhaustive. |
| Density consistency | All body slides within one density level of each other (no D1 next to D4). |
| Color consistency | Same accent color throughout. Same grey palette. No random color shifts. |
| Archetype variety | No more than 3 consecutive slides of the same archetype. |
| Recommendation presence | Decision decks must end with a recommendation slide. |
| Ask clarity | At least one slide must contain an explicit decision request (for decision decks). |
