# Generation Pipeline

Step-by-step workflow for generating individual slides and complete decks. Follow these steps in order — skipping to layout before clarifying intent is the most common cause of weak slides.

## Pre-Generation Checklist (Per Slide)

Before rendering any slide, produce these outputs in order:

| Step | Output | Validated by |
|------|--------|-------------|
| 1 | **Intent classification** — What must the audience understand or decide? | Maps to one archetype (use `design(list-archetypes)`) |
| 2 | **Context determination** — Meeting type, audience level, consumption mode | Maps to a density profile (use `design(get-context-model)`) |
| 3 | **Headline draft** — Action title in one sentence | Passes title construction rules (see slide-design-principles) |
| 4 | **"So what" statement** — Why this slide matters to the audience | Differs from headline (headline = finding, "so what" = implication) |
| 5 | **Evidence type selection** — Visual/data format that best proves the claim | Matches data shape to visual (see reference tables below) |
| 6 | **Key focal point** — Single element viewer's eye goes to first | Achievable with one accent color |
| 7 | **Footer specification** — Source/date/notes for this density level | Meets footer requirements (see slide-design-principles) |

## Post-Generation Validation

After rendering, apply these checks:

1. **3-second test** — Show slide for 3 seconds. Can the point be understood? If not, headline is weak or visual hierarchy is broken.
2. **Headline-evidence alignment** — Does every element support the headline? Remove anything that does not.
3. **Accent audit** — Is accent color used only for the focal point? If multiple things are accented, reduce to one.
4. **Footer audit** — Are notes present and visually subordinate? Missing source on a data slide = reject.
5. **Speaker-independence test** — Can the slide be understood without a presenter? Must match intended consumption mode.
6. **Scorecard calculation** — Score all eight dimensions. Apply thresholds (see slide-design-review guide).

### Visual Execution Checks (CRITICAL)

These checks catch the most common builder mistakes. Apply AFTER placing all elements:

7. **Overlap scan** — Review every element pair. No shape should overlap another unless intentionally layered (e.g., text on a colored background box). If text overlaps text, fix immediately.
8. **Readability test** — Can ALL text be read? Check specifically:
   - Title text: minimum 20pt (visible from back of room)
   - Body text: minimum 11pt (readable on screen)
   - Labels/annotations: minimum 9pt (readable up close)
   - Source/footnote: minimum 8pt
   - If text is smaller than these minimums, enlarge the text box or reduce content
9. **Space balance** — The slide should use the full content zone (y=80 to y=490). If content only fills the top half, expand elements or add supporting content. If content overflows below y=490, reduce element sizes or split the slide.
10. **Alignment check** — Elements in the same row should share the same y-position. Elements in the same column should share the same x-position. Misalignment by more than 4pt is visible and looks sloppy.
11. **Arrow/connector quality** — Arrows should be thin and subordinate to content (max 2.5pt weight for connectors). Block arrows should be smaller than the content they connect. All arrows in a diagram must be the same size and color.

## Deck Generation Workflow

For full decks, follow this expanded sequence:

```
1. Clarify context (meeting type, audience, mode)
2. Select deck sequence (use `design(get-deck-sequence)`)
3. Determine density profile (use `design(get-density-profile)`)
4. Draft headline outline (all slide headlines in sequence)
5. Validate headline flow (see deck-architecture: headline flow validation in `design(get-deck-sequence)`)
6. Generate individual slides (following pre-generation checklist)
7. Validate each slide (post-generation validation)
8. Validate deck as a whole (deck-level checks)
```

### External Client Orchestration Pattern

If an external controller or agent package is driving the workflow, use this shape:

1. **Plan phase** — derive a slide list with `{index, title, archetypeId, intent, content}`
2. **Execution phase** — build the deck slide-by-slide through ordinary MCP calls
3. **Verification phase** — inspect the generated deck and fix obvious issues with targeted edits

Important constraints:

- Do **not** assume MCP batch execution exists
- Do **not** assume subagents exist
- Keep planning/execution/verification as logical phases in one controlling client
- Reuse one PowerPoint session during execution whenever possible

## Required Inputs: Full Deck

| Input | Required? | Default if missing |
|-------|----------|-------------------|
| Objective | Yes | -- (must be specified) |
| Meeting type | Yes (or infer) | M02 (Executive steering) |
| Audience level | Yes (or infer) | L3 (SVP/VP) |
| Consumption mode | Yes (or infer) | From context matrix |
| Decision required | Strongly recommended | None stated |
| Source material / facts | Recommended | Placeholder structure |
| Slide count target | Optional | From deck sequence template |
| Brand / color constraints | Optional | Default neutral palette |

## Required Inputs: Single Slide

| Input | Required? | Default if missing |
|-------|----------|-------------------|
| Slide intent | Yes | -- |
| Archetype | Yes (or infer from intent) | Key Takeaway |
| Density profile | Yes (or derive from context) | D2 (Clean) |
| Headline | Provided or generated | Server generates |
| Data / evidence | Recommended | Placeholder |
| Source information | Recommended | "Source: [to be added]" |

## Data Shape to Visual Type Mapping

| Data shape | Primary visual | Secondary visual | Avoid |
|-----------|---------------|-----------------|-------|
| Single time series | Line chart | Bar chart (discrete periods) | Pie chart |
| Multiple time series (2-3) | Multi-line chart | Grouped bar | Stacked area |
| Multiple time series (4+) | Line with emphasis | Small multiples | Everything on one chart |
| Categorical comparison (2-8) | Horizontal bar | Grouped vertical bar | Pie chart |
| Categorical comparison (9+) | Table | Dot plot | Bar chart (too many bars) |
| Part of whole (one period) | Stacked bar | 100% bar, Treemap | Pie chart (>5 slices) |
| Part of whole (over time) | Stacked bar over time | 100% stacked bar | Stacked area |
| Movement A to B | Waterfall / bridge | Bullet chart | Pie chart |
| Correlation / tradeoff | Scatter plot | Bubble chart | Line chart |
| Distribution | Histogram | Box plot | Bar pretending to be histogram |
| Binary status (on/off track) | Scorecard / RAG | Bullet chart | Pie chart |
| Hierarchy / decomposition | Value tree / org chart | Treemap | Flat list |
| Sequence / flow | Process diagram | Swimlane | Random arrows |
| Timeline / phasing | Gantt-light / milestone | Roadmap bands | Calendar view |
| Geographic | Choropleth map | Bar by region | 3D globe |

## Intent to Archetype Mapping

| Intent | Primary archetype | Backup archetype |
|--------|------------------|-----------------|
| Summarize whole argument | Executive Summary | Recommendation |
| Present single finding | Key Takeaway | Trend or Comparison |
| Explain why something changed | Bridge / Waterfall | Driver Tree |
| Compare options or items | Comparison | KPI Scorecard or 2x2 |
| Prioritize a portfolio | 2x2 Matrix | Comparison (ranked table) |
| Show a workflow or journey | Process | Operating Model |
| Sequence actions over time | Roadmap | Process |
| Request a decision | Recommendation | Executive Summary |
| Flag risks | Risk / Mitigation | Scorecard |
| Report performance status | KPI Scorecard | Trend |
| Describe structure or model | Operating Model | Driver Tree |
| Provide backup detail | Appendix Evidence | Detailed table |
