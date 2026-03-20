# Slide Design Principles

Universal rules for professional PowerPoint slides. Apply these to EVERY presentation regardless of style.

## Slide Anatomy

Every slide has four zones. Respect them:

```
┌─────────────────────────────────────────────┐
│ ACTION TITLE (full sentence, ≤15 words)     │  ← 10% height
├─────────────────────────────────────────────┤
│                                             │
│              CONTENT AREA                   │  ← 75% height
│         (one idea, one visual)              │
│                                             │
├─────────────────────────────────────────────┤
│ Source: [citation]                          │  ← 10% height
├─────────────────────────────────────────────┤
│ Page # │ Section │ Date                     │  ← 5% height
└─────────────────────────────────────────────┘
```

## Action Titles (CRITICAL)

NEVER use topic labels. ALWAYS use action titles — complete sentences stating the takeaway.

| BAD (topic label) | GOOD (action title) |
|---|---|
| "Revenue Overview" | "Revenue grew 18% in Q1, outpacing projections by 5pp" |
| "Market Analysis" | "Three emerging markets account for 60% of growth potential" |
| "Next Steps" | "Consolidating suppliers in Q3 will unlock $50M annually" |
| "Team Structure" | "Engineering headcount must double to meet 2027 roadmap" |

Rules:
- Maximum 15 words, never more than 2 lines
- States the conclusion, not the topic
- Executive should grasp the full argument by reading only titles

**Making titles SHARP (executive-grade):**

Good titles state facts. Great titles state implications and trade-offs:

| Good (factual) | Great (insight-led) |
|---|---|
| "Revenue grew 12% in Q4" | "Q4 revenue growth of 12% validates the enterprise pivot — sustaining this requires doubling APAC investment" |
| "Three priorities for Q1" | "Three priorities will determine whether we hit 20% growth or stall at 12% — all require January action" |
| "Customer churn increased" | "Churn spike to 8% signals pricing pressure — matching competitor mid-tier preserves 60% of at-risk revenue" |

Rules for sharp titles:
- Include the specific number AND its implication
- State the trade-off or decision, not just the observation
- Answer "so what?" and "what should we do?" in the title itself
- If you can remove a word without losing meaning, remove it

**Exceptions where topic titles are acceptable:**
- Title slide (first slide) — uses the presentation topic name
- Section divider slides — uses the section name
- Appendix slides — uses "Appendix: [topic]"

**Action title formatting:**
- Position: top of slide, full width (x=36, y=24, w=888)
- Font: 20pt bold, Text Dark color
- Include specific numbers when available ("$45M", "12%", "3x")
- Separate the action title zone from content with 24pt gap minimum
- Add a thin accent-colored underline bar (w=80, h=3) beneath the title for visual weight

## One Message Per Slide

Each slide communicates exactly ONE idea. If you have two messages, make two slides.

Test: Can you summarize this slide in one sentence? If not, split it.

## Typography Hierarchy

Use a single sans-serif font family throughout (Calibri, Arial, or Segoe UI).

| Element | Size | Weight | Color |
|---|---|---|---|
| Action title | 20-24pt | Bold | Dark (near-black) |
| Section header | 18-20pt | Bold | Primary brand color |
| Body text | 14-18pt | Regular | Dark grey (#333333) |
| Bullet points | 14-16pt | Regular | Dark grey |
| Callout/highlight | 16-20pt | Bold | Accent color |
| Source/footnote | 9-10pt | Regular | Medium grey (#666666) |
| Footer | 8-9pt | Regular | Light grey (#999999) |

Rules:
- Maximum 3 font sizes per slide (title + body + footnote)
- Never use more than 2 weights (Regular + Bold)
- No italic unless for proper names or publications
- No underline (it implies hyperlinks)
- No ALL CAPS in body text (titles only, sparingly)

## Whitespace

Whitespace is not empty space — it is a design element that creates clarity.

- Margins: minimum 0.5 inches (36pt) on all sides
- Between title and content: 18-24pt gap
- Between content blocks: 14-18pt gap
- Between bullet items: 6-8pt spacing
- If a slide feels crowded, SPLIT IT — never reduce margins

### Overlap Prevention (CRITICAL)

NEVER allow elements to overlap. Overlapping shapes/text is the most common visual error.

**Before placing any element, calculate available space:**
1. Title zone: y=20 to y=70 (50pt reserved)
2. Content zone: y=80 to y=490 (410pt available)
3. Source/footer zone: y=495 to y=540 (45pt reserved)
4. Left margin: x=36. Right margin: x=924. Usable width: 888pt.

**Spacing rules between elements:**
- Minimum gap between any two elements: 8pt (NEVER 0pt or negative)
- Text boxes: leave 4pt padding inside each box so text doesn't touch edges
- When laying out grids: calculate total width = (N items x item width) + ((N-1) x gap width). If total > available width, reduce item width or reduce N.
- When stacking vertically: calculate total height = sum of all heights + gaps. If total > 410pt content zone, reduce element heights or split to two slides.

**Common overlap scenarios and fixes:**
- Table/grid cells: calculate column widths from available width BEFORE creating cells
- KPI cards in a row: 3 cards x 280pt = 840pt, leaving only 48pt for 2 gaps. Use 270pt cards + 24pt gaps instead.
- Flow diagrams: connector lines between steps — position arrows in the GAP between shapes, not on top of them
- Long text in small boxes: either enlarge the box or shorten the text

## Content Density

| Element | Maximum per slide |
|---|---|
| Bullet points | 5 (prefer 3) |
| Words per bullet | 15 (prefer 8-10) |
| Data series in chart | 5 (prefer 3) |
| Columns in table | 6 (prefer 4) |
| Rows in table | 8 (prefer 5-6) |

### Vertical Space Management

The bottom 30-40% of a slide should never be completely empty. If your content only fills the top half:
- **Expand the hero element** (chart, big number) to fill more vertical space
- **Add supporting content** below: a quote, a comparison row, a mini-table, or insight callout
- **Vertically center** the content block if the content is intentionally minimal (e.g., quote slides, big number slides)
- **Lower the content start** — move the content zone down so it's vertically balanced
- **Add a bottom accent bar** (x=0, y=520, w=960, h=20, fill=Primary) to anchor the page visually

### Avoid Orphaned Elements

When placing a callout box, insight bar, or summary beneath cards/charts:
- Keep the gap between the cards and the callout to **12-18pt maximum** — not 40+pt
- The callout should visually "attach" to the cards above, not float in isolation
- If there is too much vertical space between content blocks, either enlarge the cards or move the callout up
- Rule: if you can fit another element between two content blocks, the gap is too large

## Visual Hierarchy

Guide the eye in this order:
1. Title (top) — the takeaway
2. Hero element (center) — chart, big number, or framework
3. Supporting detail (secondary) — bullets, annotations
4. Source (bottom) — credibility

Use size, weight, and color to reinforce this hierarchy. The most important element should be the largest.

## Alignment

- All elements snap to a consistent grid
- Left-align text (never center-align body text)
- Center-align only: titles, hero numbers, single-line captions, KPI card values
- Right-align only: numbers in table columns
- Consistent margins on every slide — never shift margins between slides

## Page Numbers and Footers

**Every content slide (not title or section dividers) must have:**
- Page number: bottom-right corner (x=900, y=512, right-aligned, 9pt, Text Light)
- Optional section name: bottom-left (x=36, y=512, 9pt, Text Light)

Use the `headerfooter` command to set page numbers and footer text on all slides.

**Source bars (MANDATORY on data slides):**
- Position: x=36, y=490, w=888, h=15
- Font: 9-10pt, Text Medium (#4A5568) — NOT light grey, must be readable
- Format: "Source: [specific system/report name], [specific date/period]"
- Always include on slides with charts, numbers, metrics, or quantitative claims
- Missing sources = zero credibility in executive settings

**Source bar examples (strong vs weak):**

| Weak | Strong |
|---|---|
| "Source: internal data" | "Source: Finance ERP system, December 2025 actuals" |
| "Source: company reports" | "Source: incident log, performance telemetry — Feb 2026" |
| "Source: research" | "Source: Retail Analytics Survey 2024; Industry Digital Retail Report 2024" |

**Rule: Every slide with a number must have a source bar.** The only exceptions are title slides and CTA slides.

## Title Type Codes

Every slide title falls into one of these semantic categories. Match the title type to the slide's communicative intent:

| Type | Code | When to use | Structure |
|------|------|-------------|-----------|
| Insight | T-INS | Diagnosing, explaining, revealing | "[Subject] [verb] [finding], [driven by/because/despite] [cause]" |
| Recommendation | T-REC | Asking for action or approval | "[Action verb] [what] to [achieve outcome]" |
| Comparison | T-CMP | Evaluating alternatives | "[A] outperforms on [X], but [B] wins on [Y]" |
| Risk | T-RSK | Flagging uncertainty | "Without [X], [risk materializes]" |
| Roadmap | T-RMP | Explaining sequence and timing | "The sequence: [phase logic]" |
| Performance | T-PRF | Assessing against plan | "[Subject] is [on/off track]: [X] of [Y] are [status]" |
| Composition | T-MIX | Describing proportions | "[Part] now represents [X]% of [whole], up from [Y]%" |

### Title Construction Rules

1. Must be a complete thought, not a noun phrase
2. Conclusion-first: insight comes first, context follows
3. Tied to evidence on the slide -- if title says "pricing" but chart shows volume, there is a disconnect
4. Target 8-15 words. If >20 words, split into title + subtitle
5. Sentence case (not Title Case)
6. Active verbs: "Revenue grew 12%" not "12% revenue growth was observed"
7. Quantify where possible: "Margins improved 180bps to 14.2%" beats "Margins improved"
8. No question titles unless uncertainty IS the message

## Subtitle Types

| Role | Code | Purpose | Example |
|------|------|---------|---------|
| Scope | SH-SCP | Describe the data shown | "Revenue by region, EUR m, FY2022-2025" |
| Section label | SH-SEC | Label a deck section | "Diagnosis" or "Supporting evidence" |
| Methodology | SH-MTH | State analytical approach | "Based on n=1,250 enterprises across 41 dimensions" |
| Exhibit label | SH-EXH | Exhibit-style numbering | "Exhibit 4" |

The subtitle must never restate the headline. If headline says "Revenue grew 12%", subtitle should be "Group revenue, EUR bn, FY2020-2025" (scope), not "Revenue growth analysis."

## Footer System by Density

Footer content scales with density profile:

| Density | Footer content |
|---------|---------------|
| D1 (Minimal) | Source name only, or none for non-data slides |
| D2 (Clean) | Source + as-of date |
| D3 (Structured) | Source + as-of date + base/sample + 1 assumption/caveat |
| D4 (Detailed) | Source + date + base + definitions + assumptions + caveats |
| D5 (Dense) | Full provenance chain including methodology description |

Footer fields:

| Field | Required? | Format | Example |
|-------|----------|--------|---------|
| Source | Always (data slides) | "Source: [institution]" | "Source: Eurostat, ECB" |
| As-of date | Always (data slides) | "As of [date]" | "As of December 31, 2025" |
| Base / sample | When not obvious | "Base: [description]" | "Base: n=1,250 enterprises" |
| Metric definition | When non-standard | "Note: [metric] defined as [def]" | "Note: EBITDA adjusted for restructuring" |
| Assumptions | When material | "Assumes [assumption]" | "Assumes constant FX (EUR/USD 1.08)" |
| Caveats | When material | "Excludes [limitation]" | "Excludes discontinued operations" |

## Visual Zone Model

Every slide has three semantic zones. Respect their proportions:

```
ZONE 1: HEADLINE (top 12-15% of slide)
  Action title + optional subtitle
  ─────────────────────────────────
ZONE 2: EVIDENCE (center 55-70% of slide)
  Primary chart, diagram, or text block
  Annotations, callouts, key numbers
  ─────────────────────────────────
ZONE 3: FOOTER (bottom 8-12% of slide)
  Source, date, notes, caveats
```

Minimum spacing between zones: 3% of slide height. Space between elements within evidence zone: 2% minimum.

## Contrast and Readability

- Text-to-background contrast ratio ≥ 4.5:1 (WCAG AA)
- Never place text on busy images without a semi-transparent overlay
- Dark text on light backgrounds (default)
- Light text on dark backgrounds only for hero/accent slides
- Avoid red text (except for negative values in financial context)

### Arrows and Connectors (CRITICAL for flow/process slides)

Arrows are the most visible structural element. Bad arrows ruin otherwise good slides.

**Selection rules:**
- Flow between steps: thin connectors (1.5-2pt) or small block arrows (w=28-36pt)
- Timeline backbone: thin rectangle (h=4-6pt), NOT a fat block arrow
- Trend up/down: small triangle or arrow shape (16-20pt), filled with status color
- Transformation bridge (A to B): medium right arrow (w=48pt), accent color
- Text direction: use right arrow character in text, not a shape

**Styling rules:**
- Arrow shapes should be LESS prominent than content they connect
- All flow arrows on one slide: SAME size, SAME color, SAME weight
- Max arrow thickness for connectors: 2.5pt (never >3pt)
- Block arrows for flow: max 36pt wide (larger = too dominant)
- Color: Primary at 50-80% opacity or neutral grey, never bright/saturated
- NEVER use more than 2 arrow colors on one slide

**Common mistakes to avoid:**
- Oversized block arrows (>60pt) that dominate the content
- Mismatched arrow sizes within the same diagram
- Thick heavy connector lines that look unprofessional
- Arrows overlapping text labels
- Using block arrows where thin lines would suffice

### Dark Background Slide Rules (CRITICAL)

When using a dark background (title slides, section dividers, CTA slides):
- Title text: **always white (#FFFFFF)**, never dark colors
- Subtitle text: **white at 70-80% opacity** or **very light grey (#B0C4DE, #D0D8E0)**
- NEVER use accent colors (red, orange, yellow) as subtitle text on dark backgrounds — contrast is insufficient
- Accent colors on dark backgrounds: use only for **shapes, bars, and badges** — not body text
- Footer/date text: white at 50-60% opacity
- Test: if you squint and can't read it, the contrast is wrong

### Bullet Points and List Formatting

When presenting a list of items (pain points, features, recommendations):
- ALWAYS add visual markers — never plain text lines
- Option 1: Unicode bullets ("•") as prefix with 8pt indent
- Option 2: Small circle shapes (8-10pt diameter) as markers, left of each line
- Option 3: Number badges (for ordered lists) using circle + number
- Option 4: Check marks ("✓") for completed/included items
- Spacing: 12-16pt between bullet items for scannability
- Each bullet should be a self-contained point (no continuation across bullets)
