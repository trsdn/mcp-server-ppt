# 1b. KPI Card Dashboard (Most Common)

**When:** 2-4 key metrics shown together with context. Used in nearly every business presentation.

**Card anatomy (each card has 4 zones):**
```
┌──────────────────────────────────┐
│ LABEL STRIP (colored header)     │  ← 30pt, Primary or Accent fill, white text, ALL CAPS, 11pt
├──────────────────────────────────┤
│                                  │
│     BIG NUMBER                   │  ← 48-52pt bold, center-aligned, Primary or semantic color
│                                  │
│  context left    context right   │  ← 12pt grey, targets/comparisons
│                                  │
└──────────────────────────────────┘
```

**Color rules for KPI cards:**
- Normal metric: Primary color (navy) for the big number
- Positive trend: Positive color (green) with ▲ prefix
- Negative trend: Negative color (red) with ▼ prefix
- Highlighted/featured metric: Accent color (orange) for BOTH the header strip AND the number
- Status indicators: "● On Track" in green, "● At Risk" in amber (#F39C12), "● Off Track" in red

**CRITICAL: KPI card labels must be readable.** The header strip text (e.g., "REVENUE", "GROSS MARGIN") must be:
- White (#FFFFFF) text on the colored header strip
- 10-11pt, ALL CAPS, bold
- Verify via `text format` with `--color "#FFFFFF"` AFTER setting the shape fill
- If the label is not visible, add a separate white textbox OVERLAID on the header strip instead of using the shape's own text

**2x2 layout (4 KPIs):**
```
Card 1: x=36,  y=85,  w=435, h=195  |  Card 2: x=489, y=85,  w=435, h=195
Card 3: x=36,  y=298, w=435, h=195  |  Card 4: x=489, y=298, w=435, h=195
Gap: 18pt horizontal, 18pt vertical
```

**CRITICAL: 4-KPI Dashboard Structural Requirements:**
- **Balanced Peer Zones**: All 4 cards MUST be identical in dimensions (w=435, h=195) and visual treatment — no dominant or secondary cards allowed
- **Consistent Template**: Every card MUST use the same exact structure and field ordering for dashboard coherence:
  1. **KPI Label** (header strip: 11pt ALL CAPS, white text on colored background)
  2. **Primary Value** (big number: 48-52pt bold, center-aligned)
  3. **Delta Indicator** (trend: "▲ +12%" directly below number, same alignment, green/red)
  4. **Target Reference** (comparison: "vs Target: $125M", 10pt grey, bottom of card)
- **Label-Number-Delta Ordering**: NEVER split metrics like "Net Promoter" + "Score: 72" — keep label compact ("NPS") with value immediately below to prevent text detachment
- **Equal Completeness Check**: Before finalizing, verify ALL 4 cards have identical field count and information depth — no card should have more or fewer data points than others
- **Readable Prominence**: All big numbers must be same font size (48-52pt) and same visual weight — no hierarchy allowed within the 4-card set

**3-across layout (3 KPIs):**
```
Card 1: x=36,  y=100, w=276, h=160
Card 2: x=332, y=100, w=276, h=160
Card 3: x=628, y=100, w=276, h=160
```

**CRITICAL: 3-KPI Dashboard Container Rules:**
- MANDATORY: Create three separate shape containers (rectangles with rounded corners and fill color)
- Each container must be visually distinct peer-level containers, not loose text clusters
- Gap between cards must be exactly 20pt horizontal (never overlap or touch)
- Each container MUST contain all 4 required elements:
  1. Metric name (header strip with colored background)
  2. Current value (big number, center-aligned)
  3. Comparator/delta (trend arrow + percentage or vs target)
  4. Status indicator (● symbol with status text)
- Anti-overlap validation: ALL THREE cards must remain fully visible (no clipping)
- Middle card (Card 2) positioning is critical: x=332pt ensures 20pt gaps from both sides

**5-card layout (3+2 staggered grid):**
```
Top row (3 cards):
Card 1: x=36,  y=85,  w=276, h=140
Card 2: x=332, y=85,  w=276, h=140  
Card 3: x=628, y=85,  w=276, h=140

Bottom row (2 cards, centered):
Card 4: x=184, y=245, w=276, h=140
Card 5: x=480, y=245, w=276, h=140
```

**CRITICAL: 5-KPI Multi-Card Dashboard Consistency Rules:**
- **Container Uniformity**: All 5 cards MUST use identical dimensions (w=276, h=140) and visual treatment
- **Grid Layout**: Default to balanced 3+2 grid (3 cards top row, 2 cards bottom row, centered)
- **Standard Evidence Template**: Every card MUST include these 6 fields in this exact order:
  1. **KPI Name** (header strip: white text on colored background, 11pt ALL CAPS)
  2. **Actual Value** (big number: 36-40pt bold, center-aligned, primary/semantic color)
  3. **Plan/Target** (comparison: "Target: $125M", left-aligned, 10pt grey)
  4. **Delta vs Plan** (performance: "▲ +2.4% vs target", right-aligned, 10pt green/red)
  5. **Status** (indicator: "● On Track", center-aligned, 10pt with semantic color)
  6. **Driver Note** (optional: brief explanation, 9pt grey, bottom of card)
- **Fixed Order Requirement**: Target (field 3), delta (field 4), and status (field 5) MUST appear on EVERY card in this consistent sequence for dashboard coherence
- **Layout Discipline**: Use 3+2 staggered grid, never 5-across (insufficient width for readable cards)
- **Balanced 3+2 Recommendation**: For 5-card dashboards, the 3+2 layout with identical card templates provides optimal visual hierarchy and readability balance — ALWAYS default to this arrangement unless specifically requested otherwise
- **Gap Consistency**: 20pt horizontal gaps, 20pt vertical gap between rows
- **Color Harmony**: Primary color for normal metrics, semantic colors (green/red) only for deltas and status
- **Negative Metric Interpretation**: Consistently interpret metric direction based on business impact:
  - **Favorable reductions**: Cost down 18%, Defects -0.3pp, Waste -12% → Display as GREEN with downward trend (▼ -18%) 
  - **Concerning deteriorations**: Revenue down 5%, Satisfaction -0.8pts → Display as RED with downward trend (▼ -5%)
  - **Rule**: If "down" or negative change improves business performance → GREEN; if it harms performance → RED

**MANDATORY Product Dashboard 5-Card Template (for "kpi-product" archetype):**

When the prompt mentions product metrics (MAU, DAU, conversion, retention, feature adoption, product velocity), ALL 5 cards MUST follow this exact template structure:

```
Card Structure (FIXED ORDER — never vary):
1. **Label** (header strip: metric name in ALL CAPS, white text on colored fill, 11pt bold)
2. **Current Value** (big number: actual metric value, 36-40pt bold, center-aligned, primary color)  
3. **Target** (comparison: "Target: [value]" or "Goal: [value]", 10pt grey, left-aligned below number)
4. **Delta/Trend** (performance: "▲ +12% vs target" or "▼ -3% MoM", 10pt, right-aligned, green/red semantic)
5. **Status** (indicator: "● On Track" / "● At Risk" / "● Exceeds Goal", 10pt, center-aligned, semantic color)
6. **Driver Note** (context: brief explanation of performance, 9pt grey, bottom of card)
```

**CRITICAL Product Dashboard Rules:**
- **Label Inside Tile**: The metric label MUST appear inside the tile header strip, never external to the card
- **Multi-System Attribution**: When KPIs come from different domains (analytics, engineering, quality, reliability), the source bar MUST specify: "Source: web analytics, application telemetry, issue tracker, incident log — Feb 2026 data"
- **Fixed Field Order**: Never omit Target (field 3), Delta (field 4), or Status (field 5) — all product cards require these for comparison coherence
- **Product-Specific Metrics**: Use domain-appropriate targets (MAU: millions, Conversion: percentages, Latency: milliseconds)

**Enriching KPI cards beyond basic (for 18+/20 quality):**

Basic KPI cards show: label, number, status. Elite cards also show:
- **Target comparison**: "Target: $125M" left-aligned below the number
- **Trend delta**: "▲ +2.4% YoY" right-aligned on the same row as target
- **Prior period**: small grey text "(Q3: $121M)" beneath target
- **Mini context row**: at the bottom of the card, small grey text explaining the driver

A truly elite KPI card has 5 information layers:
```
[HEADER STRIP: "REVENUE" — white on primary]
[BIG NUMBER: "$128M" — center-aligned, primary color]
[TARGET ROW: "Target: $125M" left | "▲ +2.4% YoY" right, green]
[STATUS: "● On Track" — center, green]
[DRIVER: "Enterprise segment drove 80% of growth" — 9pt grey, bottom]
```

**Standard People Dashboard Card Schema (MANDATORY for People/HR KPIs):**

For people-focused dashboards (headcount, retention, engagement, diversity, productivity), ALL cards MUST include these 5 fields in this exact order:
```
1. **Metric Name**: Header strip with white text on colored background (11pt ALL CAPS)
   Examples: "HEADCOUNT", "RETENTION RATE", "ENGAGEMENT SCORE", "TIME TO HIRE"

2. **Current Value**: Big number (48-52pt bold, center-aligned, primary color)
   Examples: "1,247", "94%", "4.2/5", "28 days"

3. **Target/Plan**: Comparison baseline (10pt grey, left-aligned)
   Examples: "Target: 1,300", "Plan: 90%", "Goal: 4.0/5", "SLA: 30 days"

4. **Trend/Delta**: Performance vs target (10pt, right-aligned, green/red semantic color)
   Examples: "▼ -4% vs target", "▲ +4pp YoY", "→ Flat vs Q3", "▲ 2 days faster"

5. **Status**: Performance indicator (10pt, center-aligned, semantic color)
   Examples: "● At Risk", "● Exceeds Target", "● On Track", "● Below Plan"
```

**People Dashboard Source Requirements:**
Source citations for people dashboards MUST include both system name and "as of" timestamp:
- Format: "Source: [System Name], data as of [Date/Time]"
- Examples: "Source: HRIS platform, data as of March 1, 2026" or "Source: engagement survey platform, data as of Q4 2025 survey close"
- Never use generic sources like "HR data" or quarterly labels without system identification

**Common KPI metric formatting patterns:**

For **percentage metrics** (retention, satisfaction, completion):
```
[HEADER STRIP: "CUSTOMER RETENTION" — white on primary]
[BIG NUMBER: "94%" — center-aligned, primary color, no decimal places]
[TARGET ROW: "Target: 90%" left | "▲ +4pp YoY" right, green]
[STATUS: "● Exceeds Target" — center, green]
```

For **rating/score metrics** (CSAT, NPS, quality):
```  
[HEADER STRIP: "CSAT SCORE" — white on primary]
[BIG NUMBER: "4.6/5" — center-aligned, primary color, one decimal]
[TARGET ROW: "Target: 4.0/5" left | "▲ +0.3 vs Q3" right, green]
[STATUS: "● Exceeds Target" — center, green]
```

For **count/volume metrics** (users, tickets, events):
```
[HEADER STRIP: "ACTIVE USERS" — white on primary]
[BIG NUMBER: "2.3K" — center-aligned, primary color, K/M abbreviation]
[TARGET ROW: "Target: 2.0K" left | "▲ +15% MoM" right, green]
[STATUS: "● On Track" — center, green]
```