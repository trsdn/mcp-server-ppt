# 7. Comparison Slide

**When:** Side-by-side evaluation of 2-3 options.

**Layout (2-column):**
```
Title: top
Column A: left half (x=36pt, y=100pt, w=420pt, h=380pt)
Column B: right half (x=480pt, y=100pt, w=420pt, h=380pt)
  - Each column: colored header strip + comparison rows
  - Use color coding: red header for "current/before", green for "future/after"
  - **SLA/Target comparison best practice:** Explicitly label columns as "Current" and "Target" when comparing against performance targets, even when color coding makes the distinction clear
Divider: thin vertical line at x=458pt (between columns)
```

**STRUCTURED CONTENT BLOCK (prevents floating elements):**

Each column must form a cohesive comparison block with exact positioning:

```
Column Header (per column):
- Position: x=column_x, y=100pt, w=420pt, h=40pt
- Background: accent color rectangle (red/green coding)
- Text: column title (18pt bold white, centered vertically+horizontally)
- CRITICAL: Headers MUST NOT compete with metrics - use muted accent colors (30-40% opacity) to maintain visual hierarchy

Comparison Rows (stacked within each column):
- Row 1: y=140pt, h=70pt (feature + metric + description)
- Row 2: y=210pt, h=70pt  
- Row 3: y=280pt, h=70pt
- Row 4: y=350pt, h=70pt
- Each row spans full column width (w=420pt)

Within each 70pt row:
- Feature name: y=row_y+10pt, 14pt semibold
- Metric value: y=row_y+30pt, 24pt bold accent color  
- Description: y=row_y+50pt, 12pt grey
```

This structured approach prevents content from floating and ensures visual alignment across columns.

**CRITICAL: Make metrics the visual hero, not buried in body text.**

Each comparison row should lead with the metric change as the dominant element:

```
Weak: "• Manual reporting takes 3 weeks to close books"
Strong: "3 weeks → 2 days" (28pt bold, red→green) 
        "Book close cycle" (12pt grey, below)
```

Use arrows (→) or before/after pairs with the METRIC large and the description small.
This makes the slide scannable — executives see the numbers instantly.

**CONSISTENT FEATURE STRUCTURE (PREVENTS TRUNCATION):**

When comparing options with multiple features, use identical row structure across all columns to ensure nothing gets cut off:

```
Feature Row Template (per column, 40pt height minimum):
Line 1: Feature category (14pt semibold, consistent length)
Line 2: Specific capability (12pt, concise but complete)
Line 3: Metric/value (16pt bold, highlight color)
```

**Anti-truncation Rules:**
- **Match row count**: All columns must have same number of features (pad with "N/A" if needed)
- **Consistent labels**: Use identical feature names across columns ("Cost" not "Price" vs "Cost")  
- **Complete descriptions**: Never abbreviate feature details - use full capability names
- **Standard spacing**: 40pt minimum row height prevents text cutoff in any column
- **Aligned categories**: Feature categories appear in same sequence across all columns

**OPERATING-MODEL COMPARISON SELECTION (CRITICAL):**

For before/after operating-model slides, choose the layout complexity based on what drives the metric change:

**CLEAN MIRRORED METRICS (preferred when):**
- Metrics change due to direct system improvements (automation, tool upgrades)
- The "how" is obvious from the metric itself ("Manual → Automated")
- Process steps remain the same, just faster/cheaper/better
- Example: "3 weeks → 2 days book close" (tool upgrade, same process)

**ADD PROCESS-FLOW VISUALS when:**
- Operating model fundamentally changes (new workflow, restructured teams)
- Multiple interconnected process changes drive the metrics
- The "how" requires explanation of new steps or sequence
- Example: "Centralized shared services" model (new org structure, new handoffs)

**ADD DRIVER ANNOTATIONS when:**
- Single metric improvement has multiple contributing factors
- Need to prove metric change is achievable (show the levers)
- Stakeholders will ask "what specifically drives this?"
- Example: "40% cost reduction" with callouts: "20% automation + 15% consolidation + 5% renegotiation"

**NUMERIC FIDELITY (CRITICAL):** Every number, percentage, and unit displayed on comparison slides MUST exactly match the values provided in the prompt. Verify that:
- Headline claims in titles reconcile exactly to the displayed evidence
- All metrics preserve their original values and units (never round or approximate)  
- Before/after comparisons mirror the specific values given
- Calculate percentage changes precisely: if prompt says "12.5% to 18.7%", show exactly "12.5% → 18.7%" not "~13% to ~19%"
- Preserve currency symbols, decimal places, and measurement units exactly as specified
Any discrepancy between title claims and visual evidence destroys credibility.

**SCENARIO COMPARISON FORMATTING (CRITICAL for Financial Models):**

When comparing business scenarios (best/worst case, bull/bear, optimistic/conservative), apply strict numeric formatting to prevent rendering artifacts:

**Clean Number Presentation Rules:**
- Numbers with decimals: Use exactly 1 decimal place for percentages, 0 decimals for millions
- Currency: "$24.5M" not "$24,500,000" or "$24.5 million"
- Percentages: "12.5%" not "12.50%" or "~13%"  
- Ratios: "2.3x" not "2.3:1" or "2.30x"
- Remove stray punctuation: "EBITDA: $18.2M" not "EBITDA: $18.2M," or "(EBITDA): $18.2M"

**Driver Linkage Disclosure:**
When scenario spreads result from specific assumption changes, surface the key driver prominently:
- **Title enhancement**: "Revenue scenarios driven by volume assumption: +15% vs -10%"  
- **Subtitle callout**: Add driver note below title at x=40pt, y=80pt: "Volume assumption: Base 100k units, Bull 115k, Bear 90k" (12pt italic grey)
- **Column headers**: Include assumption delta: "Optimistic (+15% volume)" not just "Optimistic"

**Example Structure:**
```
Title: "Revenue Impact: Bull vs Bear Scenarios"
Subtitle: "Key driver: Volume assumption varies ±15% from 100k unit base"
Column 1 Header: "Bear Case (85k units)"  
Column 2 Header: "Base Case (100k units)"
Column 3 Header: "Bull Case (115k units)"
Metric rows: "$18.2M" → "$24.5M" → "$31.8M" (clean formatting, no extra punctuation)
```

**Layout (3-column):**
```
Column widths: 280pt each, 30pt gaps
x positions: 55pt, 365pt, 675pt
```

**DECISION MATRIX PATTERN (When comparing multiple vendors/options with criteria):**

When prompts involve vendor selection, solution comparison, or scoring against multiple criteria, use these visual markers to make the winner defensible:

**Winner Highlighting:**
- **Best-in-class cells**: Fill winning scores with green background (#E8F5E8) + bold text
- **Weighted criteria note**: Add "(Weight: 40%)" in small grey text below key criteria headers
- **Total score row**: Bold the final row, highlight the winner's total in accent green (#0D7C0D)

**Layout for Decision Matrix:**
```
Table: centered (x=80pt, y=130pt, w=800pt)
- Column 1: Criteria (200pt width, left-aligned)
- Columns 2-4: Vendor options (200pt each, center-aligned)
- Best scores: Green background (#E8F5E8) + bold text
- Winner column: Subtle green border (2pt) around entire column
- Total row: Bottom border + bold formatting
```

**BUILD VS BUY PATTERN (CRITICAL for Technology Decisions):**

When comparing internal development against vendor solutions, use this specialized 4-zone structure:

**Zone 1 — Cost Comparison (Top Left, x=50pt, y=110pt, w=400pt, h=120pt):**
- Total Cost of Ownership display: "Build: $2.4M vs Buy: $1.8M" (24pt bold)
- Breakdown bullets: Development, maintenance, licensing costs (12pt)
- Time horizon note: "3-year TCO analysis" (10pt grey)

**Zone 2 — Timeline Reality Check (Top Right, x=480pt, y=110pt, w=400pt, h=120pt):**
- Delivery timeline: "Build: 18 months vs Buy: 3 months" (24pt bold) 
- Risk callout: "Build includes 6-month integration risk buffer" (12pt orange text)
- Resource requirement: "Build requires 8 FTE developers" (12pt)

**Zone 3 — Capability Matrix (Bottom Left, x=50pt, y=260pt, w=400pt, h=140pt):**
- Feature checklist with visual indicators:
  - "Core features: Buy ✓ Build ✓" (14pt, green checkmarks)
  - "Customization: Buy ✗ Build ✓" (14pt, red X vs green checkmark)
  - "Vendor lock-in risk: Buy ⚠ Build ✓" (14pt, yellow warning vs green checkmark)

**Zone 4 — Strategic Recommendation (Bottom Right, x=480pt, y=260pt, w=400pt, h=140pt):**
- Recommendation badge: "RECOMMENDED: BUY" (18pt bold white text on green background, x=500pt, y=280pt)
- Key rationale: "60% cost savings + faster delivery outweighs customization loss" (12pt)
- Risk mitigation: "Negotiate API access to reduce lock-in" (12pt italic)
```

**Visual Hierarchy:**
1. **Immediate scan**: Winner column stands out with subtle green border
2. **Detail validation**: Best-in-class cells visually pop with green highlights  
3. **Context**: Weighted criteria notes justify the scoring methodology
4. **Conclusion**: Bold total row with clear winner emphasis

This pattern makes the selection rationale immediately defensible — stakeholders see both the winner and the evidence supporting that choice.

**QUANTITATIVE PROOF REQUIREMENT (CRITICAL for Comparison Credibility):**

Comparison slides MUST be evidence-driven, not qualitative assertions. Every comparison claim requires quantitative backing:

**MANDATORY PROOF PATTERNS for comparison metrics:**

```
Weak (qualitative): "Platform A is faster than Platform B"
Strong (quantitative): "Platform A: 2.3s response time vs Platform B: 8.7s response time"

Weak (qualitative): "Solution reduces costs significantly" 
Strong (quantitative): "$2.4M annual cost → $1.1M annual cost = 54% reduction"

Weak (qualitative): "Better user adoption rates"
Strong (quantitative): "87% daily active users vs competitor's 34% DAU"
```

**EVIDENCE REQUIREMENTS (before/after and competitive comparisons):**
- **Baseline numbers**: Current state with specific values, units, timeframes
- **Target numbers**: Future state with specific values, not percentages alone  
- **Delta calculations**: Exact change amounts ("saves $1.3M annually") with percentage ("54% reduction")
- **Proof methodology**: How metrics were measured, verified, or benchmarked
- **Context bounds**: Time periods, sample sizes, measurement conditions

**SOURCE ATTRIBUTION for competitive comparisons:**
- **Benchmark source**: "Based on Q3 2024 industry analyst evaluation" or "Internal pilot comparison, July 2024"
- **Measurement method**: "Response times measured under 1000 concurrent users" 
- **Verification**: "Third-party performance audit" or "Customer-reported metrics"
- **Scope boundaries**: "Enterprise implementations only" or "North American deployments"

**ANTI-PATTERNS TO AVOID:**
- ❌ "Much faster/cheaper/better" without numbers
- ❌ "Industry-leading" without benchmark proof
- ❌ Percentages without underlying values ("30% improvement" → improvement from what to what?)
- ❌ Generic claims ("seamless integration," "enhanced security") without measurable criteria
- ❌ Competitor comparisons without named competitors or specific metrics

**Every comparison row must answer: "Prove this claim with numbers a CFO could audit."**

**SUPPORTING EVIDENCE IN BODY CONTENT:**

Each metric must include contextual evidence that supports the quantitative claim:

```
Metric: "2.3s response time vs 8.7s response time" (24pt bold)
Supporting evidence: "Measured during peak trading hours (9-11 AM EST), 1000 concurrent users, 
3-month rolling average" (12pt grey, positioned below metric)

Metric: "$2.4M → $1.1M annual cost" (24pt bold)  
Supporting evidence: "Infrastructure savings ($800K) + operational efficiency ($500K), 
validated by Q3 2024 pilot program" (12pt grey)
```

CRITICAL: Evidence text must appear in EVERY comparison row, not just in headers or summary sections. This ensures each claim is immediately defensible when stakeholders focus on individual metrics.