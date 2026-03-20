# Archetype Registry

17 curated archetype families for slide design. Choose the right family first, then load the family-specific detail file for layout coordinates, variant rules, and anti-patterns.

The unified runtime catalog also includes learned reference families that are not authored as curated layout files here. For the full end-to-end pipeline from slide curation to sanitized runtime integration, see [Archetype Pipeline](../../../../docs/ARCHETYPE-PIPELINE.md).

## Decision Tree: Which Archetype?

```
What are you communicating?
├─ A single key metric? → big-number
├─ System reliability with target comparison + trend? → operational-kpi
├─ Multiple related metrics? → kpi-card-dashboard
├─ A trend over time? → column-bar-chart
├─ A comparison of options? → comparison
├─ A data insight with context? → chart-insight-callout
├─ A conceptual framework? → framework
├─ Detailed data/financials? → simple-table
├─ How something breaks down? → waterfall-chart
├─ A plan or timeline? → timeline-roadmap
├─ A process or workflow? → process-diagram
├─ Geographic data? → map
├─ Key findings summary? → executive-summary
├─ A voice/testimony? → quote
├─ What to do next? → recommendations
├─ Opening/closing with strategic framing? → title-slide
└─ Supporting detail? → appendix
```

## Family Index

| ID | Name | When | Variants | Best Density |
|----|------|------|----------|-------------|
| `big-number` | Big Number | One headline metric dominates | hero-only, proof-oriented, dual-proof, benchmark, pathway-breakdown | D1, D2 |
| `kpi-card-dashboard` | KPI Card Dashboard | 2-6 metrics shown together | 2x2, 3-across, 5-card, people, product | D2, D3 |
| `operational-kpi` | Operational-KPI | Target comparison + trend for reliability | target-comparison, incident-trend, dual-evidence | D2, D3 |
| `column-bar-chart` | Column/Bar Chart | Comparing values or showing trends | — | D2, D3 |
| `chart-insight-callout` | Chart + Insight Callout | Data viz with "so what" annotation | — | D2, D3 |
| `framework` | Framework | Organizing concepts, relationships | 2x2, venn, pillars, swot, maturity, raci, stakeholder, risk, canvas | D2, D3 |
| `simple-table` | Simple Table | Structured data, financials | — | D3, D4 |
| `waterfall-chart` | Waterfall Chart | Value build-up or breakdown | — | D2, D3 |
| `comparison` | Comparison Slide | Side-by-side evaluation | 2-column, 3-column, decision-matrix, build-buy, scenario, before-after | D2, D3 |
| `timeline-roadmap` | Timeline/Roadmap | Plans, milestones, phases | horizontal, swimlane, now-next-later, milestone-phases | D2, D3 |
| `process-diagram` | Process Diagram | Workflows, decision flows | linear, chevron, swimlane, funnel, stage-gate | D2, D3 |
| `executive-summary` | Executive Summary | Key findings, opening/closing | numbered-findings, incident-postmortem-cover | D1, D2 |
| `recommendations` | Recommendations | Closing with actions | structured-table | D2 |
| `quote` | Quote Slide | Human voice, testimony | quantified-testimonial, nps-before-after, saas-retention | D1, D2 |
| `map` | Map Slide | Geographic distribution | — | D2, D3 |
| `title-slide` | Title Slide | Opening with strategic framing | investor, cost-program, climate-esg, product-launch, integration-review | D1 |
| `appendix` | Appendix | Backup/reference material | — | D5 |

## Consulting Variant Map

Map named consulting patterns to archetype families:

| Catalog family | Representative variants | Default family |
|---|---|---|
| Roadmaps and timelines | Horizontal timeline, swimlane, Gantt, S-curve, chevron, now-next-later | timeline-roadmap |
| Frameworks and matrices | 2x2, 3x3, SWOT, Porter's, value chain, pyramid, Venn, risk matrix | framework |
| Grid-heavy governance | Maturity model, RACI, scoring matrix, heat-map table | simple-table or framework |
| Processes and flows | Chevron chain, swimlane, circular, funnel, stage-gate, decision tree | process-diagram |
| Comparisons | Side-by-side, before-after, gap analysis, pros-cons, TCO, scenario | comparison |
| Analytical charts | Waterfall, stacked bar, tornado, slope, dot plot, treemap, Sankey | column-bar-chart, waterfall-chart, or chart-insight-callout |
| Organizations | Org chart, stakeholder map, hub-and-spoke, ecosystem map | framework |
| Dashboards | KPI dashboard, RAG overview, balanced scorecard, traffic-light | kpi-card-dashboard or executive-summary |
| Architecture | Layered, system context, reference arch, data flow, cloud arch | framework or process-diagram |
| Strategy classics | Strategy house, BMC, Ansoff, OKR, three horizons, flywheel | framework, big-number, or comparison |
| Finance | P&L bridge, revenue breakdown, cash flow waterfall, break-even | waterfall-chart, comparison, or column-bar-chart |
| People and change | Team overview, competency radar, change impact, enablement | kpi-card-dashboard, framework, or timeline-roadmap |
| Outcomes | Key findings, impact showcase, quote, recommendations | executive-summary, big-number, quote, or recommendations |
| Consulting metaphors | Harvey balls, bubble chart, staircase, iceberg, honeycomb | comparison, framework, or executive-summary |
| Utility formats | Agenda, section divider, glossary, data table, calendar grid | appendix, timeline-roadmap, or map |
| Composite panels | Left chart + right insights, dashboard composite, icon grid | chart-insight-callout, kpi-card-dashboard, or executive-summary |
| Workshop | Voting results, affinity diagram, empathy map, action log | framework, executive-summary, or recommendations |

## Cross-Cutting Guidance

### Choosing Aspect Ratio

| Context | Ratio | Dimensions | Default? |
|---|---|---|---|
| Projected presentation | 16:9 | 960pt × 540pt | **Yes** |
| Printed handout | 4:3 | 720pt × 540pt | Board decks |

### Avoiding Layout Monotony

**Rule: Never use the same archetype on 3+ consecutive slides.** Alternate between column layouts, full-width layouts, and chart-based layouts.

| Content | Layout | Why Different |
|---|---|---|
| Strategic pillars | 3-col cards + vision bar | Hierarchy |
| Features | 3-col with icon circles | Function |
| Case studies | Full-width stacked rows | Narrative |
| Comparisons | Side-by-side with divider | Contrast |
| Timelines | Horizontal flow with milestones | Sequence |

### Headline Patterns by Archetype

| Archetype | Headline pattern |
|-----------|-----------------|
| Executive Summary | "We recommend [action] to achieve [outcome]" |
| Big Number | "[Subject] is [doing X], driven by [Y]" |
| Comparison | "[A] outperforms on [X], but [B] is stronger on [Y]" |
| Column/Bar | "Growth is reaccelerating after [event]" |
| Waterfall | "Margin decline is mostly [driver]-driven" |
| KPI Dashboard | "[X] of [Y] metrics are green — [area] needs attention" |
| Framework | "Prioritize [quadrant]: [X] high-impact, low-effort items" |
| Timeline | "A two-wave plan captures quick wins and de-risks change" |
| Process | "Current process breaks at [X] handoffs, adding [Y] days" |
| Recommendations | "Approve [action] now to unlock [value] by [date]" |

### Auto-Reject Conditions

Reject and regenerate a slide if ANY apply:

- Title is a topic label instead of an action title
- Title exceeds 20 words or contradicts visual evidence
- Multiple charts compete without comparative framing
- Chart type doesn't match data shape (pie for time series)
- More than 2 accent colors without semantic meaning
- No source or date on a data slide
- Waterfall bars don't mathematically sum
- "Other" is the largest category in a breakdown
- Roadmap has no sequencing logic
- Recommendation has no explicit ask
