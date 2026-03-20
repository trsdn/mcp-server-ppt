# 1. Big Number

**When:** One headline metric dominates the message.

**CRITICAL: Proof-Oriented vs Hero-Only Layouts:**

When prompt includes "prove the math," "show how," or "demonstrate ROI calculation," use **PROOF-ORIENTED layout** with explicit mathematical structure:
- Hero lockup: metric at top (smaller, 48pt to leave room)
- Proof zone: equation/bridge/mini-waterfall showing Investment → Components → Total Value → ROI formula
- Formula must be explicit: "ROI = (Total Value - Investment) ÷ Investment × 100%"
- **NEVER let decorative hero text overlap or crowd the proof content** — proof readability is priority

When prompt is simple metric display ("show ROI" without math proof), use **HERO-ONLY layout**:
- Hero lockup: large unified metric (60-80pt, centered)
- Supporting bullets: brief context drivers only
- No calculation detail required

**Layout (16:9, 13.33" x 7.5" = 960pt x 540pt):**
```
Title: top, full width
Hero lockup: unified metric + label, centered, 60-80pt bold, accent color  
Proof block: driver bundle below hero, subordinate to headline
Context: small text or sparkline below
```

**CRITICAL: Growth Trajectory vs Simple Growth Callouts:**

When prompts mention "trajectory," "trend," "growth path," or "progression," use **VISUAL growth proof** instead of text-only growth statements:

**Visual Growth Proof Required:**
- **Sparkline chart:** Mini line chart at x=360pt, y=320pt, w=240pt, h=60pt showing 6-12 month progression
- **Mini-timeline:** Before/after comparison showing "Jan 450 → Dec 620" with trend arrow
- **Growth proof tiles:** 3-period tiles showing "Q1: 450 → Q2: 520 → Q3: 620" progression

**Simple Growth Callouts (text only):**
- When prompt says "grew X%" or "increased by Y" without trajectory language
- Use text bullets: "Headcount grew 38% year-over-year" or "Added 170 employees in 2025"

**CRITICAL:** The word "trajectory" specifically demands temporal visual proof — never satisfy trajectory requests with static percentage increases. **Trajectory = directional evidence over time.**

**Layout Impact:** When using visual growth proof, position the hero number HIGHER (y=160pt) to accommodate the growth visual below. Ensure descriptor text does not overlap the evidence block — keep 24pt minimum gap between hero lockup and proof elements.

**CRITICAL: Unified Hero Lockup (NOT separated label/number):**
- Position: x=480pt, y=200pt (center)
- Format: "300% ROI" as ONE unified visual unit (60-80pt bold, accent color)
- NEVER separate into "300%" + "Return on Investment" — keep the label and number together
- **HIERARCHY RULE: The NUMBER must be the most prominent visual element — larger font size or brighter color than the label text**
- Example lockups: "$8M savings" (number emphasized), "94% retention" (94% emphasized), "$128M revenue" (dollar amount emphasized)
- **Visual emphasis pattern: Make the numeric portion 20-30% larger or use accent color while keeping descriptive text in primary color**

**Proof Block (visual evidence bundle):**
- Position: x=240pt, y=280pt, w=480pt (centered below hero)
- Contains: 2-3 proof drivers, visually bundled as one support unit
- Font: 14pt regular, medium grey (#666666) — subordinate to hero
- Format each driver: "Platform savings: $3.1M" + "Revenue uplift: $1.8M" + "License consolidation: $0.4M"
- Gap between drivers: 8pt line spacing
- Total math: small grey text "(= $5.3M value on $1.8M investment)"

**CRITICAL: When prompt asks for "showing each source" or "proving the math":**
When user requests component breakdown or mathematical proof, **DEFAULT to waterfall/bridge chart rather than text list**:
- Use waterfall chart showing: Investment → Component 1 → Component 2 → Component 3 → Total Value
- Each component shows additive logic visually (bars building toward total)
- Position chart at x=240pt, y=220pt, w=480pt, h=200pt (replaces text proof block)
- Label each waterfall segment with specific value and arrow showing flow
- **NEVER use bullet-point text when user asks to "show the math" — make the arithmetic visual**

**CRITICAL: Mandatory Chart Evidence Rules:**

When prompts specify exact chart forms like "benchmark bar," "trend line," "comparison chart," or "time-series," you MUST instantiate those exact evidence structures as actual charts — NEVER convert specified chart forms into text descriptions or bullet points:

- **"benchmark bar"** → Create horizontal bar chart comparing your metric vs benchmark at x=240pt, y=220pt, w=480pt, h=180pt
- **"trend line"** → Create line chart showing time progression at x=240pt, y=220pt, w=480pt, h=180pt  
- **"comparison chart"** → Create clustered column chart showing side-by-side values at x=240pt, y=220pt, w=480pt, h=180pt

**Big Number + Dual Proof Archetype:**

For KPI slides requiring BOTH external comparison AND time-series validation (e.g., "Our NPS of 72 beats industry benchmark and shows consistent growth"), use dual evidence layout:

**Layout structure:**
- Hero number: x=480pt, y=140pt (48pt bold, centered, raised position)
- **Left proof zone:** Benchmark comparison chart at x=120pt, y=220pt, w=200pt, h=160pt  
  - Horizontal bar chart: Your value vs industry benchmark
  - Labels: "Us: 72" and "Industry: 58" with delta callout "+14"
- **Right proof zone:** Trend line chart at x=600pt, y=220pt, w=200pt, h=160pt
  - Line chart showing 6-12 month progression  
  - Y-axis labeled with NPS scale, X-axis with time periods
- **Gap between charts:** 80pt minimum for visual separation
- Context line below: "Sustained outperformance across both benchmarking and trending dimensions" at y=400pt

**NEVER substitute either chart with text bullets** — both visual proof elements are mandatory when dual validation is requested.

**ALTERNATIVE for quarterly/period build-up:** When showing Q1-Q4 contribution or period-by-period build-up, **additive quarter tiles** are equally valid when the sum is obvious:
- Use 4 mini KPI cards showing: Q1 $32M, Q2 $31M, Q3 $33M, Q4 $32M → Total: $128M 
- Position cards at y=220pt, equal width distribution, with clear total indicator
- **Strong evidence pattern:** Period shares + quarter values provide strong proof for annual-total claims
- Only use waterfall if the cumulative flow needs emphasis; tiles work when quarters clearly sum to annual

**Example title:** "Customer retention reached 94%, highest in 5 years"

**CRITICAL: Benchmark Big Numbers (When Comparing to External Data):**

When the big number compares internal metrics against external benchmarks (industry average, competitor data, analyst reports), use a **visual comparison lockup** instead of text-only support:

**Visual Pattern 1 — Side-by-side comparison:**
- Position: x=480pt, y=180pt (centered)
- Format: "94% vs 82%" as unified lockup (60pt bold, accent color vs medium grey)
- Labels below: "Our retention" | "Industry avg" (12pt grey, positioned under each number)
- Emphasis: YOUR number larger/bolder than benchmark number

**Visual Pattern 2 — Delta callout:**
- Hero number: "94% retention" (60pt bold, accent color, centered)
- Delta badge: "+12pp" positioned at x=520pt, y=170pt (18pt bold, positive color, circle background)
- Context line: "vs industry benchmark (82%)" (14pt primary color, below hero, visually connected to delta badge via alignment or subtle connecting line)
- **CRITICAL: The benchmark proof line must remain clearly legible (14pt minimum) and visually tied to the delta callout — NEVER fade it to light grey**

**CRITICAL: Delta-as-hero pattern (when the gap IS the message):**
When the benchmark gap is central to your argument (even when already stated in the title), ALWAYS show the delta as a dedicated visual callout or annotation. This applies when:
- The size of the gap is the key insight ("We're beating industry by 12pp")
- The comparison gap drives a decision ("This 15% lead justifies premium pricing")
- The delta trend is changing ("Gap narrowed from 20pp to 8pp")

**Implementation:** Use Visual Pattern 2 (Delta callout) but make the delta badge the primary visual element — position it larger (24pt bold) and more prominent than the base number itself.

**CRITICAL: Support Bullet Content for Benchmark Comparisons:**

When prompts include "vs benchmark," "vs industry average," or "compared to competitors," support bullets must either:

**Option 1 — Segment performance differences:**
- "Enterprise clients: 97% retention vs 85% industry avg"
- "SMB segment: 89% retention vs 79% industry avg"  
- "New customer cohort: 91% vs 83% benchmark"

**Option 2 — Drivers of outperformance:**
- "Proactive support reduced churn by 8pp"
- "Product feature adoption 40% higher than industry"
- "Customer success program impact: +5pp retention boost"

**NEVER use generic context bullets** like "Strong customer satisfaction" or "Improved processes" for benchmark slides — the bullets must explicitly connect your outperformance to either segmented breakdowns or causal drivers that explain WHY you beat the benchmark.

**SOURCE CITATION REQUIREMENTS (Two-source rule):**
When presenting benchmark comparisons, ALWAYS cite BOTH sources separately:
- Internal source: "Source: Customer Success platform, Q4 2025 cohort analysis"
- Benchmark source: "Industry benchmark: Customer Experience Benchmark Report 2025 (n=1,200 SaaS companies)"
- Position both citations in source bar (use semicolon separator)

**CRITICAL: Data source specificity requirements:**
For ALL data slides (not just benchmarks), source citations must include:
- **System name**: Name the actual underlying system/database ("CRM platform", "analytics warehouse", "ERP system")
- **Timestamp specificity**: Include exact data pull date and time period ("Data as of Jan 15, 2025, 8am EST", "Q4 2025 actuals")  
- **Sample scope**: Specify what data is included ("All active accounts", "North America only", "Enterprise tier customers")

**Example enhanced source format:**
"Source: CRM retention analysis, all enterprise accounts, Q4 2025 actuals, data as of Jan 15, 2025"

**Example source bar format:**
"Source: CustomerHub retention analysis, Q4 2025; Industry benchmark: Gainsight SaaS Metrics Report 2025"

**CRITICAL: Competitive Ranking Requirements:**

When prompts mention "overtook competitor," "ahead of #2 player," "beat [Company X]," or similar competitive positioning language, you MUST use **ranking or side-by-side competitor comparison archetype**, NOT just text under a hero number.

**Required evidence patterns for competitive claims:**
- **Market position chart**: Horizontal bar chart showing your company vs named competitors at specific positions (e.g., "1. Us: 72%", "2. Competitor A: 68%", "3. Competitor B: 64%")
- **Ranking table**: List showing explicit rank positions with company names and values in descending order
- **Head-to-head comparison**: Side-by-side lockup format showing "Us: 72% vs [Competitor]: 68%" with clear visual emphasis on your advantage

**For dual-evidence prompts (gain-over-time + competitive overtake):**
When prompts combine both trajectory language ("growing," "trajectory," "trend") AND competitive positioning ("overtook," "ahead of"), you must show BOTH evidence types explicitly:
- **Left proof zone**: Ranking/competitive comparison chart showing current market position
- **Right proof zone**: Time-series chart showing YOUR trajectory that led to overtaking the competitor
- **NEVER satisfy both requirements with just growth trajectory** — the competitive ranking must be visually explicit

**Descriptor text sizing warning:**
Avoid oversized descriptor text that obscures the evidence layer. Keep competitor names and context text to 12-14pt maximum to ensure ranking charts and competitive comparison elements remain clearly legible. The competitive evidence is the primary proof — descriptor text is secondary.

**Big Number + Pathway Breakdown Archetype (for Sustainability Claims):**

For sustainability metrics requiring component attribution (e.g., "40% carbon reduction via pathway optimization"), use **stacked additive proof** to show how individual contributions build toward the total impact:

**Layout structure:**
- Hero number: x=480pt, y=140pt (48pt bold, centered, raised position for evidence clearance)
- **Pathway proof zone:** Stacked component chart at x=240pt, y=220pt, w=480pt, h=180pt
  - Horizontal stacked bar showing: Component 1 (12%) + Component 2 (15%) + Component 3 (13%) = Total (40%)
  - Each segment labeled with both percentage contribution and method: "Supply chain optimization: 12%", "Renewable energy: 15%", "Process efficiency: 13%"
  - Colors: Sequential green palette (#E8F5E8, #A8D8A8, #4CAF50) showing progression
- **NEVER use bullet points for pathway breakdowns** — component contributions must be visually additive to demonstrate cumulative impact
- **Descriptor sizing rule:** Pathway method labels maximum 12pt to prevent overlap with evidence bars — the stacked proof structure is primary, descriptive text is secondary

**Required for sustainability prompts mentioning:**
- "pathway," "breakdown," "components," "via," "through," or "consists of"
- Specific method attribution: "optimization," "efficiency," "renewable," "circular economy"
- Percentage or fractional contributions that sum to the hero metric

**Example implementation:**
Hero: "40% carbon reduction" (centered, 48pt bold, green accent)
Stacked bar showing three additive segments totaling 40%, with method labels positioned below each segment at 12pt maximum to maintain evidence clarity.