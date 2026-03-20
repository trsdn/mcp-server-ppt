# Icon Shapes

Common icons built from PowerPoint native shapes. Use these instead of external images for consistent, scalable visuals.

## Why Native Shapes?

- No external files needed
- Scale without pixelation
- Match presentation colors automatically
- Work in any template

## Building Icons

Use `shape(action: 'create')` with these shape types and compositions.

### Indicator Icons (Single Shape)

| Icon | Shape Type | Size | Fill | Notes |
|---|---|---|---|---|
| Checkmark | msoShapeCheckmark (custom) | 24×24pt | Positive color | Or use ✓ in a text box |
| Cross/X | msoShapeCross | 24×24pt | Negative color | |
| Circle dot | msoShapeOval | 24×24pt | Accent, 2pt border | Filled for active, hollow for inactive |
| Arrow up | msoShapeUpArrow | 20×28pt | Positive color | Trend indicator |
| Arrow down | msoShapeDownArrow | 20×28pt | Negative color | Trend indicator |
| Arrow right | msoShapeRightArrow | 28×20pt | Primary color | Flow/process |
| Star | msoShape5pointStar | 28×28pt | Accent color | Highlight/featured |
| Lightning | msoShapeLightningBolt | 24×30pt | Accent color | Energy/speed |

### Status Indicators

Build these as small shape + text combinations:

**Traffic Light (RAG Status):**
```
Green circle (16pt): On track
Amber circle (16pt): At risk  
Red circle (16pt): Off track

Shape: msoShapeOval, w=16, h=16
Colors: #2E8B57, #F39C12, #C62828
Place inline with text or in table cells
```

**Progress Bar:**
```
Background: rounded rectangle, w=120, h=12, fill=#E8ECF1
Foreground: rounded rectangle, w=percentage*1.2, h=12, fill=Primary
Stack foreground on top of background
```

**Score Badge:**
```
Circle: msoShapeOval, w=48, h=48, fill=Primary
Text inside: score value, 18pt bold, white
Label below: 10pt, text medium
```

### Category Icons (Shape Compositions)

These use 2-3 shapes grouped together:

**Person/User:**
```
Head: circle, w=20, h=20, y=0
Body: trapezoid or rounded rect, w=28, h=24, y=22
Both: fill=Primary, no border
Group together
Total size: ~28×46pt
```

**Team/Group (3 people):**
```
Three person icons offset:
  Left:   x=0,  scale=85%
  Center: x=16, scale=100% (in front)
  Right:  x=32, scale=85%
```

**Document:**
```
Rectangle: w=28, h=36, fill=white, border=Primary 1.5pt
Corner fold: right triangle, w=8, h=8, top-right, fill=#E8ECF1
Lines inside: 3 thin rectangles (w=18, h=2) representing text
```

**Chart/Graph:**
```
3 vertical bars of different heights:
  Bar 1: x=0,  w=8, h=20, y=16
  Bar 2: x=10, w=8, h=28, y=8
  Bar 3: x=20, w=8, h=36, y=0
All fill=Primary
```

**Gear/Settings:**
```
Use msoShapeGear6 or msoShapeGear9
Size: 32×32pt
Fill: Primary color
```

**Target/Bullseye:**
```
Outer circle: w=36, h=36, fill=none, border=Primary 2pt
Middle circle: w=24, h=24, centered, fill=none, border=Primary 2pt
Inner circle: w=12, h=12, centered, fill=Accent
```

**Clock/Time:**
```
Circle: w=32, h=32, fill=white, border=Primary 2pt
Hour hand: thin rectangle rotated, from center
Minute hand: thin rectangle rotated, from center, longer
```

**Money/Dollar:**
```
Circle: w=32, h=32, fill=Positive
Text "$" inside: 20pt bold, white
```

**Globe/World:**
```
Circle: w=32, h=32, fill=Primary
Two curved lines (arcs) for meridians
One horizontal line for equator
All lines: white, 1.5pt
```

### Decorative Elements

**Divider Line:**
```
Thin line: w=full content width, h=1pt, fill=Neutral
Use between sections on a slide
y-position: midpoint between sections
```

**Accent Bar:**
```
Vertical bar: w=4pt, h=content height, fill=Primary
Place at left edge of a content block for emphasis
x-position: 28pt (8pt before content margin)
```

**Quote Marks:**
```
Large opening quote: text box with " character
Font: Georgia or serif, 72pt, Primary at 30% opacity
Position: x=36, y=120 (upper-left of quote text)
```

**Callout Box:**
```
Rounded rectangle: radius=8pt, fill=Neutral, no border
Or: fill=Primary at 8% opacity for subtle accent
Padding: 16pt inside
Use for highlighting key insights next to charts
```

**Number Badge:**
```
Circle: w=36, h=36, fill=Primary
Number inside: 18pt bold, white
Use for numbered steps in process diagrams
```

## Arrows and Connectors

CRITICAL: Arrows are the most visible structural element on process, flow, and comparison slides. Poor arrows ruin otherwise good slides.

### Arrow Type Selection

| Use case | Best approach | Shape/method |
|---|---|---|
| Trend indicator (up/down) | Block arrow shape | msoShapeUpArrow / msoShapeDownArrow, 16-24pt |
| Flow between process steps | Connector line | add-connector (straight or elbow) + set-line |
| Directional emphasis (left to right) | Chevron or block arrow | msoShapeChevron or msoShapeRightArrow |
| Timeline backbone | Long thin rectangle | Rectangle, h=4-6pt, w=full span, accent fill |
| Transformation bridge (A to B) | Right arrow shape | msoShapeRightArrow or msoShapeNotchedRightArrow, w=40-60pt |
| Text-inline direction | Unicode character | Use right arrow character in text box |
| Delta/change indicator | Small triangle | msoShapeIsoscelesTriangle, 12-16pt, rotated for direction |

### Professional Arrow Styling Rules

**Size and proportion:**
- Block arrows for flow: w=36-60pt, h=20-28pt (wider than tall)
- Trend indicator arrows: w=16-20pt, h=20-28pt (taller than wide)
- Process connector thickness: set-line weight 1.5-2.5pt (never >3pt for connectors)
- Timeline backbone: rectangle h=4-6pt (thin bar, not thick block)

**Color rules:**
- Flow arrows: Primary color at 60-80% opacity, or Neutral (#888888)
- Positive trend: Positive color (green family)
- Negative trend: Negative color (red family)
- Neutral/informational: Primary or medium grey
- NEVER use more than 2 arrow colors on one slide
- Arrow color should be LESS prominent than the content it connects

**Anti-patterns (common builder mistakes):**
- Block arrows too large (>60pt) dominating the slide content
- Thick connector lines (>3pt) that look heavy and unprofessional
- Mismatched arrow sizes across the same diagram
- Arrows overlapping text or data labels
- Using block arrows where thin connectors would suffice
- Rainbow-colored arrows (use ONE color for all flow arrows)
- Arrows without clear start/end alignment to the shapes they connect

### Flow Diagram Arrow Patterns

**Linear Process (3-5 steps):**
```
Step boxes: equal-width rectangles with rounded corners (8pt radius)
Between steps: small right-arrow shapes (w=28, h=16), centered vertically
Arrow fill: Primary at 50% opacity or Neutral
Spacing: 12-16pt gap between box edge and arrow edge
Alignment: all arrows at same y-center as boxes
```

**Chevron Chain (alternative to arrows):**
```
Use msoShapeChevron shapes touching edge-to-edge
Each chevron: w=equal portion of available width, h=40-48pt
Fill: gradient from light (first) to accent (last), or all same color
Text inside each chevron: bold label, 12-14pt, white on dark fill
Better than separate boxes + arrows for 4-6 sequential steps
```

**Vertical Flow:**
```
Step boxes stacked vertically
Between steps: msoShapeDownArrow, w=16, h=20, centered horizontally
Or: thin vertical line connector (1.5pt) with no arrowhead needed
Spacing: 8-12pt between box and arrow
```

**Hub-and-Spoke:**
```
Central circle (hub): 80-100pt diameter, accent fill
Surrounding boxes: arranged in a ring
Connectors: straight connectors from hub edge to each box
Line weight: 1.5pt, primary color
All connectors same length and weight for visual balance
```

### Comparison/Bridge Arrows

**Current vs Future State bridge:**
```
Center arrow: msoShapeRightArrow or msoShapeNotchedRightArrow
Size: w=48-60pt, h=28-36pt
Fill: Primary accent color
Label on/below arrow: transformation name in 10-12pt
Position: vertically centered between the two panels
```

**Before/After delta arrows (inline with metrics):**
```
Use small triangle shapes (msoShapeIsoscelesTriangle)
Size: 12-16pt (both w and h)
Positive (improvement): rotated to point UP, positive color fill
Negative (decline): rotated to point DOWN, negative color fill
Place immediately after the metric value text
Pair with percentage text: "triangle + 8.5%" in matching color
```

### Timeline Arrow Construction

**Do NOT use a single arrow shape as timeline backbone.** Use a thin rectangle instead:
```
Backbone: rectangle, h=4-6pt, w=full timeline span, fill=Primary or Neutral
Phase markers: small circles (16-20pt) or vertical bars on top of backbone
End cap (optional): small right-arrow shape at the right end only
This produces a clean, thin timeline — NOT a chunky block arrow
```

### Connector Best Practices

**When using add-connector:**
- Prefer straight connectors (type 1) when shapes are aligned
- Use elbow connectors (type 2) when shapes are offset vertically and horizontally
- Style with set-line: weight 1.5-2.0pt, color matching the diagram theme
- Keep all connectors in a diagram the SAME weight and color

**When connectors lack arrowheads (current limitation):**
- Place a small triangle shape (8-12pt) at the end of the connector to simulate an arrowhead
- Or use block arrow shapes instead of connectors for directional flows
- Or use chevron chains which inherently show direction

## Icon Sizing Guide

| Context | Icon Size | Spacing |
|---|---|---|
| Inline with body text (14-18pt) | 16-20pt | 8pt from text |
| Bullet marker | 20-24pt | 12pt from text |
| Card/tile header | 32-40pt | 16pt from title |
| Hero element | 48-64pt | 24pt from content |
| Section divider decoration | 64-96pt | Centered |

## Color Rules for Icons

- Use palette colors, never arbitrary colors
- Single-color icons: Primary or Accent
- Status icons: Positive/Negative colors
- Decorative/background: Neutral or Primary at 10-20% opacity
- Never use more than 2 colors in one icon
- Icons on dark backgrounds: use white or Text Inverse
