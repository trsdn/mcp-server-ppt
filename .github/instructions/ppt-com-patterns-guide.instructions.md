---
applyTo: "src/PptMcp.Core/Commands/**/*.cs,tests/**/*.cs"
---

# PowerPoint COM Patterns - Quick Reference

> **Essential patterns for PowerPoint COM automation via late binding**

## Core Principles

1. **Use Late Binding** - `dynamic` types with `Type.GetTypeFromProgID("PowerPoint.Application")`
2. **1-Based Indexing** - All PowerPoint collections (Slides, Shapes, Paragraphs) start at 1
3. **Exception Propagation** - Never wrap in try-catch; let `batch.Execute()` handle exceptions
4. **msoTrue / msoFalse** - PowerPoint uses `MsoTriState`: `msoTrue = -1`, `msoFalse = 0`, `msoCTrue = 1`
5. **Points as Units** - Positions and sizes are in points (1 inch = 72 points)

---

## Decision Tree: Which Tool to Use

```
Working with presentations?
├─ Slide lifecycle (add, delete, duplicate, move, reorder)?
│  └─ Use slide tool
│
├─ Shapes on a slide (create, modify, delete, position)?
│  └─ Use shape tool
│
├─ Text inside a shape (read, write, format)?
│  └─ Use text / shape tool (TextFrame access)
│
├─ Charts embedded in a slide?
│  └─ Use chart tool
│
├─ Placeholders (title, body, footer)?
│  └─ Use placeholder tool
│
└─ Presentation-level properties (slide size, metadata)?
   └─ Use file / presentation tool
```

---

## Slide Operations

### Adding Slides

```csharp
// Get a slide layout from the first slide master
dynamic slideMaster = presentation.SlideMasters.Item(1);
dynamic customLayout = slideMaster.CustomLayouts.Item(layoutIndex); // 1-based

// Add slide at a specific position
dynamic newSlide = presentation.Slides.AddSlide(position, customLayout);
```

### Navigating Slides

```csharp
// By index (1-based)
dynamic slide = presentation.Slides.Item(slideIndex);

// By SlideID (stable across reorders)
dynamic slide = presentation.Slides.FindBySlideID(slideId);

// Iterate all slides
for (int i = 1; i <= presentation.Slides.Count; i++)
{
    dynamic slide = presentation.Slides.Item(i);
    // ... process slide ...
    ComUtilities.Release(ref slide!);
}
```

### Deleting and Reordering

```csharp
// Delete
slide.Delete();

// Move to new position
slide.MoveTo(newPosition); // 1-based target index

// Duplicate (returns SlideRange, not single slide)
dynamic slideRange = slide.Duplicate();
dynamic duplicated = slideRange.Item(1);
```

---

## Shape Lifecycle

### Creating Shapes

```csharp
// Basic shape (left, top, width, height in points)
dynamic shape = slide.Shapes.AddShape(
    1,    // msoShapeRectangle
    100f, // left
    100f, // top
    200f, // width
    150f  // height
);

// Text box
dynamic textBox = slide.Shapes.AddTextbox(
    1,    // msoTextOrientationHorizontal
    50f, 50f, 300f, 100f
);

// Picture
dynamic picture = slide.Shapes.AddPicture(
    filePath,
    0,    // msoFalse = don't link
    -1,   // msoTrue = save with document
    left, top, width, height
);

// Table
dynamic table = slide.Shapes.AddTable(
    numRows, numColumns,
    left, top, width, height
);
```

### Modifying Shapes

```csharp
// Position and size
shape.Left = 100f;
shape.Top = 50f;
shape.Width = 300f;
shape.Height = 200f;

// Rotation (degrees)
shape.Rotation = 45f;

// Name (for identification)
shape.Name = "MyShape";

// Visibility
shape.Visible = -1; // msoTrue
```

### Deleting Shapes

```csharp
shape.Delete();
// Release COM reference immediately after delete
ComUtilities.Release(ref shape!);
```

### Shape Type Detection

Use `shape.Type` (`MsoShapeType` enum values):

| Value | Constant | Description |
|-------|----------|-------------|
| 1 | msoAutoShape | Basic shapes |
| 6 | msoGroup | Grouped shapes |
| 13 | msoPicture | Images |
| 14 | msoPlaceholder | Placeholder shapes |
| 17 | msoTextBox | Text boxes |
| 19 | msoTable | Tables |
| 3 | msoChart | Charts |

```csharp
int shapeType = Convert.ToInt32(shape.Type);
if (shapeType == 14) // msoPlaceholder
{
    int phType = Convert.ToInt32(shape.PlaceholderFormat.Type);
}
```

---

## Text Manipulation

### TextFrame → TextRange → Font

```csharp
dynamic? textFrame = null;
dynamic? textRange = null;
dynamic? font = null;
try
{
    // Access text content
    textFrame = shape.TextFrame;
    textRange = textFrame.TextRange;

    // Read text
    string text = textRange.Text;

    // Write text
    textRange.Text = "New content";

    // Format text
    font = textRange.Font;
    font.Size = 24;
    font.Bold = -1;       // msoTrue
    font.Italic = 0;      // msoFalse
    font.Color.RGB = 0xFF0000; // Red (BGR format in COM)
    font.Name = "Calibri";
}
finally
{
    if (font != null) ComUtilities.Release(ref font!);
    if (textRange != null) ComUtilities.Release(ref textRange!);
    if (textFrame != null) ComUtilities.Release(ref textFrame!);
}
```

### Paragraph-Level Formatting

```csharp
dynamic? paragraphs = null;
dynamic? paragraph = null;
dynamic? paraFont = null;
try
{
    paragraphs = textFrame.TextRange.Paragraphs();
    for (int i = 1; i <= paragraphs.Count; i++)
    {
        paragraph = paragraphs.Item(i);
        paraFont = paragraph.Font;
        paraFont.Size = 18;
        ComUtilities.Release(ref paraFont!);
        ComUtilities.Release(ref paragraph!);
    }
}
finally
{
    if (paraFont != null) ComUtilities.Release(ref paraFont!);
    if (paragraph != null) ComUtilities.Release(ref paragraph!);
    if (paragraphs != null) ComUtilities.Release(ref paragraphs!);
}
```

### HasTextFrame Check

```csharp
// Not all shapes have text frames - check first
int hasText = Convert.ToInt32(shape.HasTextFrame);
if (hasText == -1) // msoTrue
{
    dynamic textFrame = shape.TextFrame;
    // ... use textFrame ...
    ComUtilities.Release(ref textFrame!);
}
```

---

## COM Object Cleanup (CRITICAL)

### Standard Pattern

```csharp
dynamic? shape = null;
dynamic? textFrame = null;
dynamic? textRange = null;
try
{
    shape = slide.Shapes.Item(1);
    textFrame = shape.TextFrame;
    textRange = textFrame.TextRange;
    // ... operations ...
}
finally
{
    // Release in reverse acquisition order
    if (textRange != null) ComUtilities.Release(ref textRange!);
    if (textFrame != null) ComUtilities.Release(ref textFrame!);
    if (shape != null) ComUtilities.Release(ref shape!);
}
```

### Loop Cleanup

```csharp
for (int i = 1; i <= slide.Shapes.Count; i++)
{
    dynamic? shape = null;
    try
    {
        shape = slide.Shapes.Item(i);
        // ... process shape ...
    }
    finally
    {
        if (shape != null) ComUtilities.Release(ref shape!);
    }
}
```

### Reverse-Order Deletion in Loops

```csharp
// When deleting shapes, iterate in reverse to avoid index shifting
for (int i = slide.Shapes.Count; i >= 1; i--)
{
    dynamic? shape = null;
    try
    {
        shape = slide.Shapes.Item(i);
        if (ShouldDelete(shape))
        {
            shape.Delete();
        }
    }
    finally
    {
        if (shape != null) ComUtilities.Release(ref shape!);
    }
}
```

---

## Common PowerPoint COM Quirks

| Quirk | Detail |
|-------|--------|
| **1-based indexing** | All collections: `Slides.Item(1)`, `Shapes.Item(1)`, `Paragraphs(1)` |
| **msoTrue = -1** | Boolean tri-state: `msoTrue = -1`, `msoFalse = 0`, `msoCTrue = 1` |
| **Points, not pixels** | Positions/sizes in points (72 points = 1 inch) |
| **BGR color order** | `Color.RGB` uses BGR: red = `0x0000FF`, blue = `0xFF0000` |
| **double returns** | Numeric properties return `double`; use `Convert.ToInt32()` |
| **HasTextFrame** | Must check before accessing `TextFrame`; not all shapes support it |
| **SlideID vs Index** | `SlideIndex` changes on reorder; `SlideID` is stable |
| **Placeholder access** | Use `shape.PlaceholderFormat.Type` only when `shape.Type == 14` |
| **Delete reindexes** | Deleting shape/slide shifts subsequent indices; iterate in reverse |
| **GroupItems** | Grouped shapes: access children via `shape.GroupItems.Item(i)` |

---

## Common Mistakes

| Mistake | Fix |
|---------|-----|
| 0-based indexing | PowerPoint is 1-based |
| `bool` for tri-state | Use `int`: -1 (true), 0 (false) |
| Pixel measurements | Use points (72pt = 1 inch) |
| RGB color order | PowerPoint COM uses BGR |
| `int x = shape.Left` | Use `Convert.ToSingle()` or `float` for position properties |
| Missing HasTextFrame check | Always check before accessing TextFrame |
| Forward-loop deletion | Iterate in reverse when deleting |
| Not releasing COM objects | `try-finally` + `ComUtilities.Release()` |
| Catching exceptions in Core | Let `batch.Execute()` handle via TaskCompletionSource |
