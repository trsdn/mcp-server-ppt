---
applyTo: "src/PptMcp.Core/Commands/**/*.cs,tests/**/*.cs"
---

# PowerPoint COM Patterns - LLM Quick Reference

> **PowerPoint-specific COM patterns and what to watch for**

## NOTE: PowerPoint vs Excel Differences

PowerPoint does not have the same data connection model as Excel. The concepts of Power Query, OLEDB/ODBC connections, QueryTables, and data loading do not apply to PowerPoint presentations.

## PowerPoint-Specific COM Patterns

### Slide Operations
```
1. Add slide → presentation.Slides.Add(index, layout)
2. Delete slide → slide.Delete()
3. Duplicate slide → slide.Duplicate()
4. Move slide → slide.MoveTo(newIndex)
```

### Shape Operations
```
1. Add shape → slide.Shapes.AddShape(type, left, top, width, height)
2. Add text box → slide.Shapes.AddTextbox(orientation, left, top, width, height)
3. Add picture → slide.Shapes.AddPicture(filename, linkToFile, saveWithDocument, left, top)
4. Add table → slide.Shapes.AddTable(numRows, numColumns, left, top, width, height)
```

### Common Mistakes to Avoid

1. **Using 0-based indexing** - PowerPoint collections are 1-based
2. **Not releasing COM objects** - Always use try-finally with ComUtilities.Release()
3. **Assuming integer return types** - COM numeric properties return `double`
4. **Forgetting to save** - Call `presentation.Save()` before close for persistence

## Security

**Always validate file paths before operations** - Never allow path traversal or access to unexpected directories.

---

## Developer Reference (Implementation Details)

<details>
<summary>Click to expand developer implementation notes</summary>

### Implementation Notes

**Slides.Add() method:**

Use the COM Add method with parameters: Index (1-based position), Layout (ppLayoutEnum value).

### Shape Type Detection

Shape types are identified via `shape.Type` property which returns `MsoShapeType` enum values. Common types:
- 1 (msoAutoShape) - Basic shapes
- 6 (msoGroup) - Grouped shapes  
- 13 (msoPicture) - Images
- 14 (msoPlaceholder) - Placeholder shapes
- 17 (msoTextBox) - Text boxes
- 19 (msoTable) - Tables

### Test Strategy

- **Slide operations** - Use for lifecycle tests (Add, Delete, Move, Duplicate)
- **Shape operations** - Use for shape CRUD and property tests
- **Text operations** - Use for TextFrame/TextRange manipulation tests

</details>
