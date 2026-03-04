---
name: ppt-cli
description: >
  Automate Microsoft PowerPoint on Windows via CLI. Use when creating, reading,
  or modifying PowerPoint presentations from scripts, CI/CD, or coding agents.
  Supports Slides, Shapes, Text, Charts, Animations, Transitions.
  Triggers: PowerPoint, presentation, pptx, pptcli, CLI automation.
---

# PowerPoint Automation with pptcli

## Preconditions

- Windows host with Microsoft PowerPoint installed (2016+)
- Uses COM interop — does NOT work on macOS or Linux
- Install: `dotnet tool install --global PptMcp.CLI`

## Workflow Checklist

| Step | Command | When |
|------|---------|------|
| 1. Session | `session create/open` | Always first |
| 2. Slides | `slide create/duplicate` | If needed |
| 3. Add content | `shape`, `text`, `chart` | Add shapes, text, charts |
| 4. Save & close | `session close --save` | Always last |

> **10+ commands?** Use `pptcli -q batch --input commands.json` — sends all commands in one process with automatic session management. See Rule 8.

## CRITICAL RULES (MUST FOLLOW)

> **⚡ Building dashboards or bulk operations?** Skip to **Rule 8: Batch Mode** — it eliminates per-command process overhead and auto-manages session IDs.

### Rule 1: NEVER Ask Clarifying Questions

Execute commands to discover the answer instead:

| DON'T ASK | DO THIS INSTEAD |
|-----------|-----------------|
| "Which file should I use?" | `pptcli -q session list` |
| "Which slide has the content?" | `pptcli -q slide list --session <id>` |
| "What shapes are on this slide?" | `pptcli -q shape list --session <id>` |

**You have commands to answer your own questions. USE THEM.**

### Rule 2: Always End With a Text Summary

**NEVER end your turn with only a command execution.** After completing all operations, always provide a brief text message confirming what was done. Silent command-only responses are incomplete.

### Rule 3: Session Lifecycle

**Creating vs Opening Files:**
```powershell
# NEW file - use session create
pptcli -q session create C:\path\newfile.pptx  # Creates file + returns session ID

# EXISTING file - use session open
pptcli -q session open C:\path\existing.pptx   # Opens file + returns session ID
```

**CRITICAL: Use `session create` for new files. `session open` on non-existent files will fail!**

**CRITICAL: ALWAYS use the session ID returned by `session create` or `session open` in subsequent commands. NEVER guess or hardcode session IDs. The session ID is in the JSON output (e.g., `{"sessionId":"abc123"}`). Parse it and use it.**

```powershell
# Example: capture session ID from output, then use it
pptcli -q session create C:\path\file.pptx     # Returns JSON with sessionId
pptcli -q slide list --session <returned-session-id>
pptcli -q session close --session <returned-session-id> --save
```

**Unclosed sessions leave PowerPoint processes running, locking files.**

### Rule 4: Slide Layout and Masters

Slides inherit layouts from slide masters:

```powershell
pptcli -q slide create --session <id> --layout "Title and Content"  # Use a layout
pptcli -q slide list --session <id>                                  # See all slides
```

### Rule 5: Shape and Text Workflow

**BEST PRACTICE: Add shapes then set their text/properties**

```powershell
# Step 1: Add a shape to a slide
pptcli -q shape create --session <id> --slide-index 1 --shape-type rectangle

# Step 2: Set text content
pptcli -q text set --session <id> --slide-index 1 --shape-name "Rectangle 1" --text "Hello World"

# Step 3: Apply formatting
pptcli -q shape format --session <id> --slide-index 1 --shape-name "Rectangle 1" --fill-color "#0078D4"
```

### Rule 6: Report File Errors Immediately

If you see "File not found" or "Path not found" - STOP and report to user. Don't retry.

### Rule 7: Use Animations and Transitions Thoughtfully

When adding animations to multiple shapes, apply them in logical order:

```powershell
# 1. Add entrance animation to title
pptcli -q animation add --session 1 --slide-index 1 --shape-name "Title" --effect "fade-in"

# 2. Add entrance animation to content
pptcli -q animation add --session 1 --slide-index 1 --shape-name "Content" --effect "fly-in" --delay 0.5

# 3. Set slide transition
pptcli -q transition set --session 1 --slide-index 1 --effect "push" --duration 1.0
```

### Rule 8: Use Batch Mode for Bulk Operations (10+ commands)

When executing 10+ commands on the same file, use `pptcli batch` to send all commands in a single process launch. This avoids per-process startup overhead and terminal buffer saturation.

```powershell
# Create a JSON file with all commands
@'
[
  {"command": "session.open", "args": {"filePath": "C:\\path\\file.pptx"}},
  {"command": "slide.create", "args": {"layout": "Blank"}},
  {"command": "shape.create", "args": {"slideIndex": 2, "shapeType": "rectangle"}},
  {"command": "session.close", "args": {"save": true}}
]
'@ | Set-Content commands.json

# Execute all commands at once
pptcli -q batch --input commands.json
```

**Key features:**
- **Session auto-capture**: `session.open`/`create` result sessionId auto-injected into subsequent commands — no need to parse and pass session IDs
- **NDJSON output**: One JSON result per line: `{"index": 0, "command": "...", "success": true, "result": {...}}`
- **`--stop-on-error`**: Exit on first failure (default: continue all)
- **`--session <id>`**: Pre-set session ID for all commands (skip session.open)

**Input formats:**
- JSON array from file: `pptcli -q batch --input commands.json`
- NDJSON from stdin: `Get-Content commands.ndjson | pptcli -q batch`

## CLI Command Reference

> Auto-generated from `pptcli --help`. Use these exact parameter names.


### animation

Animation effect operations: list, add, remove, reorder effects on slides.

**Actions:** `list`, `add`, `remove`, `clear`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | (required) |
| `--shape-name` | Name of the target shape (required for: add) |
| `--effect-type` | MsoAnimEffect integer (e.g., 1=Appear, 2=Fly, 10=Fade, 16=Wipe) (required for: add) |
| `--trigger-type` | 1=OnClick (default), 2=WithPrevious, 3=AfterPrevious (required for: add) |
| `--effect-index` | (required for: remove) |



### background

Slide background: get, set solid color, set image, reset to master.

**Actions:** `get`, `set-color`, `reset`, `set-image`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | 1-based slide index (required) |
| `--color-hex` | Hex color string (#RRGGBB) (required for: set-color) |
| `--image-path` | Path to the image file (required for: set-image) |



### chart

Embedded chart operations: create, get info, set title, set type, delete.

**Actions:** `create`, `get-info`, `set-title`, `set-type`, `delete`, `set-data`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | 1-based slide index (required) |
| `--chart-type` | XlChartType integer (e.g., 4=xlLine, 5=xlPie, 51=xlColumnClustered, -4169=xl3DColumn) (required for: create, set-type) |
| `--left` | Position from left in points (required for: create) |
| `--top` | Position from top in points (required for: create) |
| `--width` | Width in points (required for: create) |
| `--height` | Height in points (required for: create) |
| `--shape-name` | (required for: get-info, set-title, set-type, delete, set-data) |
| `--title` | (required for: set-title) |
| `--values` | 2D array of values (rows × columns) (required for: set-data) |



### comment

Slide comments: list, add, delete.

**Actions:** `list`, `add`, `delete`, `clear`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | 1-based slide index, or 0 for all slides (required) |
| `--text` | Comment text (required for: add) |
| `--author` | Author name (required for: add) |
| `--left` | Horizontal position in points (0 = top-left) (required for: add) |
| `--top` | Vertical position in points (0 = top-left) (required for: add) |
| `--comment-index` | 1-based comment index (required for: delete) |



### customshow

Custom slide show management: list, create, delete, run.

**Actions:** `list`, `create`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--show-name` | Name for the custom show (required for: create, delete) |
| `--slide-indices` | Comma-separated 1-based slide indices (e.g. "1,3,5") (required for: create) |



### design

Theme and design operations: list designs, apply themes, get theme colors.

**Actions:** `list`, `apply-theme`, `get-colors`, `list-color-schemes`

| Parameter | Description |
|-----------|-------------|
| `--theme-path` | Full path to .thmx theme file (required for: apply-theme) |
| `--design-index` | 1-based design index (0 = first design) (required for: get-colors) |



### docproperty

Document property management: read and write presentation metadata like title, author, subject, keywords.

**Actions:** `get`, `set`, `get-custom`, `set-custom`

| Parameter | Description |
|-----------|-------------|
| `--title` | Presentation title (required for: set) |
| `--subject` | Subject or topic (required for: set) |
| `--author` | Author name (required for: set) |
| `--keywords` | Keywords for search (comma-separated) (required for: set) |
| `--comments` | Description or comments (required for: set) |
| `--company` | Company or organization name (required for: set) |
| `--category` | Category (required for: set) |
| `--property-name` | Custom property name (required for: get-custom, set-custom) |
| `--property-value` | Property value (string) (required for: set-custom) |



### export

Export presentations to PDF, images, or other formats.

**Actions:** `to-pdf`, `slide-to-image`, `to-video`, `print`, `save-as`, `all-slides-to-images`

| Parameter | Description |
|-----------|-------------|
| `--destination-path` | Output PDF file path (required for: to-pdf, slide-to-image, to-video, save-as) |
| `--slide-index` | 1-based slide index (required for: slide-to-image) |
| `--width` | Image width in pixels (default: 1920) (required for: slide-to-image, all-slides-to-images) |
| `--height` | Image height in pixels (default: 1080) (required for: slide-to-image, all-slides-to-images) |
| `--default-slide-seconds` | Seconds per slide (default: 5) (required for: to-video) |
| `--resolution` | 1=1080p, 2=720p, 3=480p (required for: to-video) |
| `--copies` | Number of copies (default: 1) (required for: print) |
| `--from-slide` | First slide to print (0 = from beginning) (required for: print) |
| `--to-slide` | Last slide to print (0 = to end) (required for: print) |
| `--format` | Format code (1-7) (required for: save-as) |
| `--destination-directory` | Directory to save images (required for: all-slides-to-images) |



### file

File management commands for PowerPoint presentations. Handles file validation and metadata retrieval.

**Actions:** `test`

| Parameter | Description |
|-----------|-------------|
| `--file-path` | Path to the .pptx or .pptm file (required) |



### headerfooter

Presentation headers and footers: get settings, set date/page number/footer text.

**Actions:** `get`, `set`

| Parameter | Description |
|-----------|-------------|
| `--footer-text` | Footer text (null = don't change) |
| `--show-footer` | Show footer on slides |
| `--show-slide-number` | Show slide numbers |
| `--show-date` | Show date/time |



### hyperlink

Hyperlink management: add, remove, and get hyperlinks on shapes and text.

**Actions:** `add`, `get`, `remove`, `list`, `validate`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | 1-based slide index (required for: add, get, remove) |
| `--shape-name` | Name of the shape to add hyperlink to (required for: add, get, remove) |
| `--address` | URL (https://...) or empty for slide link (required for: add) |
| `--sub-address` | Slide number for internal links (e.g. '3' to jump to slide 3), or empty |
| `--screen-tip` | Optional tooltip text shown on hover |



### image

Image operations: insert pictures into slides.

**Actions:** `insert`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | 1-based slide index (required) |
| `--image-path` | Path to the image file (required) |
| `--left` | Position from left in points (required) |
| `--top` | Position from top in points (required) |
| `--width` | Width in points (0 = original) (required) |
| `--height` | Height in points (0 = original) (required) |



### master

Slide master and layout operations: list masters, list layouts, get placeholders.

**Actions:** `list`, `list-shapes`, `edit-shape-text`

| Parameter | Description |
|-----------|-------------|
| `--master-index` | 1-based slide master index (required for: list-shapes, edit-shape-text) |
| `--shape-name` | Name of the shape to edit (required for: edit-shape-text) |
| `--text` | New text content (required for: edit-shape-text) |



### media

Media management: insert audio and video files into slides. Supports linking or embedding media files.

**Actions:** `insert-audio`, `insert-video`, `get-info`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | 1-based slide index (required) |
| `--file-path` | Full path to the audio file (required for: insert-audio, insert-video) |
| `--left` | Position from left in points (required for: insert-audio, insert-video) |
| `--top` | Position from top in points (required for: insert-audio, insert-video) |
| `--link-to-file` | If true, link to file instead of embedding (smaller file size) (required for: insert-audio, insert-video) |
| `--save-with-document` | If true, save media with document when linking (required for: insert-audio) |
| `--width` | Width in points (0 = use video native width) (required for: insert-video) |
| `--height` | Height in points (0 = use video native height) (required for: insert-video) |
| `--shape-name` | (required for: get-info) |



### notes

Speaker notes: get, set, clear.

**Actions:** `get`, `set`, `clear`, `append`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | (required) |
| `--text` | (required for: set, append) |



### pagesetup

Slide size and page setup operations.

**Actions:** `get`, `set-size`

| Parameter | Description |
|-----------|-------------|
| `--slide-width` | Slide width in points (1 inch = 72 points). 0 = don't change. (required for: set-size) |
| `--slide-height` | Slide height in points. 0 = don't change. (required for: set-size) |



### placeholder

Slide placeholder operations: list available placeholders, fill text.

**Actions:** `list`, `set-text`, `set-image`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | 1-based slide index (required) |
| `--placeholder-index` | 1-based placeholder index (required for: set-text, set-image) |
| `--text` | Text to set (required for: set-text) |
| `--image-path` | Absolute path to the image file (required for: set-image) |



### section

Presentation section management: list, add, rename, delete, and move sections. Sections group slides for easier navigation and organization.

**Actions:** `list`, `add`, `rename`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--section-name` | Name for the new section (required for: add) |
| `--slide-index` | 1-based slide index where the section starts (required for: add) |
| `--section-index` | 1-based section index (required for: rename, delete) |
| `--new-name` | New section name (required for: rename) |



### shape

Shape management: list, read, create, move, resize, delete, z-order.

**Actions:** `list`, `read`, `add-textbox`, `add-shape`, `move-resize`, `delete`, `z-order`, `set-fill`, `set-line`, `set-rotation`, `group`, `ungroup`, `set-alt-text`, `copy-to-slide`, `set-shadow`, `add-connector`, `merge`, `duplicate`, `flip`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | (required) |
| `--shape-name` | (required for: read, move-resize, delete, z-order, set-fill, set-line, set-rotation, ungroup, set-alt-text, copy-to-slide, set-shadow, duplicate, flip) |
| `--left` | Position from left in points (required for: add-textbox, add-shape) |
| `--top` | Position from top in points (required for: add-textbox, add-shape) |
| `--width` | Width in points (required for: add-textbox, add-shape) |
| `--height` | Height in points (required for: add-textbox, add-shape) |
| `--text` | Initial text content (required for: add-textbox) |
| `--auto-shape-type` | MsoAutoShapeType integer (1=Rectangle, 9=Oval, etc.) (required for: add-shape) |
| `--z-order-cmd` | 1=BringToFront, 2=SendToBack, 3=BringForward, 4=SendBackward (required for: z-order) |
| `--color-hex` | Hex color string like #FF0000 for red, or 'none' for no fill (required for: set-fill, set-line) |
| `--line-width` | Line width in points (default 0.75) (required for: set-line) |
| `--degrees` | (required for: set-rotation) |
| `--shape-names` | Comma-separated list of shape names to group (required for: group, merge) |
| `--alt-text` | (required for: set-alt-text) |
| `--target-slide-index` | 1-based target slide index (required for: copy-to-slide) |
| `--visible` | Show or hide shadow (required for: set-shadow) |
| `--offset-x` | Shadow offset X in points (required for: set-shadow) |
| `--offset-y` | Shadow offset Y in points (required for: set-shadow) |
| `--connector-type` | 1=Straight, 2=Elbow, 3=Curve (required for: add-connector) |
| `--start-shape-name` | Starting shape name (required for: add-connector) |
| `--end-shape-name` | Ending shape name (required for: add-connector) |
| `--merge-type` | 1=Union, 2=Combine, 3=Fragment, 4=Intersect, 5=Subtract (required for: merge) |
| `--flip-type` | 0=Horizontal, 1=Vertical (required for: flip) |



### shapealign

Shape alignment and distribution operations.

**Actions:** `align`, `distribute`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | 1-based slide index (required) |
| `--shape-names` | Comma-separated shape names (required) |
| `--align-type` | Alignment type (0-5) (required for: align) |
| `--distribute-type` | 0=Horizontally, 1=Vertically (required for: distribute) |



### slide

Slide lifecycle commands: list, read, create, duplicate, move, delete.

**Actions:** `list`, `read`, `create`, `duplicate`, `move`, `delete`, `apply-layout`, `set-name`, `clone-with-replace`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | 1-based slide index (required for: read, duplicate, move, delete, apply-layout, set-name, clone-with-replace) |
| `--position` | 1-based insert position (0 = at end) (required for: create) |
| `--layout-name` | Layout name from the slide master (e.g. "Title Slide", "Blank") (required for: create, apply-layout) |
| `--new-position` | 1-based target position (required for: move) |
| `--name` | New name for the slide (required for: set-name) |
| `--count` | Number of clones to create (required for: clone-with-replace) |
| `--search-text` | Text to search for in each clone (required for: clone-with-replace) |
| `--replace-text` | Text to replace with in each clone (required for: clone-with-replace) |



### slideimport

Import slides from another presentation file.

**Actions:** `import`

| Parameter | Description |
|-----------|-------------|
| `--source-file-path` | Path to the source .pptx file (required) |
| `--slide-indices` | Comma-separated 1-based slide indices to import (empty = all) (required) |
| `--insert-at` | Position to insert (0 = at end) (required) |



### slideshow

Slideshow presentation mode: start, stop, navigate, get status.

**Actions:** `start`, `stop`, `goto-slide`, `get-status`

| Parameter | Description |
|-----------|-------------|
| `--start-slide` | 1-based slide to start from (0 = beginning) (required for: start) |
| `--slide-index` | 1-based target slide index (required for: goto-slide) |



### slidetable

Table shape operations: create, read, write cells, add/delete rows and columns, merge cells.

**Actions:** `create`, `read`, `write-cell`, `add-row`, `add-column`, `delete-row`, `delete-column`, `merge-cells`, `read-cell`, `format-cell`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | 1-based slide index (required) |
| `--rows` | Number of rows (required for: create) |
| `--columns` | Number of columns (required for: create) |
| `--left` | Position from left in points (required for: create) |
| `--top` | Position from top in points (required for: create) |
| `--width` | Width in points (required for: create) |
| `--height` | Height in points (required for: create) |
| `--shape-name` | Name of the table shape (required for: read, write-cell, add-row, add-column, delete-row, delete-column, merge-cells, read-cell, format-cell) |
| `--row` | 1-based row index (required for: write-cell, delete-row, read-cell, format-cell) |
| `--column` | 1-based column index (required for: write-cell, delete-column, read-cell, format-cell) |
| `--value` | Cell value to set (required for: write-cell) |
| `--position` | 1-based position to insert (-1 = at end) (required for: add-row, add-column) |
| `--start-row` | 1-based start row (required for: merge-cells) |
| `--start-column` | 1-based start column (required for: merge-cells) |
| `--end-row` | 1-based end row (required for: merge-cells) |
| `--end-column` | 1-based end column (required for: merge-cells) |
| `--fill-color` | Hex fill color (#RRGGBB) or null to skip |
| `--font-bold` | Set bold (null = don't change) |
| `--font-size` | Set font size (0 = don't change) (required for: format-cell) |
| `--text-align` | Text alignment: left, center, right (null = don't change) |



### smartart

SmartArt diagram operations: create, add/remove nodes, change layout.

**Actions:** `get-info`, `add-node`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | 1-based slide index (required) |
| `--shape-name` | Name of the SmartArt shape (required) |
| `--text` | Text for the new node (required for: add-node) |



### tag

Custom tags/metadata on slides and shapes.

**Actions:** `list`, `set`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | 1-based slide index (required) |
| `--shape-name` | Shape name (null/empty = slide-level tags) |
| `--tag-name` | Tag name (case-insensitive) (required for: set, delete) |
| `--tag-value` | Tag value (required for: set) |



### text

Text operations within shapes: get, set, format, find, replace.

**Actions:** `get`, `set`, `find`, `replace`, `format`, `format-advanced`, `word-count`, `alt-text-audit`, `empty-placeholder-audit`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | (required) |
| `--shape-name` | (required for: get, set, format, format-advanced) |
| `--text` | (required for: set) |
| `--search-text` | Text to find (required for: find, replace) |
| `--replace-text` | Replacement text (required for: replace) |
| `--font-name` |  |
| `--font-size` |  |
| `--bold` |  |
| `--italic` |  |
| `--color` |  |
| `--alignment` |  |
| `--vertical-alignment` |  |
| `--underline` | Set underline (null = don't change) |
| `--strikethrough` | Set strikethrough (null = don't change) |
| `--subscript` | Set subscript (null = don't change) |
| `--superscript` | Set superscript (null = don't change) |



### transition

Slide transition effects: get, set, remove.

**Actions:** `get`, `set`, `remove`, `copy-to-all`

| Parameter | Description |
|-----------|-------------|
| `--slide-index` | (required) |
| `--transition-type` | PpEntryEffect enum value (e.g. 3844=Fade, 3849=Push) (required for: set) |
| `--duration` | Duration in seconds (required for: set) |
| `--advance-on-click` | Whether to advance on mouse click (required for: set) |
| `--advance-after-time` | Auto-advance after N seconds (0 = disabled) (required for: set) |



### vba

VBA macro operations: list modules, view/import/delete code, run macros. Requires VBA trust settings enabled in PowerPoint.

**Actions:** `list`, `view`, `import`, `delete`, `run`

| Parameter | Description |
|-----------|-------------|
| `--module-name` | (required for: view, import, delete) |
| `--code` | VBA code to import (required for: import) |
| `--module-type` | 1=Standard, 2=ClassModule (default: 1) (required for: import) |
| `--macro-name` | Fully qualified macro name (e.g., "Module1.MyMacro") (required for: run) |



### window

PowerPoint window management: get info, minimize, restore, maximize.

**Actions:** `get-info`, `minimize`, `restore`, `maximize`, `set-zoom`

| Parameter | Description |
|-----------|-------------|
| `--zoom-percent` | Zoom percentage (e.g. 100 for 100%) (required for: set-zoom) |




## Common Pitfalls

### Slide Indices Are 1-Based

Slide indices start at 1, not 0. `--slide-index 0` is invalid and will error.

### --timeout Must Be Greater Than Zero

When using `--timeout`, the value must be a positive integer (seconds). `--timeout 0` is invalid and will error. Omit `--timeout` entirely to use the default (300 seconds for most operations).

### Shape Names Must Be Exact

Shape names are case-sensitive and must match exactly. Use `shape list` to discover correct names before targeting shapes.

### JSON Values Format

`--values` takes a JSON string wrapped in single quotes:
```powershell
# CORRECT: JSON with single-quote wrapper
--values '{"text": "Hello World", "fontSize": 24}'

# WRONG: Missing quotes
--values {text: Hello World}
```

### List Parameters Use JSON Arrays

Parameters that accept lists require JSON array format:
```powershell
# CORRECT: JSON array with single-quote wrapper
--selected-items '["Slide 1","Slide 3"]'

# WRONG: Comma-separated string (not valid)
--selected-items "Slide 1,Slide 3"
```

## Reference Documentation

- [Core execution rules and LLM guidelines](./references/behavioral-rules.md)
- [Common mistakes to avoid](./references/anti-patterns.md)
- [Presentation workflows and patterns](./references/workflows.md)
- [Charts, positioning, and formatting](./references/chart.md)
- [CLI commands reference](./references/cli-commands.md)
- [Presentation design best practices](./references/dashboard.md)
- [Screenshot and visual verification](./references/screenshot.md)
- [Table operations](./references/table.md)
- [Window management](./references/window.md)
- [Agent mode patterns](./references/ppt_agent_mode.md)
