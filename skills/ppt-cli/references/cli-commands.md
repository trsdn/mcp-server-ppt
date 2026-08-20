# CLI Command Reference

> Auto-generated from \pptcli --help\. Do not edit manually.

## service

Service lifecycle management: start, stop, status

## batch

Execute multiple commands from a JSON file or stdin. Outputs NDJSON (one result

| Parameter | Description |
|-----------|-------------|
| `--input` | JSON file with command array. Use '-' for stdin |
| `--session` | Default session ID for all commands. Overridden |

## diag

Diagnostic commands: ping, echo, validate-params

## accessibility

Accessibility audit: check alt text, title placeholders, reading order

**Actions:** `audit`, `get-reading-order`, `set-reading-order`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--slide-index` | 1-based slide index (required for: |
| `--shape-names` | Comma-separated shape names in desired |
| `--output` | Write output to file instead of stdout. |

## animation

Animation effect operations: list, add, remove, reorder effects on slides

**Actions:** `list`, `add`, `remove`, `clear`, `set-timing`, `reorder`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--slide-index` | 1-based slide index (required) |
| `--shape-name` | Name of the target shape (required for: |
| `--effect-type` | MsoAnimEffect integer (e.g., 1=Appear, |
| `--trigger-type` | 1=OnClick (default), 2=WithPrevious, |
| `--effect-index` | 1-based index of the effect in the |
| `--duration` | Duration in seconds (required for: |
| `--delay` | Delay before start in seconds (required |
| `--new-index` | 1-based target position in the sequence |
| `--output` | Write output to file instead of stdout. |

## background

Slide background: get, set solid color, set image, reset to master

**Actions:** `get`, `set-color`, `reset`, `set-image`, `set-gradient`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' |
| `--slide-index` | 1-based slide index (required) |
| `--color-hex` | Hex color string (#RRGGBB) (required |
| `--image-path` | Path to the image file (required |
| `--color1` | First gradient color as hex |
| `--color2` | Second gradient color as hex |
| `--gradient-style` | 1=Horizontal, 2=Vertical, |
| `--output` | Write output to file instead of |

## chart

Embedded chart operations: create, get info, set title, set type, delete

**Actions:** `create`, `get-info`, `set-title`, `set-type`, `delete`, `set-data`, `set-legend`, `read-data`, `set-axis-title`, `toggle-data-table`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--slide-index` | 1-based slide index (required) |
| `--chart-type` | XlChartType integer (e.g., 4=xlLine, |
| `--left` | Position from left in points (required |
| `--top` | Position from top in points (required for: |
| `--width` | Width in points (required for: create) |
| `--height` | Height in points (required for: create) |
| `--shape-name` | Name of the chart shape (required for: |
| `--title` | Title text for the axis (required for: |
| `--values` | 2D array of values (rows × columns) |
| `--visible` | Whether the legend is visible (required |
| `--position` | Legend position: -4107=Bottom, -4131=Left, |
| `--axis-type` | Axis type: 1=Category(X), 2=Value(Y) |
| `--output` | Write output to file instead of stdout. |

## comment

Slide comments: list, add, delete

**Actions:** `list`, `add`, `delete`, `clear`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--slide-index` | 1-based slide index, or 0 for all |
| `--text` | Comment text (required for: add) |
| `--author` | Author name (required for: add) |
| `--left` | Horizontal position in points (0 = |
| `--top` | Vertical position in points (0 = |
| `--comment-index` | 1-based comment index (required for: |
| `--output` | Write output to file instead of |

## customshow

Custom slide show management: list, create, delete, run

**Actions:** `list`, `create`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--show-name` | Name for the custom show (required |
| `--slide-indices` | Comma-separated 1-based slide indices |
| `--output` | Write output to file instead of |

## design

Design operations: themes, colors, fonts, and design knowledge catalog. THEME

**Actions:** `list`, `apply-theme`, `get-colors`, `list-color-schemes`, `get-fonts`, `list-archetypes`, `get-archetype`, `list-palettes`, `get-palette`, `list-style-profiles`, `get-style-profile`, `list-layout-grids`, `get-layout-grid`, `list-density-profiles`, `get-density-profile`, `get-context-model`, `get-deck-sequence`, `get-slide-patterns`, `get-icon-shapes`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--theme-path` | Full path to .thmx theme file (required |
| `--design-index` | 1-based design index (0 = first design) |
| `--archetype-id` | Archetype id: big-number, |
| `--palette-id` | Palette id: corporate-blue, |
| `--profile-id` | Profile id: consulting, corporate, |
| `--grid-id` | Grid id: single-column, |
| `--density-id` | Density id: D1, D2, D3, D4, D5 (required |
| `--sequence-id` | Sequence id: S1 (Decision), S2 |
| `--output` | Write output to file instead of stdout. |

## docproperty

Document property management: read and write presentation metadata like title,

**Actions:** `get`, `set`, `get-custom`, `set-custom`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' |
| `--title` | Presentation title (required for: |
| `--subject` | Subject or topic (required for: set) |
| `--author` | Author name (required for: set) |
| `--keywords` | Keywords for search |
| `--comments` | Description or comments (required |
| `--company` | Company or organization name |
| `--category` | Category (required for: set) |
| `--property-name` | Custom property name (required for: |
| `--property-value` | Property value (string) (required |
| `--output` | Write output to file instead of |

## export

Export presentations to PDF, images, or other formats

**Actions:** `to-pdf`, `slide-to-image`, `to-video`, `print`, `save-as`, `all-slides-to-images`, `extract-text`, `extract-images`, `save-copy`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from |
| `--destination-path` | Output PDF file path |
| `--slide-index` | 1-based slide index |
| `--width` | Image width in pixels |
| `--height` | Image height in pixels |
| `--default-slide-seconds` | Seconds per slide |
| `--resolution` | 1=1080p, 2=720p, |
| `--copies` | Number of copies |
| `--from-slide` | First slide to print |
| `--to-slide` | Last slide to print (0 |
| `--format` | Format code (1-7) |
| `--destination-directory` | Directory to save |
| `--output` | Write output to file |

## file

File management commands for PowerPoint presentations. Handles file validation

**Actions:** `test`

| Parameter | Description |
|-----------|-------------|
| `--file-path` | Path to the .pptx or .pptm file (required) |
| `--output` | Write output to file instead of stdout. For |

## headerfooter

Presentation headers and footers: get settings, set date/page number/footer text

**Actions:** `get`, `set`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' |
| `--footer-text` | Footer text (null = don't |
| `--show-footer` | Show footer on slides |
| `--show-slide-number` | Show slide numbers |
| `--show-date` | Show date/time |
| `--output` | Write output to file instead of |

## hyperlink

Hyperlink management: add, remove, and get hyperlinks on shapes and text

**Actions:** `add`, `get`, `remove`, `list`, `validate`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--slide-index` | 1-based slide index (required for: add, |
| `--shape-name` | Name of the shape to add hyperlink to |
| `--address` | URL (https://...) or empty for slide link |
| `--sub-address` | Slide number for internal links (e.g. '3' |
| `--screen-tip` | Optional tooltip text shown on hover |
| `--output` | Write output to file instead of stdout. |

## image

Image operations: insert pictures into slides

**Actions:** `insert`, `crop`, `set-brightness-contrast`, `set-transparent-color`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--slide-index` | 1-based slide index (required) |
| `--image-path` | Path to the image file (required for: |
| `--left` | Position from left in points (required |
| `--top` | Position from top in points (required for: |
| `--width` | Width in points (0 = original) (required |
| `--height` | Height in points (0 = original) (required |
| `--shape-name` | Name of the picture shape (required for: |
| `--crop-left` | Crop from left in points (0 = no crop) |
| `--crop-right` | Crop from right in points (0 = no crop) |
| `--crop-top` | Crop from top in points (0 = no crop) |
| `--crop-bottom` | Crop from bottom in points (0 = no crop) |
| `--brightness` | Brightness value (0.0 to 1.0) (required |
| `--contrast` | Contrast value (0.0 to 1.0) (required for: |
| `--color-hex` | Hex color string (#RRGGBB) to make |
| `--output` | Write output to file instead of stdout. |

## master

Slide master and layout operations: list masters, list layouts, get placeholders

**Actions:** `list`, `list-shapes`, `edit-shape-text`, `list-layouts`, `delete-unused`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--master-index` | 1-based slide master index (required |
| `--shape-name` | Name of the shape to edit (required for: |
| `--text` | New text content (required for: |
| `--output` | Write output to file instead of stdout. |

## media

Media management: insert audio and video files into slides. Supports linking or

**Actions:** `insert-audio`, `insert-video`, `get-info`, `set-playback`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session |
| `--slide-index` | 1-based slide index |
| `--file-path` | Full path to the audio file |
| `--left` | Position from left in points |
| `--top` | Position from top in points |
| `--link-to-file` | If true, link to file instead |
| `--save-with-document` | If true, save media with |
| `--width` | Width in points (0 = use |
| `--height` | Height in points (0 = use |
| `--shape-name` | Name of the media shape |
| `--volume` | Volume level (0.0 to 1.0), |
| `--muted` | Mute state, null to leave |
| `--fade-in-seconds` | Fade-in duration in seconds, |
| `--fade-out-seconds` | Fade-out duration in seconds, |
| `--output` | Write output to file instead |

## notes

Speaker notes: get, set, clear

**Actions:** `get`, `set`, `clear`, `append`, `read-all`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--slide-index` | (required for: get, set, clear, append) |
| `--text` | (required for: set, append) |
| `--output` | Write output to file instead of stdout. |

## pagesetup

Slide size and page setup operations

**Actions:** `get`, `set-size`, `set-first-number`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session |
| `--slide-width` | Slide width in points (1 inch |
| `--slide-height` | Slide height in points. 0 = |
| `--first-slide-number` | The number to assign to the |
| `--output` | Write output to file instead |

## placeholder

Slide placeholder operations: list available placeholders, fill text

**Actions:** `list`, `set-text`, `set-image`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' |
| `--slide-index` | 1-based slide index (required) |
| `--placeholder-index` | 1-based placeholder index |
| `--text` | Text to set (required for: |
| `--image-path` | Absolute path to the image |
| `--output` | Write output to file instead |

## printoptions

Manage print options: output type, color mode, framing, fit-to-page, hidden

**Actions:** `get`, `set`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session |
| `--output-type` | 1=Slides, |
| `--color-type` | 1=Color, 2=Grayscale, |
| `--frame-slides` | Whether to frame slides |
| `--fit-to-page` | Whether to fit slides to |
| `--print-hidden-slides` | Whether to include hidden |
| `--output` | Write output to file |

## proofing

Proofing and language operations: check spelling, get/set language for text

**Actions:** `check-spelling`, `set-language`, `get-language`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--slide-index` | 0 for all slides, or specific 1-based |
| `--shape-name` | Empty string for all shapes on slide, or |
| `--language-id` | MsoLanguageID value (e.g. 1033 for English |
| `--output` | Write output to file instead of stdout. |

## section

Presentation section management: list, add, rename, delete, and move sections.

**Actions:** `list`, `add`, `rename`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--section-name` | Name for the new section (required |
| `--slide-index` | 1-based slide index where the section |
| `--section-index` | 1-based section index (required for: |
| `--new-name` | New section name (required for: |
| `--output` | Write output to file instead of |

## shape

Shape management: list, read, create, move, resize, delete, z-order

**Actions:** `list`, `read`, `add-textbox`, `add-shape`, `move-resize`, `delete`, `z-order`, `set-fill`, `set-line`, `set-rotation`, `group`, `ungroup`, `set-alt-text`, `copy-to-slide`, `set-shadow`, `add-connector`, `merge`, `duplicate`, `flip`, `set-text-frame`, `set-gradient-fill`, `set-glow`, `set-reflection`, `set-opacity`, `read-fill`, `read-line`, `find-by-type`, `copy-formatting`, `set-action-settings`, `scale`, `lock-aspect-ratio`, `set-soft-edge`, `read-shadow`, `add-text-effect`, `set-3d`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session |
| `--slide-index` | 1-based slide index |
| `--shape-name` | Name of the shape (required |
| `--left` | Position from left in points |
| `--top` | Position from top in points |
| `--width` | Width in points (required |
| `--height` | Height in points (required |
| `--text` | Initial text content |
| `--auto-shape-type` | MsoAutoShapeType integer |
| `--z-order-cmd` | 1=BringToFront, 2=SendToBack, |
| `--color-hex` | Hex color string like #FF0000 |
| `--line-width` | Line width in points (default |
| `--degrees` | (required for: set-rotation) |
| `--shape-names` | Comma-separated list of shape |
| `--alt-text` | (required for: set-alt-text) |
| `--target-slide-index` | 1-based target slide index |
| `--visible` | Show or hide shadow (required |
| `--offset-x` | Shadow offset X in points |
| `--offset-y` | Shadow offset Y in points |
| `--connector-type` | 1=Straight, 2=Elbow, 3=Curve |
| `--start-shape-name` | Starting shape name (required |
| `--end-shape-name` | Ending shape name (required |
| `--merge-type` | 1=Union, 2=Combine, |
| `--flip-type` | 0=Horizontal, 1=Vertical |
| `--margin-left` | Left margin in points (null = |
| `--margin-right` | Right margin in points (null |
| `--margin-top` | Top margin in points (null = |
| `--margin-bottom` | Bottom margin in points (null |
| `--word-wrap` | Enable/disable word wrap |
| `--auto-size` | 0=None, 1=ShapeToFitText, |
| `--color1` | First gradient color as hex |
| `--color2` | Second gradient color as hex |
| `--gradient-style` | 1=Horizontal, 2=Vertical, |
| `--radius` | Glow radius in points (0 = |
| `--reflection-type` | 0=None, |
| `--opacity` | Opacity value from 0.0 (fully |
| `--shape-type` | MsoShapeType integer |
| `--source-shape-name` | Name of the shape to copy |
| `--target-shape-name` | Name of the shape to apply |
| `--action-type` | 0=None, 1=NextSlide, |
| `--hyperlink-address` | URL for actionType=7 |
| `--scale-x` | Width scale factor (e.g. 1.5 |
| `--scale-y` | Height scale factor (e.g. 1.5 |
| `--locked` | True to lock aspect ratio, |
| `--preset-effect` | MsoPresetTextEffect integer |
| `--font-name` | Font name (e.g. "Arial") |
| `--font-size` | Font size in points (required |
| `--rotation-x` | X-axis rotation in degrees |
| `--rotation-y` | Y-axis rotation in degrees |
| `--rotation-z` | Z-axis rotation in degrees |
| `--bevel-type` | Bevel top type: 0=None, |
| `--bevel-depth` | Bevel top depth in points |
| `--output` | Write output to file instead |

## shapealign

Shape alignment and distribution operations

**Actions:** `align`, `distribute`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' |
| `--slide-index` | 1-based slide index (required) |
| `--shape-names` | Comma-separated shape names |
| `--align-type` | Alignment type (0-5) (required |
| `--distribute-type` | 0=Horizontally, 1=Vertically |
| `--output` | Write output to file instead of |

## slide

Slide lifecycle commands: list, read, create, duplicate, move, delete

**Actions:** `list`, `read`, `create`, `duplicate`, `move`, `delete`, `apply-layout`, `set-name`, `clone-with-replace`, `hide`, `unhide`, `get-thumbnail`, `summary`, `set-display-master`, `copy`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' |
| `--slide-index` | 1-based slide index (required |
| `--position` | 1-based insert position (0 = at |
| `--layout-name` | Layout name from the slide |
| `--new-position` | 1-based target position |
| `--name` | New name for the slide (required |
| `--count` | Number of clones to create |
| `--search-text` | Text to search for in each clone |
| `--replace-text` | Text to replace with in each |
| `--destination-path` | Full path for the output PNG |
| `--display` | Whether to display master shapes |
| `--output` | Write output to file instead of |

## slideimport

Import slides from another presentation file

**Actions:** `import`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' |
| `--source-file-path` | Path to the source .pptx file |
| `--slide-indices` | Comma-separated 1-based slide |
| `--insert-at` | Position to insert (0 = at end) |
| `--output` | Write output to file instead of |

## slideshow

Slideshow presentation mode: start, stop, navigate, get status

**Actions:** `start`, `stop`, `goto-slide`, `get-status`, `configure`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session |
| `--start-slide` | 1-based slide to start from |
| `--slide-index` | 1-based target slide index |
| `--show-type` | 1=Speaker (full screen), |
| `--loop-until-stopped` | Whether to loop the |
| `--show-with-animation` | Whether to show animations |
| `--show-with-narration` | Whether to play narrations |
| `--output` | Write output to file |

## slidetable

Table shape operations: create, read, write cells, add/delete rows and columns,

**Actions:** `create`, `read`, `write-cell`, `add-row`, `add-column`, `delete-row`, `delete-column`, `merge-cells`, `read-cell`, `format-cell`, `write-row`, `read-row`, `set-cell-border`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--slide-index` | 1-based slide index (required) |
| `--rows` | Number of rows (required for: create) |
| `--columns` | Number of columns (required for: create) |
| `--left` | Position from left in points (required |
| `--top` | Position from top in points (required |
| `--width` | Width in points (required for: create, |
| `--height` | Height in points (required for: create) |
| `--shape-name` | Name of the table shape (required for: |
| `--row` | 1-based row index (required for: |
| `--column` | 1-based column index (required for: |
| `--value` | Cell value to set (required for: |
| `--position` | 1-based position to insert (-1 = at end) |
| `--start-row` | 1-based start row (required for: |
| `--start-column` | 1-based start column (required for: |
| `--end-row` | 1-based end row (required for: |
| `--end-column` | 1-based end column (required for: |
| `--fill-color` | Hex fill color (#RRGGBB) or null to skip |
| `--font-bold` | Set bold (null = don't change) |
| `--font-size` | Set font size (0 = don't change) |
| `--text-align` | Text alignment: left, center, right |
| `--values` | Comma-separated values for the row |
| `--color-hex` | Border color as hex (#RRGGBB) (required |
| `--output` | Write output to file instead of stdout. |

## smartart

SmartArt diagram operations: create, add/remove nodes, change layout

**Actions:** `get-info`, `add-node`, `set-layout`, `set-style`, `delete-node`, `change-level`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--slide-index` | 1-based slide index (required) |
| `--shape-name` | Name of the SmartArt shape (required) |
| `--text` | Text for the new node (required for: |
| `--layout-index` | 1-based index into |
| `--style-index` | 1-based index into |
| `--node-index` | 1-based index of the node to delete |
| `--promote` | True to promote (decrease level), false |
| `--output` | Write output to file instead of stdout. |

## tag

Custom tags/metadata on slides and shapes

**Actions:** `list`, `set`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--slide-index` | 1-based slide index (required) |
| `--shape-name` | Shape name (null/empty = slide-level tags) |
| `--tag-name` | Tag name (case-insensitive) (required for: |
| `--tag-value` | Tag value (required for: set) |
| `--output` | Write output to file instead of stdout. |

## text

Text operations within shapes: get, set, format, find, replace

**Actions:** `get`, `set`, `find`, `replace`, `format`, `format-advanced`, `word-count`, `alt-text-audit`, `empty-placeholder-audit`, `set-spacing`, `set-bullets`, `insert-link`, `change-case`, `read-spacing`, `read-bullets`, `insert-symbol`, `insert-datetime`, `insert-slide-number`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session |
| `--slide-index` | 0 for all slides, or |
| `--shape-name` | Shape name (required for: |
| `--text` | (required for: set) |
| `--search-text` | Text to find (required for: |
| `--replace-text` | Replacement text (required |
| `--font-name` | Font name containing the |
| `--font-size` | FontSize |
| `--bold` | Bold |
| `--italic` | Italic |
| `--color` | Color |
| `--alignment` | Alignment |
| `--vertical-alignment` | VerticalAlignment |
| `--underline` | Set underline (null = don't |
| `--strikethrough` | Set strikethrough (null = |
| `--subscript` | Set subscript (null = don't |
| `--superscript` | Set superscript (null = |
| `--line-spacing` | Line spacing in points (null |
| `--space-before` | Space before paragraph in |
| `--space-after` | Space after paragraph in |
| `--character-spacing` | Character spacing in points |
| `--bullet-type` | 0=None, 1=Unnumbered |
| `--bullet-character` | Custom bullet character |
| `--indent-level` | Indent level 0-4 (required |
| `--link-text` | Text to find and make into a |
| `--url` | URL for the hyperlink |
| `--case-type` | 1=Sentence, 2=Lower, |
| `--char-number` | Unicode/character code of |
| `--date-time-format` | PpDateTimeFormat value |
| `--output` | Write output to file instead |

## transition

Slide transition effects: get, set, remove

**Actions:** `get`, `set`, `remove`, `copy-to-all`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session |
| `--slide-index` | 1-based slide index |
| `--transition-type` | PpEntryEffect enum value |
| `--duration` | Duration in seconds (required |
| `--advance-on-click` | Whether to advance on mouse |
| `--advance-after-time` | Auto-advance after N seconds |
| `--output` | Write output to file instead |

## vba

VBA macro operations: list modules, view/import/delete code, run macros.

**Actions:** `list`, `view`, `import`, `delete`, `run`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--module-name` | Name for the new module (required for: |
| `--code` | VBA code to import (required for: import) |
| `--module-type` | 1=Standard, 2=ClassModule (default: 1) |
| `--macro-name` | Fully qualified macro name (e.g., |
| `--output` | Write output to file instead of stdout. |

## window

PowerPoint window management: get info, minimize, restore, maximize

**Actions:** `get-info`, `minimize`, `restore`, `maximize`, `set-zoom`, `set-view`, `get-view`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--zoom-percent` | Zoom percentage (e.g. 100 for 100%) |
| `--view-type` | 1=Normal, 2=Outline, 3=SlideSorter, |
| `--output` | Write output to file instead of stdout. |
