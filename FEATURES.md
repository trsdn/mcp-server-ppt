# PptMcp - Complete Feature Reference

**33 specialized tools with 223 operations for comprehensive PowerPoint automation**

> Generated from the source of truth: the action enums emitted by the source generators
> (`ServiceRegistry.*.g.cs`) plus the hand-written `file` tool. Do not edit counts by hand.

## Official Automation Layers

In addition to the PowerPoint tool surface, the repository documents three official automation layers:

- **MCP Server** — conversational tool surface for rich tool discovery
- **CLI** (`pptcli`) — compact scripting and coding-agent surface
- **Agent Client** (`src\PptMcp.Agent`) — source-side Copilot SDK orchestrator for plan → execute → verify → repair workflows on top of the MCP server

Related docs:

- [Agent Client](src/PptMcp.Agent/README.md)
- [Agent Client Architecture](docs/AGENT-CLIENT.md)
- [Eval Framework](eval/README.md)
- [Archetype Pipeline](docs/ARCHETYPE-PIPELINE.md)

---

## 📁 File Operations (6 operations)

MCP tool / CLI command: `file`

- `open`
- `close`
- `create`
- `list`
- `test`
- `save`

---

## 📄 Slide Operations (15 operations)

MCP tool / CLI command: `slide`

- `list`
- `read`
- `create`
- `duplicate`
- `move`
- `delete`
- `apply-layout`
- `set-name`
- `clone-with-replace`
- `hide`
- `unhide`
- `get-thumbnail`
- `summary`
- `set-display-master`
- `copy`

---

## 🔷 Shape Operations (35 operations)

MCP tool / CLI command: `shape`

- `list`
- `read`
- `add-textbox`
- `add-shape`
- `move-resize`
- `delete`
- `z-order`
- `set-fill`
- `set-line`
- `set-rotation`
- `group`
- `ungroup`
- `set-alt-text`
- `copy-to-slide`
- `set-shadow`
- `add-connector`
- `merge`
- `duplicate`
- `flip`
- `set-text-frame`
- `set-gradient-fill`
- `set-glow`
- `set-reflection`
- `set-opacity`
- `read-fill`
- `read-line`
- `find-by-type`
- `copy-formatting`
- `set-action-settings`
- `scale`
- `lock-aspect-ratio`
- `set-soft-edge`
- `read-shadow`
- `add-text-effect`
- `set-3d`

---

## ✏️ Text Operations (18 operations)

MCP tool / CLI command: `text`

- `get`
- `set`
- `find`
- `replace`
- `format`
- `format-advanced`
- `word-count`
- `alt-text-audit`
- `empty-placeholder-audit`
- `set-spacing`
- `set-bullets`
- `insert-link`
- `change-case`
- `read-spacing`
- `read-bullets`
- `insert-symbol`
- `insert-datetime`
- `insert-slide-number`

---

## 📋 Table Operations (13 operations)

MCP tool / CLI command: `slidetable`

- `create`
- `read`
- `write-cell`
- `add-row`
- `add-column`
- `delete-row`
- `delete-column`
- `merge-cells`
- `read-cell`
- `format-cell`
- `write-row`
- `read-row`
- `set-cell-border`

---

## 📊 Chart Operations (10 operations)

MCP tool / CLI command: `chart`

- `create`
- `get-info`
- `set-title`
- `set-type`
- `delete`
- `set-data`
- `set-legend`
- `read-data`
- `set-axis-title`
- `toggle-data-table`

---

## 🖼️ Image Operations (4 operations)

MCP tool / CLI command: `image`

- `insert`
- `crop`
- `set-brightness-contrast`
- `set-transparent-color`

---

## 🎵 Media Operations (4 operations)

MCP tool / CLI command: `media`

- `insert-audio`
- `insert-video`
- `get-info`
- `set-playback`

---

## 🧩 SmartArt Operations (6 operations)

MCP tool / CLI command: `smartart`

- `get-info`
- `add-node`
- `set-layout`
- `set-style`
- `delete-node`
- `change-level`

---

## 🎨 Design Operations (19 operations)

MCP tool / CLI command: `design`

- `list`
- `apply-theme`
- `get-colors`
- `list-color-schemes`
- `get-fonts`
- `list-archetypes`
- `get-archetype`
- `list-palettes`
- `get-palette`
- `list-style-profiles`
- `get-style-profile`
- `list-layout-grids`
- `get-layout-grid`
- `list-density-profiles`
- `get-density-profile`
- `get-context-model`
- `get-deck-sequence`
- `get-slide-patterns`
- `get-icon-shapes`

---

## 🖌️ Background Operations (5 operations)

MCP tool / CLI command: `background`

- `get`
- `set-color`
- `reset`
- `set-image`
- `set-gradient`

---

## 🎭 Master & Layout Operations (5 operations)

MCP tool / CLI command: `master`

- `list`
- `list-shapes`
- `edit-shape-text`
- `list-layouts`
- `delete-unused`

---

## 📌 Placeholder Operations (3 operations)

MCP tool / CLI command: `placeholder`

- `list`
- `set-text`
- `set-image`

---

## 📃 Header & Footer Operations (2 operations)

MCP tool / CLI command: `headerfooter`

- `get`
- `set`

---

## 📐 Page Setup Operations (3 operations)

MCP tool / CLI command: `pagesetup`

- `get`
- `set-size`
- `set-first-number`

---

## ↔️ Shape Alignment Operations (2 operations)

MCP tool / CLI command: `shapealign`

- `align`
- `distribute`

---

## 🎬 Animation Operations (6 operations)

MCP tool / CLI command: `animation`

- `list`
- `add`
- `remove`
- `clear`
- `set-timing`
- `reorder`

---

## 🔀 Transition Operations (4 operations)

MCP tool / CLI command: `transition`

- `get`
- `set`
- `remove`
- `copy-to-all`

---

## ▶️ Slideshow Operations (5 operations)

MCP tool / CLI command: `slideshow`

- `start`
- `stop`
- `goto-slide`
- `get-status`
- `configure`

---

## 🎞️ Custom Show Operations (3 operations)

MCP tool / CLI command: `customshow`

- `list`
- `create`
- `delete`

---

## 📂 Section Operations (4 operations)

MCP tool / CLI command: `section`

- `list`
- `add`
- `rename`
- `delete`

---

## 📥 Slide Import Operations (1 operation)

MCP tool / CLI command: `slideimport`

- `import`

---

## 📝 Notes Operations (5 operations)

MCP tool / CLI command: `notes`

- `get`
- `set`
- `clear`
- `append`
- `read-all`

---

## 💬 Comment Operations (4 operations)

MCP tool / CLI command: `comment`

- `list`
- `add`
- `delete`
- `clear`

---

## 🔗 Hyperlink Operations (5 operations)

MCP tool / CLI command: `hyperlink`

- `add`
- `get`
- `remove`
- `list`
- `validate`

---

## 🏷️ Tag Operations (3 operations)

MCP tool / CLI command: `tag`

- `list`
- `set`
- `delete`

---

## 🗂️ Document Property Operations (4 operations)

MCP tool / CLI command: `docproperty`

- `get`
- `set`
- `get-custom`
- `set-custom`

---

## ♿ Accessibility Operations (3 operations)

MCP tool / CLI command: `accessibility`

- `audit`
- `get-reading-order`
- `set-reading-order`

---

## 🔤 Proofing Operations (3 operations)

MCP tool / CLI command: `proofing`

- `check-spelling`
- `set-language`
- `get-language`

---

## 📤 Export Operations (9 operations)

MCP tool / CLI command: `export`

- `to-pdf`
- `slide-to-image`
- `to-video`
- `print`
- `save-as`
- `all-slides-to-images`
- `extract-text`
- `extract-images`
- `save-copy`

---

## 🖨️ Print Options Operations (2 operations)

MCP tool / CLI command: `printoptions`

- `get`
- `set`

---

## ⚙️ VBA Macros Operations (5 operations)

MCP tool / CLI command: `vba`

- `list`
- `view`
- `import`
- `delete`
- `run`

---

## 🪟 Window Operations (7 operations)

MCP tool / CLI command: `window`

- `get-info`
- `minimize`
- `restore`
- `maximize`
- `set-zoom`
- `set-view`
- `get-view`

---

## 📊 Total Operations Summary

| Tool | Operations |
|------|-----------|
| `file` | 6 |
| `slide` | 15 |
| `shape` | 35 |
| `text` | 18 |
| `slidetable` | 13 |
| `chart` | 10 |
| `image` | 4 |
| `media` | 4 |
| `smartart` | 6 |
| `design` | 19 |
| `background` | 5 |
| `master` | 5 |
| `placeholder` | 3 |
| `headerfooter` | 2 |
| `pagesetup` | 3 |
| `shapealign` | 2 |
| `animation` | 6 |
| `transition` | 4 |
| `slideshow` | 5 |
| `customshow` | 3 |
| `section` | 4 |
| `slideimport` | 1 |
| `notes` | 5 |
| `comment` | 4 |
| `hyperlink` | 5 |
| `tag` | 3 |
| `docproperty` | 4 |
| `accessibility` | 3 |
| `proofing` | 3 |
| `export` | 9 |
| `printoptions` | 2 |
| `vba` | 5 |
| `window` | 7 |
| **Total (33 tools)** | **223** |

---

## 🚀 Key Capabilities

**Slide Management:**
- Full slide lifecycle (create, duplicate, move, delete)
- Layout and master slide support
- Section organization
- Import slides from other presentations

**Content Creation:**
- Rich shape creation and manipulation (textboxes, auto-shapes, connectors)
- Table creation with cell-level control
- Chart creation and configuration
- Image, audio, and video insertion
- SmartArt management

**Design & Formatting:**
- Theme and color scheme management
- Shape fill, line, shadow, and rotation
- Text formatting (font, size, color, bold, italic)
- Slide background customization
- Shape alignment and distribution

**Presentation Delivery:**
- Slideshow control (start, stop, navigate)
- Animation and transition effects
- Custom slide shows
- Speaker notes management

**Automation & Export:**
- VBA macro execution and management
- Export to PDF, image, and video
- Print support
- Find and replace across slides

**Metadata & Organization:**
- Document properties management
- Comments and review workflow
- Tags for custom metadata
- Header and footer configuration
- Placeholder content management

---

---

## 🔧 Tool Selection Quick Reference

| Task | Tool |
|------|------|
| Open/save/create files | `file` |
| Add/manage slides | `slide` |
| Add/modify shapes | `shape` |
| Edit text content | `text` or `placeholder` |
| Create tables | `slidetable` |
| Create charts | `chart` |
| Insert images | `image` |
| Insert audio/video | `media` |
| Build diagrams | `smartart` |
| Change theme/colors | `design` |
| Set slide background | `background` |
| Edit masters/layouts | `master` |
| Align/distribute shapes | `shapealign` |
| Add animations | `animation` |
| Set transitions | `transition` |
| Run slideshow | `slideshow` |
| Organize sections | `section` |
| Import slides | `slideimport` |
| Manage speaker notes | `notes` |
| Manage comments | `comment` |
| Check alt text/accessibility | `accessibility` |
| Spell check / language | `proofing` |
| Export presentation | `export` |
| Configure printing | `printoptions` |
| Script automation | `vba` |
| Control PowerPoint windows | `window` |
