# OnlyOffice Spreadsheet Analysis

What to learn from and what to avoid.

## Architecture

- **Canvas 2D rendering** (same as us) — no DOM cells, no WebAssembly
- **Formula engine**: Pure JS, recursive descent parser, ~1.9MB across category files
- **Main repo**: `ONLYOFFICE/sdkjs` (324 stars) — cell/, common/, word/, slide/
- **UI shell**: `ONLYOFFICE/web-apps` — Backbone.js toolbar/ribbon/panels
- **Collaboration**: WebSocket + cell-level locking (not OT)
- **File format**: C++ core for XLSX (780MB!) + JS serialization layer

## What to Steal

### 1. SheetMemory Pattern (HIGH VALUE)
`SheetMemory.js` — uses `ArrayBuffer` with typed array views instead of JS objects for cell storage. Fixed-size structs indexed by row, lazy allocation. Much more memory-efficient than object-per-cell. **We should adopt this for large sheets.**

### 2. Number Format Engine
`NumFormat.js` (447KB) — handles all Excel number format codes. Our `NumberFormat.js` is basic. Their implementation is the reference for XLSX fidelity.

### 3. Rendering Pipeline Order
Grid lines → cell backgrounds (merge-aware) → cell text → sparklines → borders (separate pass) → frozen panes (re-rendered per quadrant) → selection. **Their frozen pane approach (re-draw per quadrant) is correct.**

### 4. Keyboard Shortcut System
Declarative: each shortcut is a data object `{type, keyCode, ctrl, shift, alt}` in a map. Separates definitions from handlers. Has "unlocked shortcuts" concept for protected mode.

### 5. Clipboard Binary Format
Internal copy/paste uses custom binary format (fast). Cross-app paste parses HTML tables. Special paste with math operations (add, subtract, multiply). **We should add paste-with-operation.**

### 6. Conditional Formatting: Data Bars + Icon Sets
`_drawCellCFDataBar()` and `_drawCellCFIconSet()` — visual indicators drawn directly in cells. We have basic CF but not data bars or icon sets.

### 7. Protected Ranges with Granular Permissions
Per-user range protection with SHA hashing. Permissions: format cells, insert/delete rows, sort, autofilter, pivot. **Critical for our agent+user use case.**

### 8. Sparklines
Line/column/win-loss sparklines drawn directly onto the cell canvas via the chart engine. Small but high-value feature.

## What NOT to Copy

### 1. God-Object Files
`WorksheetView.js` = 1,116KB (14K+ lines). `Workbook.js` = 909KB. Untestable monoliths. **We're already better here with modular files.**

### 2. No Code Splitting
Entire formula engine (1.9MB), chart system (2.5MB), model (4.4MB) loaded upfront. No lazy loading. A 5-cell sheet loads 15+MB. **This is why it's bulky.**

### 3. Pre-ES6 Everything
IIFEs, prototype chains, global namespaces. No tree-shaking possible. **We use ES modules — already better.**

### 4. Inline Data Blobs
`ChartStyles.js` = 1,547KB of style definitions inline in JS. SVG icons base64-encoded in code. **Keep data separate.**

### 5. Full Canvas Redraw
No dirty-rect tracking. Every change redraws all ~2000 visible cells. **We should add dirty-rect optimization.**

### 6. C++ File Format Dependency
780MB of C++ for XLSX. We use SheetJS (JS-only, ~500KB). **Much lighter.**

### 7. Single-Thread Formulas
All formula computation on main thread. **We should consider Web Workers.**

### 8. Monolithic Collaboration Server
RabbitMQ + Redis + PostgreSQL. Overkill. **Simple WebSocket relay or CRDTs better for us.**

## Priority Adoptions for Our Library

| Feature | OnlyOffice Source | Effort | Impact |
|---------|------------------|--------|--------|
| ArrayBuffer cell storage | SheetMemory.js | Large | Performance on large sheets |
| Data bars in CF | WorksheetView._drawCellCFDataBar | Medium | Visual polish |
| Icon sets in CF | WorksheetView._drawCellCFIconSet | Medium | Visual polish |
| Sparklines | _drawSparklines + chart engine | Medium | Feature parity |
| Protected ranges | protectRange.js | Medium | Agent+user security |
| Paste with operations | clipboard.js | Small | Power user feature |
| Dirty-rect rendering | N/A (they don't have it) | Medium | Performance |
| Web Worker formulas | N/A (they don't have it) | Medium | Performance |
