# Sheets

A full-featured Google Sheets clone as an embeddable JavaScript library. Canvas-rendered, zero-dependency, 115KB minified.

## Quick Start

```html
<link rel="stylesheet" href="sheets.css">
<script src="sheets.js"></script>
<div id="spreadsheet" style="width:100%; height:600px;"></div>
<script>
  const ss = new Sheets.Spreadsheet('#spreadsheet');
</script>
```

## Features

### Grid & Rendering
- Canvas-based rendering with virtual scrolling (handles 100K+ rows)
- Row/column headers with resize (drag or double-click to auto-fit)
- Frozen rows and columns
- Cell merging
- Custom column widths and row heights
- Hidden rows/columns

### Editing
- Click to select, type to replace, Enter/F2 to edit in-place
- Formula bar editing with synchronized cell editor
- Multi-sheet support with tab bar (add, rename, duplicate, delete)
- Undo/redo (Ctrl+Z / Ctrl+Y)
- Copy/cut/paste with formatting
- Auto-fill (drag fill handle to extend patterns)
- Find & Replace (Ctrl+F / Ctrl+H)

### Formulas
- **142 built-in functions** across math, statistics, text, logical, date, lookup, financial, and trigonometry categories
- Full expression parser with operator precedence
- Cell references (A1, $A$1, A$1, $A1) with F4 cycling
- Range references (A1:B10)
- Cross-sheet references (Sheet1!A1)
- Dependency tracking and automatic recalculation
- **Formula autocomplete** — dropdown with function names and descriptions as you type
- **Signature hints** — parameter tooltip showing current argument
- **Point-and-click cell selection** — click or arrow-key to insert cell references while editing formulas

### Formatting
- Bold, italic, underline, strikethrough
- Font family and size
- Text and background colors
- Text alignment (left, center, right) and vertical alignment
- Text wrapping
- Number formats (general, number, currency, percent, date, time, scientific, accounting)
- Cell borders
- Conditional formatting

### Keyboard Shortcuts (Google Sheets compatible)
| Action | Shortcut |
|--------|----------|
| Edit cell | Enter / F2 |
| Cancel edit | Escape |
| Commit & move down/up | Enter / Shift+Enter |
| Commit & move right/left | Tab / Shift+Tab |
| Navigate | Arrow keys |
| Jump to edge | Ctrl+Arrow |
| Select range | Shift+Arrow |
| Select all | Ctrl+A |
| Bold / Italic / Underline | Ctrl+B / I / U |
| Copy / Cut / Paste | Ctrl+C / X / V |
| Undo / Redo | Ctrl+Z / Ctrl+Y |
| Find / Replace | Ctrl+F / Ctrl+H |
| Cycle absolute reference | F4 (in formula) |
| Insert cell reference | Arrow keys / Click (in formula) |
| Delete cell content | Delete / Backspace |

### Import / Export
- CSV and TSV (built-in)
- XLSX via SheetJS integration
- JSON serialization (full state save/restore)

## Embedding API

Every UI action has a corresponding API method. The library exposes rich state for host app integration.

### Initialization

```js
const ss = new Sheets.Spreadsheet('#container', {
  rows: 1000,           // default row count
  cols: 26,             // default column count (A-Z)
  toolbar: true,        // show formatting toolbar
  formulaBar: true,     // show formula bar
  sheetTabs: true,      // show sheet tab bar
  contextMenu: true,    // enable right-click menu
  readOnly: false,      // disable editing
  data: null,           // initial JSON data (from toJSON())
});
```

### Cell Operations

```js
// Read / write cell values
ss.setCellValue('A1', 'Hello');
ss.setCellValue('B1', 42);
ss.getCellValue('A1');              // 'Hello'

// Formulas
ss.setCellFormula('C1', '=A1&" "&B1');

// Bulk data
ss.setData([
  ['Name', 'Age', 'Score'],
  ['Alice', 30, 95],
  ['Bob', 25, 87],
]);

// Read data
ss.getData();                       // 2D array of all used cells
```

### Styling

```js
ss.setStyle({ bold: true, fontSize: 14, textColor: '#ff0000' });
ss.toggleBold();
ss.toggleItalic();
ss.toggleUnderline();
ss.toggleStrikethrough();
ss.toggleWrap();
ss.setAllBorders({ color: '#000', width: 1 });
ss.clearBorders();
ss.clearFormatting();
ss.adjustDecimals(1);               // increase decimal places
ss.adjustDecimals(-1);              // decrease decimal places
```

### Selection & Navigation

```js
// Read selection state
ss.selection;                       // CellRange object
ss.activeRow;                       // current row index
ss.activeCol;                       // current column index

// Programmatic selection
ss.selectionManager.select(0, 0);           // select A1
ss.selectionManager.selectRange(range);     // select a CellRange
ss.selectionManager.selectRow(5);           // select entire row 6
ss.selectionManager.selectCol(2);           // select entire column C
ss.selectionManager.selectAll();            // select all
ss.selectionManager.move(1, 0);             // move down
ss.selectionManager.moveToEdge(0, 1);       // Ctrl+Right behavior
```

### Sheet Management

```js
ss.addSheet('Sheet2');                      // add new sheet
ss.deleteSheet(sheetId);                    // delete sheet
ss.duplicateSheet(sheetId);                 // duplicate sheet
ss.setActiveSheet(sheetId);                 // switch to sheet
ss.sheets;                                  // array of all Sheet objects
ss.activeSheet;                             // current Sheet

// Sheet properties
const sheet = ss.activeSheet;
sheet.name;                                 // 'Sheet1'
sheet.rowCount;                             // 1000
sheet.colCount;                             // 26
sheet.frozenRows;                           // 0
sheet.frozenCols;                           // 0
sheet.getUsedRange();                       // CellRange of non-empty area
```

### Row / Column Operations

```js
ss.insertRowAbove();
ss.insertRowBelow();
ss.insertColLeft();
ss.insertColRight();
ss.deleteSelectedRows();
ss.deleteSelectedCols();
ss.freezeRows(2);                           // freeze top 2 rows
ss.freezeCols(1);                           // freeze first column

// Direct sheet manipulation
sheet.setColWidth(0, 150);                  // set column A width
sheet.setRowHeight(0, 40);                  // set row 1 height
sheet.insertRows(5, 3);                     // insert 3 rows at row 5
sheet.deleteCols(2, 1);                     // delete column C
```

### Clipboard & History

```js
ss.copy();
ss.cut();
ss.paste();
ss.undo();
ss.redo();
ss.deleteSelection();                       // clear selected cells
```

### Sorting & Merging

```js
ss.sortAscending();
ss.sortDescending();
ss.mergeSelection();
ss.unmergeSelection();
```

### Find & Replace

```js
ss.showFindReplace();                       // open find dialog
ss.showFindReplace(true);                   // open find & replace dialog
```

### Import / Export

```js
// CSV
ss.importCSV(csvString, ',');
const csv = ss.exportCSV(',');

// JSON (full state)
const state = ss.toJSON();
ss.fromJSON(state);
```

### Events

```js
ss.on('selectionChanged', (range) => { ... });
ss.on('sheetAdded', (sheet) => { ... });
ss.on('save', () => { ... });               // Ctrl+S
```

### Conditional Formatting

```js
ss.addConditionalFormat('A1:A100', 'greaterThan', {
  value: 50,
  style: { bgColor: '#c6efce', textColor: '#006100' },
});
```

### Cleanup

```js
ss.destroy();                               // remove from DOM, release resources
```

## Architecture

```
src/
  core/           Spreadsheet, Sheet, Cell, CellRange (data model)
  formula/        Tokenizer → Parser → AST → Evaluator, 142 functions
  render/         Canvas renderer with virtual scrolling
  interaction/    Keyboard, mouse, clipboard, selection handlers
  ui/             Editor, toolbar, formula bar, sheet tabs, context menu
  features/       Undo/redo, merge, sort, conditional formatting
  format/         Number formatting, cell styles
  io/             CSV parser, JSON serialization
  utils/          EventBus, helpers, constants
```

## Building

```bash
npm install
npm run build        # outputs dist/sheets.js and dist/sheets.esm.js
npm run dev          # watch mode
```

## License

MIT
