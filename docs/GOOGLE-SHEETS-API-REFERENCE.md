# Google Sheets API Design Reference

Reference for designing the sheets-clone public API, based on Google Apps Script SpreadsheetApp and Google Sheets REST API v4.

## 1. Class Hierarchy

```
SpreadsheetApp (static entry point)
  -> Spreadsheet (a file/workbook)
       -> Sheet (a single tab)
            -> Range (a rectangular cell selection)
                 -> RangeList (non-contiguous multi-range)
```

Each level returns the next level down. All mutating methods return `this` for chaining.

## 2. SpreadsheetApp (Static Entry Point)

### Active State
```
getActive() -> Spreadsheet
getActiveSheet() -> Sheet
getActiveRange() -> Range
getCurrentCell() -> Range|null
getSelection() -> Selection
setActiveRange(range) -> Range
setCurrentCell(cell) -> Range
```

### Builders
```
newConditionalFormatRule() -> ConditionalFormatRuleBuilder
newDataValidation() -> DataValidationBuilder
newFilterCriteria() -> FilterCriteriaBuilder
newTextStyle() -> TextStyleBuilder
newColor() -> ColorBuilder
```

### Key Enums
```
BorderStyle     { DOTTED, DASHED, SOLID, SOLID_MEDIUM, SOLID_THICK, DOUBLE }
CopyPasteType   { PASTE_NORMAL, PASTE_NO_BORDERS, PASTE_FORMAT, PASTE_FORMULA, PASTE_VALUES }
Direction       { UP, DOWN, NEXT, PREVIOUS }
WrapStrategy    { WRAP, OVERFLOW, CLIP }
ValueType       { EMPTY, STRING, NUMBER, BOOLEAN, DATE, ERROR }
```

## 3. Spreadsheet (Workbook)

### Sheet Management
```
getSheets() -> Sheet[]
getSheetByName(name) -> Sheet|null
getNumSheets() -> Integer
insertSheet(name?) -> Sheet
deleteSheet(sheet) -> void
getActiveSheet() -> Sheet
setActiveSheet(sheet) -> Sheet
duplicateActiveSheet() -> Sheet
moveActiveSheet(pos) -> void
```

### Range Access (delegates to active sheet)
```
getRange(a1Notation) -> Range
getRangeByName(name) -> Range|null
getDataRange() -> Range
getActiveCell() -> Range
getActiveRange() -> Range
```

### Named Ranges
```
setNamedRange(name, range) -> void
getNamedRanges() -> NamedRange[]
getRangeByName(name) -> Range|null
removeNamedRange(name) -> void
```

### Document Metadata
```
getName() -> String
rename(newName) -> void
getLastRow() -> Integer
getLastColumn() -> Integer
```

## 4. Sheet (Tab)

### Identity
```
getName() -> String
setName(name) -> Sheet
getSheetId() -> Integer
getIndex() -> Integer
activate() -> Sheet
copyTo(spreadsheet) -> Sheet
```

### Range Access (core gateway)
```
getRange(row, column) -> Range                    // 1-based
getRange(row, column, numRows, numColumns) -> Range
getRange(a1Notation) -> Range                     // "A1:B5"
getDataRange() -> Range                           // all data
getActiveCell() -> Range
getActiveRange() -> Range
getSelection() -> Selection
setActiveRange(range) -> Range
setCurrentCell(cell) -> Range
```

### Data Dimensions
```
getLastRow() -> Integer
getLastColumn() -> Integer
getMaxRows() -> Integer
getMaxColumns() -> Integer
```

### Row Operations
```
insertRow(rowIndex) -> void
insertRows(rowIndex, numRows) -> void
insertRowAfter(afterPosition) -> Sheet
insertRowBefore(beforePosition) -> Sheet
deleteRow(rowPosition) -> Sheet
deleteRows(rowPosition, howMany) -> void
appendRow(rowContents[]) -> Sheet
moveRows(rowSpec, destinationIndex) -> void
getRowHeight(rowPosition) -> Integer
setRowHeight(rowPosition, height) -> Sheet
setRowHeights(startRow, numRows, height) -> Sheet
autoResizeRows(startRow, numRows) -> Sheet
hideRows(rowIndex, numRows?) -> void
showRows(rowIndex, numRows?) -> void
isRowHiddenByUser(rowPosition) -> Boolean
isRowHiddenByFilter(rowPosition) -> Boolean
```

### Column Operations
```
insertColumn(columnIndex) -> void
insertColumns(columnIndex, numColumns) -> void
deleteColumn(columnPosition) -> Sheet
deleteColumns(columnPosition, howMany) -> void
getColumnWidth(columnPosition) -> Integer
setColumnWidth(columnPosition, width) -> Sheet
autoResizeColumn(columnPosition) -> Sheet
autoResizeColumns(startColumn, numColumns) -> Sheet
hideColumns(columnIndex, numColumns?) -> void
showColumns(columnIndex, numColumns?) -> void
isColumnHiddenByUser(columnPosition) -> Boolean
```

### Freezing
```
setFrozenRows(rows) -> void
setFrozenColumns(columns) -> void
getFrozenRows() -> Integer
getFrozenColumns() -> Integer
```

### Clearing
```
clear() -> Sheet
clear({formatOnly, contentsOnly}) -> Sheet
clearContents() -> Sheet
clearFormats() -> Sheet
clearNotes() -> Sheet
clearConditionalFormatRules() -> void
```

### Sorting
```
sort(columnPosition) -> Sheet
sort(columnPosition, ascending) -> Sheet
```

### Visibility
```
hideSheet() -> Sheet
showSheet() -> Sheet
isSheetHidden() -> Boolean
setTabColor(color) -> Sheet
```

### Conditional Formatting
```
getConditionalFormatRules() -> ConditionalFormatRule[]
setConditionalFormatRules(rules[]) -> void
```

### Filtering
```
getFilter() -> Filter|null
```

### Search
```
createTextFinder(findText) -> TextFinder
```

## 5. Range (Most Important Class)

### KEY PATTERN: Singular/Plural Pairs
```
getValue()  -> Object        // top-left cell
getValues() -> Object[][]   // 2D array [row][col]
setValue(value) -> Range     // sets all cells, returns this
setValues(values[][]) -> Range
```

### Cell Values & Formulas
```
getValue() / getValues()
setValue(value) / setValues(values[][])
getDisplayValue() / getDisplayValues()       // formatted strings
getFormula() / getFormulas()
setFormula(formula) / setFormulas(formulas[][])
```

### Background Color
```
getBackground() / getBackgrounds()            // "#ffffff"
setBackground(color) / setBackgrounds(colors[][])
setBackgroundRGB(red, green, blue)
```

### Font Properties
```
getFontColor() / setFontColor(color)
getFontFamily() / setFontFamily(fontFamily)
getFontSize() / setFontSize(size)
getFontStyle() / setFontStyle(style)          // "normal" | "italic"
getFontWeight() / setFontWeight(weight)       // "normal" | "bold"
getFontLine() / setFontLine(line)             // "none" | "underline" | "line-through"
// All have plural variants: getFontColors(), setFontSizes(sizes[][]), etc.
```

### Alignment
```
getHorizontalAlignment() / setHorizontalAlignment(alignment)   // "left"|"center"|"right"
getVerticalAlignment() / setVerticalAlignment(alignment)       // "top"|"middle"|"bottom"
// Plural variants exist
```

### Wrapping & Number Format
```
getWrap() / setWrap(wrap)
getWrapStrategy() / setWrapStrategy(strategy)    // WRAP|OVERFLOW|CLIP
getNumberFormat() / setNumberFormat(format)       // "#,##0.00", "$#,##0"
// Plural variants exist
```

### Borders
```
setBorder(top, left, bottom, right, vertical, horizontal) -> Range
setBorder(top, left, bottom, right, vertical, horizontal, color, style) -> Range
// top/left/bottom/right/vertical/horizontal are Boolean|null
```

### Range Metadata
```
getA1Notation() -> String         // "A1:B5"
getRow() / getColumn()            // 1-based
getLastRow() / getLastColumn()
getNumRows() / getNumColumns()
getSheet() -> Sheet
```

### Cell Navigation
```
getCell(row, column) -> Range     // relative within range
offset(rowOffset, columnOffset, numRows?, numColumns?) -> Range
getNextDataCell(direction) -> Range    // Ctrl+arrow behavior
getDataRegion() -> Range              // expand to contiguous data
```

### Merging
```
merge() -> Range
mergeAcross() -> Range
mergeVertically() -> Range
getMergedRanges() -> Range[]
isPartOfMerge() -> Boolean
breakApart() -> Range
```

### Notes
```
getNote() / getNotes()
setNote(note) / setNotes(notes[][])
clearNote()
```

### Data Validation
```
getDataValidation() / getDataValidations()
setDataValidation(rule) / setDataValidations(rules[][])
clearDataValidations()
```

### Clearing
```
clear() / clear({contentsOnly, formatOnly})
clearContent()
clearFormat()
```

### Copy/Move
```
copyTo(destination) -> void
copyTo(destination, copyPasteType, transposed) -> void
moveTo(target) -> void
```

### Checkboxes
```
insertCheckboxes() -> Range
removeCheckboxes() -> Range
check() / uncheck()
isChecked() -> Boolean|null
```

### Filtering & Sorting
```
createFilter() -> Filter
sort(sortSpecObj) -> Range
randomize() -> Range
removeDuplicates(columnsToCompare[]?) -> Range
```

### Banding (alternating colors)
```
applyRowBanding(theme?, showHeader?, showFooter?) -> Banding
applyColumnBanding(theme?, showHeader?, showFooter?) -> Banding
```

## 6. Event Model

### onEdit(e)
```js
e.range      // Range (edited cells)
e.source     // Spreadsheet
e.value      // new value (single cell only)
e.oldValue   // previous value (single cell only)
```

### onSelectionChange(e)
```js
e.range      // Range (new selection)
e.source     // Spreadsheet
```

### onChange(e)
```js
e.changeType // "EDIT"|"INSERT_ROW"|"INSERT_COLUMN"|"REMOVE_ROW"|"REMOVE_COLUMN"|"FORMAT"|"OTHER"
e.source     // Spreadsheet
```

## 7. REST API v4 Patterns

### Values Operations
| Operation | Method | Pattern |
|-----------|--------|---------|
| Read range | GET | `/values/{range}` |
| Read multiple | GET | `/values:batchGet?ranges=A1:B2&ranges=C1:D5` |
| Write range | PUT | `/values/{range}?valueInputOption=USER_ENTERED` |
| Write multiple | POST | `/values:batchUpdate` |
| Append | POST | `/values/{range}:append` |

### ValueInputOption
- `RAW` — literal string, formulas stored as text
- `USER_ENTERED` — parsed like typing in UI (formulas evaluated, dates parsed)

### A1 Notation
```
"A1"                 Single cell
"A1:B5"              Rectangle
"A:A"                Entire column
"1:1"                Entire row
"Sheet1!A1:B5"       Sheet-qualified
"'My Sheet'!A1"      Quoted for special chars
```

## 8. Key Design Patterns

1. **Method chaining** — All setters return `this`
2. **Singular/plural pairs** — `getValue()`/`getValues()` for every property
3. **Builder pattern** — Complex objects (DataValidation, ConditionalFormat, TextStyle) use builders
4. **Active state model** — Always one active sheet, range, and cell
5. **A1 notation** — Universal range addressing alongside numeric (row, col, numRows, numCols)
6. **Overloaded getRange** — 4 signatures, the most-used method in the API
7. **Clear separation** — Values vs formatting vs structure are distinct operations
8. **Event objects are plain data** — `{range, source, value, oldValue}`
