import Cell, { CellStyle } from './Cell.js';
import CellRange from './CellRange.js';
import Range from '../api/Range.js';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, DEFAULT_ROWS, DEFAULT_COLS, CELL_TYPE } from '../utils/constants.js';
import { cellKey, keyToRC, generateId, indexToCol, parseCellRef } from '../utils/helpers.js';

export default class Sheet {
  constructor(name, opts = {}) {
    this.id = generateId();
    this.name = name || 'Sheet1';
    this.cells = new Map();
    this.rowCount = opts.rows || DEFAULT_ROWS;
    this.colCount = opts.cols || DEFAULT_COLS;
    this.colWidths = new Map();
    this.rowHeights = new Map();
    this.hiddenRows = new Set();
    this.hiddenCols = new Set();
    this.frozenRows = 0;
    this.frozenCols = 0;
    this.merges = [];
    this.conditionalFormats = [];
    this.dataValidations = [];
    this.filterRange = null;
    this.filterCriteria = new Map();
    this.filteredRows = new Set();
    this.defaultColWidth = opts.defaultColWidth || DEFAULT_COL_WIDTH;
    this.defaultRowHeight = opts.defaultRowHeight || DEFAULT_ROW_HEIGHT;
  }

  getCell(row, col) {
    return this.cells.get(cellKey(row, col)) || null;
  }

  getOrCreateCell(row, col) {
    const key = cellKey(row, col);
    let cell = this.cells.get(key);
    if (!cell) {
      cell = new Cell();
      this.cells.set(key, cell);
    }
    return cell;
  }

  setCell(row, col, cell) {
    if (cell.isEmpty && !cell.style && !cell.mergeParent && !cell.mergeSpan) {
      this.cells.delete(cellKey(row, col));
    } else {
      this.cells.set(cellKey(row, col), cell);
    }
  }

  getCellValue(row, col) {
    const cell = this.getCell(row, col);
    if (!cell) return null;
    return cell.computedValue !== null ? cell.computedValue : cell.rawValue;
  }

  setCellValue(row, col, value) {
    const cell = this.getOrCreateCell(row, col);
    cell.setValue(value);
    if (cell.isEmpty && !cell.style) {
      this.cells.delete(cellKey(row, col));
    }
    return cell;
  }

  setCellFormula(row, col, formula) {
    const cell = this.getOrCreateCell(row, col);
    cell.setFormula(formula);
    return cell;
  }

  setCellStyle(row, col, props) {
    const cell = this.getOrCreateCell(row, col);
    cell.setStyle(props);
    return cell;
  }

  clearCell(row, col) {
    this.cells.delete(cellKey(row, col));
  }

  clearRange(range) {
    range.forEach((r, c) => this.clearCell(r, c));
  }

  clearRangeContent(range) {
    range.forEach((r, c) => {
      const cell = this.getCell(r, c);
      if (cell) {
        cell.clearContent();
        if (cell.isEmpty && !cell.style) {
          this.cells.delete(cellKey(r, c));
        }
      }
    });
  }

  clearRangeFormat(range) {
    range.forEach((r, c) => {
      const cell = this.getCell(r, c);
      if (cell) {
        cell.clearFormat();
        if (cell.isEmpty && !cell.style) {
          this.cells.delete(cellKey(r, c));
        }
      }
    });
  }

  getRangeValues(range) {
    return range.map((r, c) => this.getCellValue(r, c));
  }

  setRangeValues(startRow, startCol, values) {
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        this.setCellValue(startRow + r, startCol + c, values[r][c]);
      }
    }
  }

  setRangeStyle(range, props) {
    range.forEach((r, c) => this.setCellStyle(r, c, props));
  }

  getColWidth(col) {
    if (this.hiddenCols.has(col)) return 0;
    return this.colWidths.get(col) || this.defaultColWidth;
  }

  setColWidth(col, width) {
    this.colWidths.set(col, Math.max(20, width));
  }

  getRowHeight(row) {
    if (this.hiddenRows.has(row)) return 0;
    return this.rowHeights.get(row) || this.defaultRowHeight;
  }

  setRowHeight(row, height) {
    this.rowHeights.set(row, Math.max(5, height));
  }

  isRowVisible(row) {
    return !this.hiddenRows.has(row) && !this.filteredRows.has(row);
  }

  isColVisible(col) {
    return !this.hiddenCols.has(col);
  }

  getColX(col) {
    let x = 0;
    for (let c = 0; c < col; c++) {
      if (this.isColVisible(c)) x += this.getColWidth(c);
    }
    return x;
  }

  getRowY(row) {
    let y = 0;
    for (let r = 0; r < row; r++) {
      if (this.isRowVisible(r)) y += this.getRowHeight(r);
    }
    return y;
  }

  getColAtX(x) {
    let cx = 0;
    for (let c = 0; c < this.colCount; c++) {
      if (!this.isColVisible(c)) continue;
      const w = this.getColWidth(c);
      if (x < cx + w) return c;
      cx += w;
    }
    return this.colCount - 1;
  }

  getRowAtY(y) {
    let cy = 0;
    for (let r = 0; r < this.rowCount; r++) {
      if (!this.isRowVisible(r)) continue;
      const h = this.getRowHeight(r);
      if (y < cy + h) return r;
      cy += h;
    }
    return this.rowCount - 1;
  }

  getTotalWidth() {
    let w = 0;
    for (let c = 0; c < this.colCount; c++) {
      if (this.isColVisible(c)) w += this.getColWidth(c);
    }
    return w;
  }

  getTotalHeight() {
    let h = 0;
    for (let r = 0; r < this.rowCount; r++) {
      if (this.isRowVisible(r)) h += this.getRowHeight(r);
    }
    return h;
  }

  insertRows(at, count = 1) {
    const newCells = new Map();
    for (const [key, cell] of this.cells) {
      const { row, col } = keyToRC(key);
      const newRow = row >= at ? row + count : row;
      newCells.set(cellKey(newRow, col), cell);
    }
    this.cells = newCells;

    const newRowHeights = new Map();
    for (const [row, h] of this.rowHeights) {
      newRowHeights.set(row >= at ? row + count : row, h);
    }
    this.rowHeights = newRowHeights;

    const newHidden = new Set();
    for (const row of this.hiddenRows) {
      newHidden.add(row >= at ? row + count : row);
    }
    this.hiddenRows = newHidden;

    this.rowCount += count;
    this._shiftMerges(at, count, true);
  }

  deleteRows(at, count = 1) {
    const newCells = new Map();
    for (const [key, cell] of this.cells) {
      const { row, col } = keyToRC(key);
      if (row >= at && row < at + count) continue;
      const newRow = row >= at + count ? row - count : row;
      newCells.set(cellKey(newRow, col), cell);
    }
    this.cells = newCells;

    const newRowHeights = new Map();
    for (const [row, h] of this.rowHeights) {
      if (row >= at && row < at + count) continue;
      newRowHeights.set(row >= at + count ? row - count : row, h);
    }
    this.rowHeights = newRowHeights;

    this.rowCount -= count;
    this._shiftMerges(at, -count, true);
  }

  insertCols(at, count = 1) {
    const newCells = new Map();
    for (const [key, cell] of this.cells) {
      const { row, col } = keyToRC(key);
      const newCol = col >= at ? col + count : col;
      newCells.set(cellKey(row, newCol), cell);
    }
    this.cells = newCells;

    const newColWidths = new Map();
    for (const [col, w] of this.colWidths) {
      newColWidths.set(col >= at ? col + count : col, w);
    }
    this.colWidths = newColWidths;

    const newHidden = new Set();
    for (const col of this.hiddenCols) {
      newHidden.add(col >= at ? col + count : col);
    }
    this.hiddenCols = newHidden;

    this.colCount += count;
    this._shiftMerges(at, count, false);
  }

  deleteCols(at, count = 1) {
    const newCells = new Map();
    for (const [key, cell] of this.cells) {
      const { row, col } = keyToRC(key);
      if (col >= at && col < at + count) continue;
      const newCol = col >= at + count ? col - count : col;
      newCells.set(cellKey(row, newCol), cell);
    }
    this.cells = newCells;

    const newColWidths = new Map();
    for (const [col, w] of this.colWidths) {
      if (col >= at && col < at + count) continue;
      newColWidths.set(col >= at + count ? col - count : col, w);
    }
    this.colWidths = newColWidths;

    this.colCount -= count;
    this._shiftMerges(at, -count, false);
  }

  _shiftMerges(at, delta, isRow) {
    this.merges = this.merges.map(m => {
      const range = CellRange.fromString(m);
      if (!range) return m;
      let { startRow, startCol, endRow, endCol } = range;
      if (isRow) {
        if (startRow >= at) startRow += delta;
        if (endRow >= at) endRow += delta;
      } else {
        if (startCol >= at) startCol += delta;
        if (endCol >= at) endCol += delta;
      }
      return new CellRange(startRow, startCol, endRow, endCol).toString();
    }).filter(m => {
      const range = CellRange.fromString(m);
      return range && range.startRow >= 0 && range.startCol >= 0;
    });
  }

  addMerge(range) {
    const key = range.toString();
    if (!this.merges.includes(key)) {
      this.merges.push(key);
      const topLeft = this.getOrCreateCell(range.startRow, range.startCol);
      topLeft.mergeSpan = { rows: range.rowCount, cols: range.colCount };
      range.forEach((r, c) => {
        if (r === range.startRow && c === range.startCol) return;
        const cell = this.getOrCreateCell(r, c);
        cell.mergeParent = { row: range.startRow, col: range.startCol };
      });
    }
  }

  removeMerge(range) {
    const key = range.toString();
    const idx = this.merges.indexOf(key);
    if (idx !== -1) {
      this.merges.splice(idx, 1);
      range.forEach((r, c) => {
        const cell = this.getCell(r, c);
        if (cell) {
          cell.mergeParent = null;
          cell.mergeSpan = null;
        }
      });
    }
  }

  getMergeAt(row, col) {
    const cell = this.getCell(row, col);
    if (!cell) return null;
    if (cell.mergeSpan) {
      return new CellRange(row, col, row + cell.mergeSpan.rows - 1, col + cell.mergeSpan.cols - 1);
    }
    if (cell.mergeParent) {
      const parent = this.getCell(cell.mergeParent.row, cell.mergeParent.col);
      if (parent && parent.mergeSpan) {
        return new CellRange(
          cell.mergeParent.row, cell.mergeParent.col,
          cell.mergeParent.row + parent.mergeSpan.rows - 1,
          cell.mergeParent.col + parent.mergeSpan.cols - 1,
        );
      }
    }
    return null;
  }

  getUsedRange() {
    let minR = Infinity, minC = Infinity, maxR = -Infinity, maxC = -Infinity;
    for (const key of this.cells.keys()) {
      const { row, col } = keyToRC(key);
      if (row < minR) minR = row;
      if (col < minC) minC = col;
      if (row > maxR) maxR = row;
      if (col > maxC) maxC = col;
    }
    if (minR === Infinity) return null;
    return new CellRange(minR, minC, maxR, maxC);
  }

  // ── Google Sheets-compatible API ──

  /**
   * getRange — supports 4 signatures (matching Google Apps Script):
   *   getRange('A1')              A1 notation, single cell
   *   getRange('A1:B5')           A1 notation, range
   *   getRange(row, col)          1-based row/col, single cell
   *   getRange(row, col, numRows, numCols)  1-based, rectangle
   */
  getRange(rowOrA1, col, numRows, numCols) {
    if (typeof rowOrA1 === 'string') {
      return Range.fromA1Notation(this, rowOrA1);
    }
    // 1-based to 0-based
    const r = rowOrA1 - 1;
    const c = (col || 1) - 1;
    return new Range(this, r, c, numRows || 1, numCols || 1);
  }

  getDataRange() {
    const used = this.getUsedRange();
    if (!used) return new Range(this, 0, 0, 1, 1);
    return new Range(this, used.startRow, used.startCol,
      used.endRow - used.startRow + 1, used.endCol - used.startCol + 1);
  }

  getLastRow() {
    let maxR = 0;
    for (const key of this.cells.keys()) {
      const r = key >> 16;
      if (r + 1 > maxR) maxR = r + 1;
    }
    return maxR; // 1-based
  }

  getLastColumn() {
    let maxC = 0;
    for (const key of this.cells.keys()) {
      const c = key & 0xffff;
      if (c + 1 > maxC) maxC = c + 1;
    }
    return maxC; // 1-based
  }

  getMaxRows() { return this.rowCount; }
  getMaxColumns() { return this.colCount; }

  appendRow(values) {
    const r = this.getLastRow(); // next empty row (0-based = r, since getLastRow is 1-based)
    for (let c = 0; c < values.length; c++) {
      this.setCellValue(r, c, values[c]);
    }
    return this;
  }

  getSheetId() { return this.id; }
  getSheetName() { return this.name; }
  getName() { return this.name; }
  setName(name) { this.name = name; return this; }

  getFrozenRows() { return this.frozenRows; }
  getFrozenColumns() { return this.frozenCols; }
  setFrozenRows(n) { this.frozenRows = n; }
  setFrozenColumns(n) { this.frozenCols = n; }

  sort(columnPosition, ascending = true) {
    const dataRange = this.getDataRange();
    dataRange.sort({ column: typeof columnPosition === 'number' ? columnPosition : 1, ascending });
    return this;
  }

  clearContents() {
    for (const [key, cell] of this.cells) { cell.clearContent(); }
    return this;
  }

  clearFormats() {
    for (const [key, cell] of this.cells) { cell.clearFormat(); }
    return this;
  }

  autoFitColWidth(col, ctx) {
    let maxW = 40;
    for (let r = 0; r < Math.min(this.rowCount, 500); r++) {
      const cell = this.getCell(r, col);
      if (!cell || cell.isEmpty) continue;
      const style = cell.getStyle();
      ctx.font = style.getFont();
      const w = ctx.measureText(cell.displayValue).width + 16;
      if (w > maxW) maxW = w;
    }
    this.setColWidth(col, Math.min(maxW, 500));
  }

  toJSON() {
    const cellData = {};
    for (const [key, cell] of this.cells) {
      const { row, col } = keyToRC(key);
      const ref = indexToCol(col) + (row + 1);
      const json = cell.toJSON();
      if (Object.keys(json).length > 0) {
        cellData[ref] = json;
      }
    }
    return {
      name: this.name,
      rowCount: this.rowCount,
      colCount: this.colCount,
      cells: cellData,
      colWidths: Object.fromEntries(this.colWidths),
      rowHeights: Object.fromEntries(this.rowHeights),
      frozenRows: this.frozenRows,
      frozenCols: this.frozenCols,
      merges: this.merges,
      hiddenRows: [...this.hiddenRows],
      hiddenCols: [...this.hiddenCols],
    };
  }

  static fromJSON(data) {
    const sheet = new Sheet(data.name, { rows: data.rowCount, cols: data.colCount });
    if (data.cells) {
      for (const [ref, cellData] of Object.entries(data.cells)) {
        const parsed = parseCellRef(ref);
        if (parsed) {
          sheet.cells.set(cellKey(parsed.row, parsed.col), Cell.fromJSON(cellData));
        }
      }
    }
    if (data.colWidths) {
      for (const [col, w] of Object.entries(data.colWidths)) {
        sheet.colWidths.set(Number(col), w);
      }
    }
    if (data.rowHeights) {
      for (const [row, h] of Object.entries(data.rowHeights)) {
        sheet.rowHeights.set(Number(row), h);
      }
    }
    sheet.frozenRows = data.frozenRows || 0;
    sheet.frozenCols = data.frozenCols || 0;
    sheet.merges = data.merges || [];
    sheet.hiddenRows = new Set(data.hiddenRows || []);
    sheet.hiddenCols = new Set(data.hiddenCols || []);
    return sheet;
  }
}
