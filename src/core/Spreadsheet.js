import EventBus from '../utils/EventBus.js';
import { el, cellKey, keyToRC, indexToCol, deepClone, parseCellRef, cellRefToString } from '../utils/helpers.js';
import { DEFAULT_ROWS, DEFAULT_COLS, CELL_TYPE } from '../utils/constants.js';
import Sheet from './Sheet.js';
import Cell, { CellStyle } from './Cell.js';
import CellRange from './CellRange.js';
import FormulaEngine from '../formula/FormulaEngine.js';
import Renderer from '../render/Renderer.js';
import SelectionManager from '../interaction/SelectionManager.js';
import KeyboardHandler from '../interaction/KeyboardHandler.js';
import MouseHandler from '../interaction/MouseHandler.js';
import ClipboardHandler from '../interaction/ClipboardHandler.js';
import Editor from '../ui/Editor.js';
import FormulaHelper from '../ui/FormulaHelper.js';
import Toolbar from '../ui/Toolbar.js';
import FormulaBar from '../ui/FormulaBar.js';
import SheetTabs from '../ui/SheetTabs.js';
import ContextMenu from '../ui/ContextMenu.js';
import FindReplace from '../ui/FindReplace.js';
import StatusBar from '../ui/StatusBar.js';
import CommandManager from '../features/CommandManager.js';
import DataValidationManager from '../features/DataValidation.js';
import NamedRangesManager from '../features/NamedRanges.js';
import ConditionalFormatUI from '../features/ConditionalFormatUI.js';
import NumberFormat from '../format/NumberFormat.js';
import { parseCSV, generateCSV } from '../io/CSV.js';

export default class Spreadsheet {
  constructor(container, options = {}) {
    if (typeof container === 'string') {
      container = document.querySelector(container);
    }
    if (!container) throw new Error('Container element not found');

    this._events = new EventBus();
    this.options = {
      rows: options.rows || DEFAULT_ROWS,
      cols: options.cols || DEFAULT_COLS,
      toolbar: options.toolbar !== false,
      formulaBar: options.formulaBar !== false,
      sheetTabs: options.sheetTabs !== false,
      contextMenu: options.contextMenu !== false,
      readOnly: options.readOnly || false,
      data: options.data || null,
      ...options,
    };

    // Core state
    this.sheets = [];
    this.activeSheetIndex = 0;
    this.copyRange = null;

    // Managers
    this.formulaEngine = new FormulaEngine(this);
    this.numberFormat = new NumberFormat();
    this.commandManager = new CommandManager(this);
    this.selectionManager = new SelectionManager(this);
    this.renderer = new Renderer(this);
    this.keyboard = new KeyboardHandler(this);
    this.mouse = new MouseHandler(this);
    this.clipboard = new ClipboardHandler(this);
    this.editor = new Editor(this);
    this.formulaHelper = new FormulaHelper(this);
    this.toolbar = this.options.toolbar ? new Toolbar(this) : null;
    this.formulaBar = this.options.formulaBar ? new FormulaBar(this) : null;
    this.sheetTabs = this.options.sheetTabs ? new SheetTabs(this) : null;
    this.contextMenu = this.options.contextMenu ? new ContextMenu(this) : null;
    this.findReplace = new FindReplace(this);
    this.statusBar = new StatusBar(this);
    this.dataValidation = new DataValidationManager(this);
    this.namedRanges = new NamedRangesManager(this);
    this.conditionalFormatUI = new ConditionalFormatUI(this);

    // Build DOM first so UI components are initialized
    this._buildDOM(container);

    // Create initial sheet (after DOM so sheetTabs.update works)
    this.addSheet('Sheet1');

    // Load initial data
    if (this.options.data) {
      this.fromJSON(this.options.data);
    }

    // Initial render
    this.recalculate();
    this.render();
  }

  // ── DOM ──

  _buildDOM(parentContainer) {
    this.container = el('div', {
      className: 'sheets-container',
      tabIndex: '0',
    });
    this.container.style.width = '100%';
    this.container.style.height = '100%';
    this.container.style.position = 'relative';
    this.container.style.outline = 'none';
    this.container.style.overflow = 'hidden';
    this.container.style.display = 'flex';
    this.container.style.flexDirection = 'column';
    this.container.style.fontFamily = 'Arial, sans-serif';

    // Toolbar
    if (this.toolbar) {
      this.toolbar.init(this.container);
    }

    // Formula bar
    if (this.formulaBar) {
      this.formulaBar.init(this.container);
    }

    // Grid area
    this.gridArea = el('div', { className: 'sheets-grid-area' });
    this.gridArea.style.flex = '1';
    this.gridArea.style.position = 'relative';
    this.gridArea.style.overflow = 'hidden';

    this.canvas = el('canvas', { className: 'sheets-canvas' });
    this.canvas.style.position = 'absolute';
    this.canvas.style.top = '0';
    this.canvas.style.left = '0';

    this.gridArea.appendChild(this.canvas);
    this.container.appendChild(this.gridArea);

    // Init renderer on canvas
    this.renderer.init(this.canvas);

    // Init editor and formula helper (textarea overlay)
    this.editor.init(this.gridArea);
    this.formulaHelper.init(this.gridArea);

    // Init interaction handlers
    this.keyboard.init(this.container);
    this.mouse.init(this.canvas);
    this.clipboard.init(this.container);
    this.findReplace.init(this.gridArea);

    // Sheet tabs + status bar
    if (this.sheetTabs) {
      this.sheetTabs.init(this.container);
    }
    this.statusBar.init(this.container);

    parentContainer.appendChild(this.container);

    // Resize observer
    this._resizeObserver = new ResizeObserver(() => {
      this.renderer.resize();
    });
    this._resizeObserver.observe(this.gridArea);

    this.container.focus();
  }

  // ── Google Sheets-compatible API ──

  /** Get a sheet by name (Apps Script: getSheetByName) */
  sheet(name) {
    return this.sheets.find(s => s.name === name) || null;
  }

  /** Get active range as a Range object */
  getActiveRange() {
    const sel = this.selection;
    if (!sel || !this.activeSheet) return null;
    return this.activeSheet.getRange(sel.startRow + 1, sel.startCol + 1,
      sel.endRow - sel.startRow + 1, sel.endCol - sel.startCol + 1);
  }

  /** Get active cell as a Range object */
  getActiveCell() {
    if (!this.activeSheet) return null;
    return this.activeSheet.getRange(this.activeRow + 1, this.activeCol + 1);
  }

  /** Get range on active sheet using A1 notation */
  getRange(a1Notation) {
    if (!this.activeSheet) return null;
    return this.activeSheet.getRange(a1Notation);
  }

  /** Navigate to a cell — switches sheet + scrolls + selects */
  navigateTo(sheetName, a1Ref) {
    if (sheetName) {
      const sheet = this.getSheetByName(sheetName);
      if (sheet) this.setActiveSheet(sheet.id);
    }
    if (a1Ref) {
      const parsed = parseCellRef(a1Ref);
      if (parsed) {
        this.selectionManager.select(parsed.row, parsed.col);
        this.renderer.ensureCellVisible(parsed.row, parsed.col);
      }
    }
    this.render();
  }

  /** Batch mode — defers recalculate and render until done */
  batch(fn) {
    const origRecalc = this.recalculate.bind(this);
    const origRender = this.render.bind(this);
    this.recalculate = () => {};
    this.render = () => {};
    try {
      fn();
    } finally {
      this.recalculate = origRecalc;
      this.render = origRender;
    }
    this.recalculate();
    this.render();
  }

  /** Full state snapshot for agent context */
  getState() {
    const sheet = this.activeSheet;
    const sel = this.selectionManager;
    const editor = this.editor;

    return {
      activeSheet: sheet ? sheet.name : null,
      sheets: this.sheets.map(s => ({
        name: s.name,
        id: s.id,
        rowCount: s.rowCount,
        colCount: s.colCount,
        lastRow: s.getLastRow(),
        lastColumn: s.getLastColumn(),
        frozenRows: s.frozenRows,
        frozenCols: s.frozenCols,
      })),
      selection: {
        range: this.selection ? this.selection.toString() : null,
        activeCell: sheet ? cellRefToString(sel.activeCol, sel.activeRow) : null,
        activeRow: sel.activeRow + 1,  // 1-based
        activeCol: sel.activeCol + 1,
      },
      editing: {
        active: editor ? editor.isActive : false,
        mode: editor ? editor.mode : 'ready',
        value: editor && editor.isActive ? editor.textarea.value : null,
        cell: editor && editor.isActive ? cellRefToString(editor.editCol, editor.editRow) : null,
      },
    };
  }

  // ── Sheets ──

  get activeSheet() {
    return this.sheets[this.activeSheetIndex] || null;
  }

  get selection() {
    return this.selectionManager.selection;
  }

  get activeRow() {
    return this.selectionManager.activeRow;
  }

  get activeCol() {
    return this.selectionManager.activeCol;
  }

  addSheet(name) {
    const idx = this.sheets.length;
    if (!name) name = 'Sheet' + (idx + 1);
    const sheet = new Sheet(name, { rows: this.options.rows, cols: this.options.cols });
    this.sheets.push(sheet);
    this.activeSheetIndex = idx;
    if (this.sheetTabs) this.sheetTabs.update();
    this.emit('sheetAdded', sheet);
    return sheet;
  }

  deleteSheet(id) {
    if (this.sheets.length <= 1) return;
    const idx = this.sheets.findIndex(s => s.id === id);
    if (idx === -1) return;
    this.sheets.splice(idx, 1);
    if (this.activeSheetIndex >= this.sheets.length) {
      this.activeSheetIndex = this.sheets.length - 1;
    }
    if (this.sheetTabs) this.sheetTabs.update();
    this.render();
  }

  duplicateSheet(id) {
    const src = this.sheets.find(s => s.id === id);
    if (!src) return;
    const json = src.toJSON();
    json.name = src.name + ' (Copy)';
    const sheet = new Sheet(json.name, { rows: json.rowCount, cols: json.colCount });
    // Copy cells
    for (const [key, cell] of src.cells) {
      sheet.cells.set(key, cell.clone());
    }
    sheet.colWidths = new Map(src.colWidths);
    sheet.rowHeights = new Map(src.rowHeights);
    sheet.frozenRows = src.frozenRows;
    sheet.frozenCols = src.frozenCols;
    sheet.merges = [...src.merges];
    this.sheets.push(sheet);
    this.activeSheetIndex = this.sheets.length - 1;
    if (this.sheetTabs) this.sheetTabs.update();
    this.render();
  }

  setActiveSheet(id) {
    const idx = this.sheets.findIndex(s => s.id === id);
    if (idx !== -1) {
      this.activeSheetIndex = idx;
      this.selectionManager.select(0, 0);
      this.renderer.scrollTo(0, 0);
      this.recalculate();
      this.render();
      if (this.sheetTabs) this.sheetTabs.update();
    }
  }

  getSheetById(id) {
    return this.sheets.find(s => s.id === id) || null;
  }

  getSheetByName(name) {
    return this.sheets.find(s => s.name === name) || null;
  }

  // ── Cell operations ──

  setCellValue(ref, value) {
    const parsed = parseCellRef(ref);
    if (!parsed) return;
    this.activeSheet.setCellValue(parsed.row, parsed.col, value);
    this.recalculate();
    this.render();
  }

  getCellValue(ref) {
    const parsed = parseCellRef(ref);
    if (!parsed) return null;
    return this.activeSheet.getCellValue(parsed.row, parsed.col);
  }

  setCellFormula(ref, formula) {
    const parsed = parseCellRef(ref);
    if (!parsed) return;
    this.activeSheet.setCellFormula(parsed.row, parsed.col, formula);
    this.recalculate();
    this.render();
  }

  // ── Style operations ──

  setStyle(props) {
    const sel = this.selection;
    const sheet = this.activeSheet;
    if (!sel || !sheet) return;

    this.commandManager.beginBatch();
    sel.forEach((r, c) => {
      const cell = sheet.getCell(r, c);
      const oldProps = {};
      for (const key of Object.keys(props)) {
        oldProps[key] = cell && cell.style ? cell.style[key] : undefined;
      }
      this.commandManager.pushToCurrentBatch({
        type: 'setStyle', row: r, col: c,
        oldProps, newProps: { ...props },
      });
      sheet.setCellStyle(r, c, props);
    });
    this.commandManager.endBatch();
    this.render();
  }

  _activeCell() {
    const sheet = this.activeSheet;
    return sheet ? sheet.getCell(this.activeRow, this.activeCol) : null;
  }

  toggleBold() {
    const cell = this._activeCell();
    this.setStyle({ bold: !(cell && cell.style ? cell.style.bold : false) });
  }

  toggleItalic() {
    const cell = this._activeCell();
    this.setStyle({ italic: !(cell && cell.style ? cell.style.italic : false) });
  }

  toggleUnderline() {
    const cell = this._activeCell();
    this.setStyle({ underline: !(cell && cell.style ? cell.style.underline : false) });
  }

  toggleStrikethrough() {
    const cell = this._activeCell();
    this.setStyle({ strikethrough: !(cell && cell.style ? cell.style.strikethrough : false) });
  }

  toggleWrap() {
    const cell = this._activeCell();
    this.setStyle({ wrap: !(cell && cell.style ? cell.style.wrap : false) });
  }

  setAllBorders(border) {
    this.setStyle({
      borderTop: border,
      borderRight: border,
      borderBottom: border,
      borderLeft: border,
    });
  }

  clearBorders() {
    this.setStyle({
      borderTop: null,
      borderRight: null,
      borderBottom: null,
      borderLeft: null,
    });
  }

  clearFormatting() {
    const sel = this.selection;
    const sheet = this.activeSheet;
    if (!sel || !sheet) return;
    sheet.clearRangeFormat(sel);
    this.render();
  }

  // ── Format painter ──

  startFormatPainter() {
    const sheet = this.activeSheet;
    if (!sheet) return;
    const cell = sheet.getCell(this.activeRow, this.activeCol);
    this._formatPainterStyle = cell && cell.style ? cell.style.clone() : null;
    if (this.canvas) this.canvas.style.cursor = 'copy';
  }

  applyFormatPainter(range) {
    if (!this._formatPainterStyle) return;
    const sheet = this.activeSheet;
    if (!sheet) return;
    range.forEach((r, c) => {
      sheet.setCellStyle(r, c, this._formatPainterStyle);
    });
    this._formatPainterStyle = null;
    if (this.canvas) this.canvas.style.cursor = '';
    this.render();
  }

  // ── Paste special ──

  pasteValuesOnly() {
    const cb = this.clipboard;
    if (!cb.copiedData) return;
    const sheet = this.activeSheet;
    if (!sheet) return;
    const sel = this.selectionManager;

    for (const cellData of cb.copiedData.cells) {
      const r = sel.activeRow + cellData.row;
      const c = sel.activeCol + cellData.col;
      if (r >= sheet.rowCount || c >= sheet.colCount) continue;
      // Only paste computed values, no formulas or styles
      const val = cellData.value;
      sheet.setCellValue(r, c, val);
    }
    this.recalculate();
    this.render();
  }

  // ── Alternating colors (banding) ──

  applyAlternatingColors(color1 = '#ffffff', color2 = '#f3f3f3', headerColor = '#e8eaed') {
    const sel = this.selection;
    const sheet = this.activeSheet;
    if (!sel || !sheet) return;

    for (let r = sel.startRow; r <= sel.endRow; r++) {
      const bg = r === sel.startRow ? headerColor : (r % 2 === 0 ? color1 : color2);
      for (let c = sel.startCol; c <= sel.endCol; c++) {
        sheet.setCellStyle(r, c, { bgColor: bg });
      }
      if (r === sel.startRow) {
        for (let c = sel.startCol; c <= sel.endCol; c++) {
          sheet.setCellStyle(r, c, { bold: true });
        }
      }
    }
    this.render();
  }

  pasteFormatOnly() {
    const cb = this.clipboard;
    if (!cb.copiedData) return;
    const sheet = this.activeSheet;
    if (!sheet) return;
    const sel = this.selectionManager;

    for (const cellData of cb.copiedData.cells) {
      const r = sel.activeRow + cellData.row;
      const c = sel.activeCol + cellData.col;
      if (r >= sheet.rowCount || c >= sheet.colCount) continue;
      if (cellData.style) {
        sheet.setCellStyle(r, c, cellData.style);
      }
    }
    this.render();
  }

  adjustDecimals(delta) {
    const cell = this._activeCell();
    const style = cell ? cell.getStyle() : new CellStyle();
    let fmt = style.numberFormat || 'General';

    if (fmt === 'General') fmt = '#,##0';

    const parts = fmt.split('.');
    let decimals = parts[1] ? parts[1].replace(/[^0#]/g, '').length : 0;
    decimals = Math.max(0, decimals + delta);

    const newFmt = decimals > 0 ? parts[0] + '.' + '0'.repeat(decimals) : parts[0];
    this.setStyle({ numberFormat: newFmt });
  }

  // ── Clipboard ──

  copy() { this.clipboard.copy(false); }
  cut() { this.clipboard.copy(true); }
  paste() { this.clipboard.paste(this.activeRow, this.activeCol); }

  // ── Undo / Redo ──

  undo() { this.commandManager.undo(); }
  redo() { this.commandManager.redo(); }

  // ── Selection operations ──

  deleteSelection() {
    const sel = this.selection;
    const sheet = this.activeSheet;
    if (!sel || !sheet) return;

    this.commandManager.beginBatch();
    sel.forEach((r, c) => {
      const cell = sheet.getCell(r, c);
      if (cell) {
        this.commandManager.pushToCurrentBatch({
          type: 'setCellValue', row: r, col: c,
          oldValue: cell.rawValue,
          oldFormula: cell.formula,
          newValue: null,
        });
      }
    });
    this.commandManager.endBatch();

    sheet.clearRangeContent(sel);
    this.recalculate();
    this.render();
  }

  // ── Row/Col operations ──

  insertRowAbove() {
    const row = this.selectionManager.activeRow;
    this.activeSheet.insertRows(row);
    this.commandManager.execute({ type: 'insertRows', at: row, count: 1 });
    this.emit('change', { changeType: 'INSERT_ROW', source: this });
    this.recalculate();
    this.render();
  }

  insertRowBelow() {
    const sel = this.selection;
    const at = sel ? sel.endRow + 1 : this.activeRow + 1;
    this.activeSheet.insertRows(at);
    this.commandManager.execute({ type: 'insertRows', at, count: 1 });
    this.emit('change', { changeType: 'INSERT_ROW', source: this });
    this.recalculate();
    this.render();
  }

  insertColLeft() {
    const col = this.selectionManager.activeCol;
    this.activeSheet.insertCols(col);
    this.commandManager.execute({ type: 'insertCols', at: col, count: 1 });
    this.emit('change', { changeType: 'INSERT_COLUMN', source: this });
    this.recalculate();
    this.render();
  }

  insertColRight() {
    const sel = this.selection;
    const at = sel ? sel.endCol + 1 : this.activeCol + 1;
    this.activeSheet.insertCols(at);
    this.commandManager.execute({ type: 'insertCols', at, count: 1 });
    this.emit('change', { changeType: 'INSERT_COLUMN', source: this });
    this.recalculate();
    this.render();
  }

  deleteSelectedRows() {
    const sel = this.selection;
    const sheet = this.activeSheet;
    if (!sel || !sheet) return;
    const count = sel.endRow - sel.startRow + 1;

    // Save cell data for undo
    const savedData = [];
    for (let r = sel.startRow; r <= sel.endRow; r++) {
      for (let c = 0; c < sheet.colCount; c++) {
        const cell = sheet.getCell(r, c);
        if (cell) savedData.push({ row: r, col: c, cellJSON: cell.toJSON() });
      }
    }

    sheet.deleteRows(sel.startRow, count);
    this.commandManager.execute({ type: 'deleteRows', at: sel.startRow, count, savedData });
    this.selectionManager.select(Math.min(sel.startRow, sheet.rowCount - 1), this.activeCol);
    this.emit('change', { changeType: 'REMOVE_ROW', source: this });
    this.recalculate();
    this.render();
  }

  deleteSelectedCols() {
    const sel = this.selection;
    const sheet = this.activeSheet;
    if (!sel || !sheet) return;
    const count = sel.endCol - sel.startCol + 1;

    // Save cell data for undo
    const savedData = [];
    for (let r = 0; r < sheet.rowCount; r++) {
      for (let c = sel.startCol; c <= sel.endCol; c++) {
        const cell = sheet.getCell(r, c);
        if (cell) savedData.push({ row: r, col: c, cellJSON: cell.toJSON() });
      }
    }

    sheet.deleteCols(sel.startCol, count);
    this.commandManager.execute({ type: 'deleteCols', at: sel.startCol, count, savedData });
    this.selectionManager.select(this.activeRow, Math.min(sel.startCol, sheet.colCount - 1));
    this.emit('change', { changeType: 'REMOVE_COLUMN', source: this });
    this.recalculate();
    this.render();
  }

  // ── Merge ──

  mergeSelection() {
    const sel = this.selection;
    if (!sel || sel.isSingleCell) return;
    this.activeSheet.addMerge(sel);
    this.commandManager.execute({
      type: 'merge', rangeStr: sel.toString(),
      startRow: sel.startRow, startCol: sel.startCol, endRow: sel.endRow, endCol: sel.endCol,
    });
    this.render();
  }

  unmergeSelection() {
    const sel = this.selection;
    if (!sel) return;
    const toRemove = [];
    for (const m of this.activeSheet.merges) {
      const range = CellRange.fromString(m);
      if (range && range.intersects(sel)) toRemove.push(range);
    }
    this.commandManager.beginBatch();
    for (const range of toRemove) {
      this.activeSheet.removeMerge(range);
      this.commandManager.pushToCurrentBatch({
        type: 'unmerge', rangeStr: range.toString(),
        startRow: range.startRow, startCol: range.startCol, endRow: range.endRow, endCol: range.endCol,
      });
    }
    this.commandManager.endBatch();
    this.render();
  }

  // ── Freeze ──

  freezeRows(count) {
    this.activeSheet.frozenRows = count;
    this.render();
  }

  freezeCols(count) {
    this.activeSheet.frozenCols = count;
    this.render();
  }

  // ── Sort ──

  sortAscending() { this._sortSelection(true); }
  sortDescending() { this._sortSelection(false); }

  _sortSelection(asc) {
    const sel = this.selection;
    if (!sel) return;
    const sheet = this.activeSheet;
    const col = this.activeCol;

    // Collect rows
    const rows = [];
    for (let r = sel.startRow; r <= sel.endRow; r++) {
      const rowData = [];
      for (let c = sel.startCol; c <= sel.endCol; c++) {
        const cell = sheet.getCell(r, c);
        rowData.push(cell ? cell.clone() : new Cell());
      }
      const sortVal = sheet.getCellValue(r, col);
      rows.push({ data: rowData, key: sortVal });
    }

    rows.sort((a, b) => {
      const va = a.key, vb = b.key;
      if (va === null && vb === null) return 0;
      if (va === null) return 1;
      if (vb === null) return -1;
      const cmp = va < vb ? -1 : va > vb ? 1 : 0;
      return asc ? cmp : -cmp;
    });

    // Write back
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].data.length; c++) {
        sheet.setCell(sel.startRow + r, sel.startCol + c, rows[r].data[c]);
      }
    }

    this.recalculate();
    this.render();
  }

  // ── Auto-fill ──

  autoFill(srcRange, destRange) {
    const sheet = this.activeSheet;

    // Determine fill direction
    const fillDown = destRange.endRow > srcRange.endRow;
    const fillRight = destRange.endCol > srcRange.endCol;
    const fillUp = destRange.startRow < srcRange.startRow;
    const fillLeft = destRange.startCol < srcRange.startCol;

    if (fillDown || fillUp) {
      const startR = fillDown ? srcRange.endRow + 1 : destRange.startRow;
      const endR = fillDown ? destRange.endRow : srcRange.startRow - 1;

      for (let r = startR; r <= endR; r++) {
        const srcRow = srcRange.startRow + ((r - srcRange.startRow) % srcRange.rowCount);
        for (let c = srcRange.startCol; c <= srcRange.endCol; c++) {
          const srcCell = sheet.getCell(srcRow, c);
          if (!srcCell) { sheet.clearCell(r, c); continue; }

          if (srcCell.formula) {
            const dRow = r - srcRow;
            const adjusted = FormulaEngine.copyFormula(srcCell.formula, dRow, 0);
            sheet.setCellFormula(r, c, adjusted);
          } else if (typeof srcCell.computedValue === 'number') {
            // Detect pattern
            const vals = [];
            for (let sr = srcRange.startRow; sr <= srcRange.endRow; sr++) {
              const sc = sheet.getCell(sr, c);
              if (sc && typeof sc.computedValue === 'number') vals.push(sc.computedValue);
            }
            if (vals.length >= 2) {
              const step = vals[vals.length - 1] - vals[vals.length - 2];
              const offset = r - srcRange.endRow;
              sheet.setCellValue(r, c, vals[vals.length - 1] + step * offset);
            } else {
              sheet.setCellValue(r, c, srcCell.rawValue);
            }
          } else {
            sheet.setCellValue(r, c, srcCell.rawValue);
          }

          if (srcCell.style) {
            sheet.setCellStyle(r, c, srcCell.style.clone());
          }
        }
      }
    }

    if (fillRight || fillLeft) {
      const startC = fillRight ? srcRange.endCol + 1 : destRange.startCol;
      const endC = fillRight ? destRange.endCol : srcRange.startCol - 1;

      for (let c = startC; c <= endC; c++) {
        const srcCol = srcRange.startCol + ((c - srcRange.startCol) % srcRange.colCount);
        for (let r = srcRange.startRow; r <= srcRange.endRow; r++) {
          const srcCell = sheet.getCell(r, srcCol);
          if (!srcCell) { sheet.clearCell(r, c); continue; }

          if (srcCell.formula) {
            const dCol = c - srcCol;
            const adjusted = FormulaEngine.copyFormula(srcCell.formula, 0, dCol);
            sheet.setCellFormula(r, c, adjusted);
          } else {
            sheet.setCellValue(r, c, srcCell.rawValue);
          }

          if (srcCell.style) {
            sheet.setCellStyle(r, c, srcCell.style.clone());
          }
        }
      }
    }

    this.recalculate();
    this.render();
  }

  // ── Conditional formatting ──

  getConditionalStyle(sheet, row, col) {
    for (const cf of sheet.conditionalFormats) {
      const range = CellRange.fromString(cf.range);
      if (!range || !range.contains(row, col)) continue;

      const cellValue = sheet.getCellValue(row, col);
      if (this._matchCondition(cellValue, cf)) {
        return cf.style;
      }
    }
    return null;
  }

  _matchCondition(value, cf) {
    switch (cf.type) {
      case 'greaterThan': return typeof value === 'number' && value > cf.value;
      case 'lessThan': return typeof value === 'number' && value < cf.value;
      case 'equalTo': return value === cf.value;
      case 'between': return typeof value === 'number' && value >= cf.min && value <= cf.max;
      case 'text_contains': return typeof value === 'string' && value.toLowerCase().includes(cf.value.toLowerCase());
      case 'duplicate': return false; // TODO: implement
      default: return false;
    }
  }

  addConditionalFormat(rangeStr, type, opts) {
    this.activeSheet.conditionalFormats.push({
      range: rangeStr,
      type,
      ...opts,
    });
    this.render();
  }

  // ── Filter ──

  toggleFilter() {
    const sheet = this.activeSheet;
    if (!sheet) return;

    if (sheet.filterRange) {
      // Remove filter
      sheet.filterRange = null;
      sheet.filterCriteria.clear();
      sheet.filteredRows.clear();
      this.render();
      return;
    }

    // Auto-detect data range or use selection
    const sel = this.selection;
    let range;
    if (sel && !sel.isSingleCell) {
      range = sel;
    } else {
      range = sheet.getUsedRange();
    }
    if (!range) return;

    sheet.filterRange = range.toString();
    sheet.filterCriteria.clear();
    sheet.filteredRows.clear();
    this.render();
  }

  showFilterDropdown(col) {
    const sheet = this.activeSheet;
    if (!sheet || !sheet.filterRange) return;

    const range = CellRange.fromString(sheet.filterRange);
    if (!range) return;

    // Collect unique values in the column
    const values = new Map();
    for (let r = range.startRow + 1; r <= range.endRow; r++) {
      const val = sheet.getCellValue(r, col);
      const key = val === null || val === undefined ? '' : String(val);
      values.set(key, (values.get(key) || 0) + 1);
    }

    const existing = sheet.filterCriteria.get(col) || null;
    const hidden = existing ? new Set(existing) : new Set();

    // Build filter dropdown
    const overlay = document.createElement('div');
    overlay.className = 'sheets-filter-overlay';

    const dialog = document.createElement('div');
    dialog.className = 'sheets-filter-dialog';

    // Sort A→Z / Z→A
    const sortSection = document.createElement('div');
    sortSection.className = 'sheets-filter-sort';
    const sortAZ = document.createElement('div');
    sortAZ.className = 'sheets-filter-sort-btn';
    sortAZ.textContent = 'Sort A → Z';
    sortAZ.onclick = () => { this._filterSort(col, true); overlay.remove(); };
    const sortZA = document.createElement('div');
    sortZA.className = 'sheets-filter-sort-btn';
    sortZA.textContent = 'Sort Z → A';
    sortZA.onclick = () => { this._filterSort(col, false); overlay.remove(); };
    sortSection.appendChild(sortAZ);
    sortSection.appendChild(sortZA);
    dialog.appendChild(sortSection);

    dialog.appendChild(document.createElement('hr'));

    // Select all / Deselect all
    const toggleRow = document.createElement('div');
    toggleRow.className = 'sheets-filter-toggle';
    const selectAll = document.createElement('a');
    selectAll.textContent = 'Select all';
    selectAll.href = '#';
    selectAll.onclick = (e) => { e.preventDefault(); checkboxes.forEach(cb => cb.checked = true); };
    const deselectAll = document.createElement('a');
    deselectAll.textContent = 'Clear';
    deselectAll.href = '#';
    deselectAll.onclick = (e) => { e.preventDefault(); checkboxes.forEach(cb => cb.checked = false); };
    toggleRow.appendChild(selectAll);
    toggleRow.appendChild(document.createTextNode(' | '));
    toggleRow.appendChild(deselectAll);
    dialog.appendChild(toggleRow);

    // Checkbox list
    const list = document.createElement('div');
    list.className = 'sheets-filter-list';
    const checkboxes = [];
    const sortedKeys = [...values.keys()].sort();

    for (const key of sortedKeys) {
      const row = document.createElement('label');
      row.className = 'sheets-filter-item';
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = !hidden.has(key);
      cb.dataset.value = key;
      checkboxes.push(cb);
      row.appendChild(cb);
      row.appendChild(document.createTextNode(' ' + (key || '(Blanks)') + ' (' + values.get(key) + ')'));
      list.appendChild(row);
    }
    dialog.appendChild(list);

    // OK / Cancel
    const actions = document.createElement('div');
    actions.className = 'sheets-filter-actions';
    const okBtn = document.createElement('button');
    okBtn.textContent = 'OK';
    okBtn.className = 'sheets-filter-ok';
    okBtn.onclick = () => {
      const newHidden = new Set();
      checkboxes.forEach(cb => { if (!cb.checked) newHidden.add(cb.dataset.value); });
      if (newHidden.size > 0) {
        sheet.filterCriteria.set(col, newHidden);
      } else {
        sheet.filterCriteria.delete(col);
      }
      this._applyFilter();
      overlay.remove();
    };
    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    cancelBtn.onclick = () => overlay.remove();
    actions.appendChild(cancelBtn);
    actions.appendChild(okBtn);
    dialog.appendChild(actions);

    overlay.appendChild(dialog);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });
    document.body.appendChild(overlay);

    // Position near the column header
    const rect = this.renderer._getCellRect(sheet, range.startRow, col);
    const containerRect = this.container.getBoundingClientRect();
    dialog.style.left = (containerRect.left + rect.x) + 'px';
    dialog.style.top = (containerRect.top + rect.y + rect.h) + 'px';
  }

  _filterSort(col, asc) {
    const sheet = this.activeSheet;
    if (!sheet || !sheet.filterRange) return;
    const range = CellRange.fromString(sheet.filterRange);
    if (!range) return;

    // Collect rows (skip header)
    const rows = [];
    for (let r = range.startRow + 1; r <= range.endRow; r++) {
      const rowData = [];
      for (let c = range.startCol; c <= range.endCol; c++) {
        const cell = sheet.getCell(r, c);
        rowData.push(cell ? cell.clone() : null);
      }
      rows.push({ data: rowData, key: sheet.getCellValue(r, col) });
    }

    rows.sort((a, b) => {
      const va = a.key, vb = b.key;
      if (va === null && vb === null) return 0;
      if (va === null) return 1;
      if (vb === null) return -1;
      const cmp = va < vb ? -1 : va > vb ? 1 : 0;
      return asc ? cmp : -cmp;
    });

    for (let i = 0; i < rows.length; i++) {
      for (let c = 0; c < rows[i].data.length; c++) {
        const cell = rows[i].data[c];
        if (cell) sheet.setCell(range.startRow + 1 + i, range.startCol + c, cell);
        else sheet.clearCell(range.startRow + 1 + i, range.startCol + c);
      }
    }

    this.recalculate();
    this.render();
  }

  _applyFilter() {
    const sheet = this.activeSheet;
    if (!sheet || !sheet.filterRange) return;
    const range = CellRange.fromString(sheet.filterRange);
    if (!range) return;

    sheet.filteredRows.clear();

    for (let r = range.startRow + 1; r <= range.endRow; r++) {
      for (const [col, hiddenSet] of sheet.filterCriteria) {
        const val = sheet.getCellValue(r, col);
        const key = val === null || val === undefined ? '' : String(val);
        if (hiddenSet.has(key)) {
          sheet.filteredRows.add(r);
          break;
        }
      }
    }

    this.render();
  }

  // ── Find / Replace ──

  showFindReplace(replace = false) {
    this.findReplace.show(replace);
  }

  // ── Data Validation ──

  showDataValidation() {
    this.dataValidation.showDialog();
  }

  setDataValidation(rangeStr, rule) {
    const sheet = this.activeSheet;
    if (!sheet) return;
    const range = CellRange.fromString(rangeStr);
    if (!range) return;
    this.dataValidation.setValidation(sheet, range.startRow, range.startCol, range.endRow, range.endCol, rule);
  }

  // ── Named Ranges ──

  showNamedRanges() {
    this.namedRanges.showDialog();
  }

  defineNamedRange(name, rangeStr) {
    const sheet = this.activeSheet;
    if (!sheet) return;
    const range = CellRange.fromString(rangeStr);
    if (!range) return;
    this.namedRanges.define(name, sheet.id, range);
  }

  getNamedRange(name) {
    return this.namedRanges.resolve(name);
  }

  // ── Conditional Formatting ──

  showConditionalFormatting() {
    this.conditionalFormatUI.showDialog();
  }

  // ── Formula recalculation ──

  recalculate() {
    // Evaluate formulas on ALL sheets (not just active) for cross-sheet references
    for (const sheet of this.sheets) {
      for (const [key, cell] of sheet.cells) {
        if (cell.formula) {
          const row = key >> 16;
          const col = key & 0xffff;
          const result = this.formulaEngine.evaluate(cell.formula, sheet.id, row, col);
          cell.computedValue = result === null || result === undefined ? '' : result;
        }
      }
    }
  }

  // ── Rendering ──

  render() {
    this.renderer.requestRender();
    if (this.formulaBar) this.formulaBar.update();
  }

  // ── Events ──

  on(event, fn) { return this._events.on(event, fn); }
  off(event, fn) { this._events.off(event, fn); }
  emit(event, ...args) { this._events.emit(event, ...args); }

  // ── Import / Export ──

  toJSON() {
    return {
      sheets: this.sheets.map(s => s.toJSON()),
      activeSheet: this.activeSheetIndex,
    };
  }

  fromJSON(data) {
    if (data.sheets) {
      this.sheets = [];
      for (const sheetData of data.sheets) {
        const sheet = new Sheet(sheetData.name, { rows: sheetData.rowCount, cols: sheetData.colCount });
        if (sheetData.cells) {
          for (const [ref, cellData] of Object.entries(sheetData.cells)) {
            const parsed = parseCellRef(ref);
            if (parsed) {
              sheet.cells.set(cellKey(parsed.row, parsed.col), Cell.fromJSON(cellData));
            }
          }
        }
        if (sheetData.colWidths) {
          for (const [col, w] of Object.entries(sheetData.colWidths)) {
            sheet.colWidths.set(Number(col), w);
          }
        }
        if (sheetData.rowHeights) {
          for (const [row, h] of Object.entries(sheetData.rowHeights)) {
            sheet.rowHeights.set(Number(row), h);
          }
        }
        sheet.frozenRows = sheetData.frozenRows || 0;
        sheet.frozenCols = sheetData.frozenCols || 0;
        sheet.merges = sheetData.merges || [];
        sheet.hiddenRows = new Set(sheetData.hiddenRows || []);
        sheet.hiddenCols = new Set(sheetData.hiddenCols || []);
        this.sheets.push(sheet);
      }
      this.activeSheetIndex = data.activeSheet || 0;
    }
    if (this.sheetTabs) this.sheetTabs.update();
    this.recalculate();
    this.render();
  }

  importCSV(text, delimiter) {
    const rows = parseCSV(text, delimiter);
    const sheet = this.activeSheet;
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        sheet.setCellValue(r, c, rows[r][c]);
      }
    }
    this.recalculate();
    this.render();
  }

  exportCSV(delimiter) {
    return generateCSV(this.activeSheet, delimiter);
  }

  // ── Bulk data ──

  setData(data, startRow = 0, startCol = 0) {
    const sheet = this.activeSheet;
    for (let r = 0; r < data.length; r++) {
      if (!Array.isArray(data[r])) continue;
      for (let c = 0; c < data[r].length; c++) {
        sheet.setCellValue(startRow + r, startCol + c, data[r][c]);
      }
    }
    this.recalculate();
    this.render();
  }

  getData(range) {
    const sheet = this.activeSheet;
    if (!range) {
      const used = sheet.getUsedRange();
      if (!used) return [];
      range = used;
    }
    return range.map((r, c) => sheet.getCellValue(r, c));
  }

  // ── Cleanup ──

  destroy() {
    this.renderer.destroy();
    this.keyboard.destroy();
    this.mouse.destroy();
    this.clipboard.destroy();
    this.editor.destroy();
    this.formulaHelper.destroy();
    if (this.toolbar) this.toolbar.destroy();
    if (this.formulaBar) this.formulaBar.destroy();
    if (this.sheetTabs) this.sheetTabs.destroy();
    if (this.contextMenu) this.contextMenu.destroy();
    if (this.findReplace) this.findReplace.destroy();
    if (this.statusBar) this.statusBar.destroy();
    if (this._resizeObserver) this._resizeObserver.disconnect();
    if (this.container && this.container.parentElement) {
      this.container.parentElement.removeChild(this.container);
    }
    this._events.removeAll();
  }
}
