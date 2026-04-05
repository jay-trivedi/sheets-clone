import CellRange from '../core/CellRange.js';
import FormulaEngine from '../formula/FormulaEngine.js';
import { cellKey, keyToRC, indexToCol } from '../utils/helpers.js';

export default class ClipboardHandler {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.copiedData = null;
    this.copyRange = null;
    this.isCut = false;

    this._onPaste = this._onPaste.bind(this);
    this._onCopy = this._onCopy.bind(this);
    this._onCut = this._onCut.bind(this);
  }

  init(container) {
    this.container = container;
    container.addEventListener('paste', this._onPaste);
    container.addEventListener('copy', this._onCopy);
    container.addEventListener('cut', this._onCut);
  }

  destroy() {
    this.container.removeEventListener('paste', this._onPaste);
    this.container.removeEventListener('copy', this._onCopy);
    this.container.removeEventListener('cut', this._onCut);
  }

  copy(asCut = false) {
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    const sel = ss.selection;
    if (!sel) return;

    this.copyRange = new CellRange(sel.startRow, sel.startCol, sel.endRow, sel.endCol);
    this.isCut = asCut;

    // Build clipboard data
    const data = { cells: [], range: this.copyRange, sheetId: sheet.id };
    const textRows = [];

    for (let r = sel.startRow; r <= sel.endRow; r++) {
      const row = [];
      const textCols = [];
      for (let c = sel.startCol; c <= sel.endCol; c++) {
        const cell = sheet.getCell(r, c);
        if (cell) {
          row.push({
            row: r - sel.startRow,
            col: c - sel.startCol,
            value: cell.rawValue,
            formula: cell.formula,
            style: cell.style ? cell.style.clone() : null,
          });
          textCols.push(cell.displayValue);
        } else {
          row.push({ row: r - sel.startRow, col: c - sel.startCol, value: null });
          textCols.push('');
        }
      }
      data.cells.push(...row);
      textRows.push(textCols.join('\t'));
    }

    this.copiedData = data;
    ss.copyRange = this.copyRange;

    // Also put tab-separated text on system clipboard
    const text = textRows.join('\n');
    try {
      navigator.clipboard.writeText(text).catch(() => {});
    } catch (e) {}

    ss.render();
  }

  paste(targetRow, targetCol) {
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;

    if (!this.copiedData) return;

    const data = this.copiedData;
    const srcRange = data.range;
    const dRow = targetRow - srcRange.startRow;
    const dCol = targetCol - srcRange.startCol;

    ss.commandManager.beginBatch();

    for (const cellData of data.cells) {
      const r = targetRow + cellData.row;
      const c = targetCol + cellData.col;

      if (r >= sheet.rowCount || c >= sheet.colCount) continue;

      const oldCell = sheet.getCell(r, c);
      ss.commandManager.pushToCurrentBatch({
        type: 'setCellValue',
        row: r, col: c,
        oldValue: oldCell ? oldCell.rawValue : null,
        oldFormula: oldCell ? oldCell.formula : null,
        oldStyle: oldCell && oldCell.style ? oldCell.style.clone() : null,
      });

      if (cellData.formula) {
        const adjusted = FormulaEngine.copyFormula(cellData.formula, dRow, dCol);
        sheet.setCellFormula(r, c, adjusted);
      } else {
        sheet.setCellValue(r, c, cellData.value);
      }

      if (cellData.style) {
        sheet.setCellStyle(r, c, cellData.style);
      }
    }

    // If cut, clear source — also track for undo
    if (this.isCut) {
      for (const cellData of data.cells) {
        const r = srcRange.startRow + cellData.row;
        const c = srcRange.startCol + cellData.col;
        const srcCell = sheet.getCell(r, c);
        // Track the source cell clearing so undo restores it
        ss.commandManager.pushToCurrentBatch({
          type: 'setCellValue',
          row: r, col: c,
          oldValue: srcCell ? srcCell.rawValue : null,
          oldFormula: srcCell ? srcCell.formula : null,
          oldStyle: srcCell && srcCell.style ? srcCell.style.clone() : null,
        });
        sheet.clearCell(r, c);
      }
      this.copiedData = null;
      this.copyRange = null;
      ss.copyRange = null;
      this.isCut = false;
    }

    ss.commandManager.endBatch();
    ss.recalculate();
    ss.render();
  }

  _onCopy(e) {
    if (this.spreadsheet.editor && this.spreadsheet.editor.isActive) return;
    e.preventDefault();
    this.copy(false);
  }

  _onCut(e) {
    if (this.spreadsheet.editor && this.spreadsheet.editor.isActive) return;
    e.preventDefault();
    this.copy(true);
  }

  _onPaste(e) {
    if (this.spreadsheet.editor && this.spreadsheet.editor.isActive) return;
    e.preventDefault();

    const ss = this.spreadsheet;
    const sel = ss.selectionManager;

    // Try internal paste first
    if (this.copiedData) {
      this.paste(sel.activeRow, sel.activeCol);
      return;
    }

    // External paste (from clipboard)
    const text = e.clipboardData.getData('text/plain');
    if (text) {
      this._pasteText(text, sel.activeRow, sel.activeCol);
    }
  }

  _pasteText(text, startRow, startCol) {
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    const rows = text.split('\n').map(r => r.split('\t'));

    ss.commandManager.beginBatch();

    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        const row = startRow + r;
        const col = startCol + c;
        if (row >= sheet.rowCount || col >= sheet.colCount) continue;

        const oldCell = sheet.getCell(row, col);
        ss.commandManager.pushToCurrentBatch({
          type: 'setCellValue',
          row, col,
          oldValue: oldCell ? oldCell.rawValue : null,
          oldFormula: oldCell ? oldCell.formula : null,
        });

        sheet.setCellValue(row, col, rows[r][c]);
      }
    }

    ss.commandManager.endBatch();
    ss.recalculate();
    ss.render();
  }
}
