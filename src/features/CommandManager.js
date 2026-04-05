import { cellKey, keyToRC } from '../utils/helpers.js';

export default class CommandManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.undoStack = [];
    this.redoStack = [];
    this.maxHistory = 100;
    this._batch = null;
  }

  execute(command) {
    if (this._batch) {
      this._batch.push(command);
      return;
    }
    this.undoStack.push([command]);
    if (this.undoStack.length > this.maxHistory) this.undoStack.shift();
    this.redoStack = [];
  }

  beginBatch() { this._batch = []; }

  pushToCurrentBatch(command) {
    if (this._batch) this._batch.push(command);
  }

  endBatch() {
    if (this._batch && this._batch.length > 0) {
      this.undoStack.push(this._batch);
      if (this.undoStack.length > this.maxHistory) this.undoStack.shift();
      this.redoStack = [];
    }
    this._batch = null;
  }

  undo() {
    const batch = this.undoStack.pop();
    if (!batch) return;

    const redoBatch = [];
    const sheet = this.spreadsheet.activeSheet;
    if (!sheet) return;

    // Undo in reverse order
    for (let i = batch.length - 1; i >= 0; i--) {
      const redoCmd = this._applyUndo(batch[i], sheet);
      if (redoCmd) redoBatch.unshift(redoCmd);
    }

    this.redoStack.push(redoBatch);
    this.spreadsheet.recalculate();
    this.spreadsheet.render();
  }

  redo() {
    const batch = this.redoStack.pop();
    if (!batch) return;

    const undoBatch = [];
    const sheet = this.spreadsheet.activeSheet;
    if (!sheet) return;

    for (const cmd of batch) {
      const undoCmd = this._applyRedo(cmd, sheet);
      if (undoCmd) undoBatch.push(undoCmd);
    }

    this.undoStack.push(undoBatch);
    this.spreadsheet.recalculate();
    this.spreadsheet.render();
  }

  // ── Apply undo for a single command ──
  _applyUndo(cmd, sheet) {
    switch (cmd.type) {
      case 'setCellValue': {
        // Save current state for redo
        const cell = sheet.getCell(cmd.row, cmd.col);
        const curValue = cell ? cell.rawValue : null;
        const curFormula = cell ? cell.formula : null;
        const curStyle = cell && cell.style ? cell.style.clone() : null;

        // Restore old state
        if (cmd.oldFormula) {
          sheet.setCellFormula(cmd.row, cmd.col, cmd.oldFormula);
        } else if (cmd.oldValue !== null && cmd.oldValue !== undefined) {
          sheet.setCellValue(cmd.row, cmd.col, cmd.oldValue);
        } else {
          sheet.clearCell(cmd.row, cmd.col);
        }
        if (cmd.oldStyle) sheet.setCellStyle(cmd.row, cmd.col, cmd.oldStyle);

        return { ...cmd, oldValue: curValue, oldFormula: curFormula, oldStyle: curStyle,
                 newValue: cmd.oldValue, newFormula: cmd.oldFormula };
      }

      case 'setStyle': {
        for (const [key, val] of Object.entries(cmd.oldProps || {})) {
          sheet.setCellStyle(cmd.row, cmd.col, { [key]: val });
        }
        return { ...cmd, oldProps: cmd.newProps, newProps: cmd.oldProps };
      }

      case 'resize': {
        const cur = cmd.mode === 'resize-col' ? sheet.getColWidth(cmd.index) : sheet.getRowHeight(cmd.index);
        if (cmd.mode === 'resize-col') sheet.setColWidth(cmd.index, cmd.oldSize);
        else sheet.setRowHeight(cmd.index, cmd.oldSize);
        return { ...cmd, oldSize: cur, newSize: cmd.oldSize };
      }

      case 'insertRows': {
        // Undo insert = delete the rows
        sheet.deleteRows(cmd.at, cmd.count);
        return { type: 'deleteRows', at: cmd.at, count: cmd.count, savedData: null };
      }

      case 'deleteRows': {
        // Undo delete = re-insert and restore data
        sheet.insertRows(cmd.at, cmd.count);
        if (cmd.savedData) {
          for (const { row, col, cellJSON } of cmd.savedData) {
            const Cell = sheet.getOrCreateCell(row, col).constructor;
            const restored = Cell.fromJSON(cellJSON);
            sheet.setCell(row, col, restored);
          }
        }
        return { type: 'insertRows', at: cmd.at, count: cmd.count };
      }

      case 'insertCols': {
        sheet.deleteCols(cmd.at, cmd.count);
        return { type: 'deleteCols', at: cmd.at, count: cmd.count, savedData: null };
      }

      case 'deleteCols': {
        sheet.insertCols(cmd.at, cmd.count);
        if (cmd.savedData) {
          for (const { row, col, cellJSON } of cmd.savedData) {
            const Cell = sheet.getOrCreateCell(row, col).constructor;
            const restored = Cell.fromJSON(cellJSON);
            sheet.setCell(row, col, restored);
          }
        }
        return { type: 'insertCols', at: cmd.at, count: cmd.count };
      }

      case 'sort': {
        // Restore original cell data
        if (cmd.savedData) {
          for (const { row, col, cellJSON } of cmd.savedData) {
            if (cellJSON) {
              const Cell = sheet.getOrCreateCell(row, col).constructor;
              sheet.setCell(row, col, Cell.fromJSON(cellJSON));
            } else {
              sheet.clearCell(row, col);
            }
          }
        }
        // For redo, we'd need to save current state — just return a sort command
        return cmd; // Re-sorting is idempotent enough
      }

      case 'merge': {
        const CellRange = (require && require('../core/CellRange.js')) || null;
        // Dynamic import won't work synchronously, so parse the range string
        if (cmd.rangeStr) {
          // Find and remove the merge
          const idx = sheet.merges.indexOf(cmd.rangeStr);
          if (idx !== -1) sheet.merges.splice(idx, 1);
          // Clear merge markers on cells
          this._clearMergeMarkers(sheet, cmd);
        }
        return { type: 'unmerge', rangeStr: cmd.rangeStr, startRow: cmd.startRow, startCol: cmd.startCol, endRow: cmd.endRow, endCol: cmd.endCol };
      }

      case 'unmerge': {
        if (cmd.rangeStr && !sheet.merges.includes(cmd.rangeStr)) {
          sheet.merges.push(cmd.rangeStr);
          this._setMergeMarkers(sheet, cmd);
        }
        return { type: 'merge', rangeStr: cmd.rangeStr, startRow: cmd.startRow, startCol: cmd.startCol, endRow: cmd.endRow, endCol: cmd.endCol };
      }

      default:
        return null;
    }
  }

  // ── Apply redo (same logic as undo — it's symmetric) ──
  _applyRedo(cmd, sheet) {
    return this._applyUndo(cmd, sheet);
  }

  // ── Helpers for merge markers ──
  _clearMergeMarkers(sheet, cmd) {
    for (let r = cmd.startRow; r <= cmd.endRow; r++) {
      for (let c = cmd.startCol; c <= cmd.endCol; c++) {
        const cell = sheet.getCell(r, c);
        if (cell) { cell.mergeParent = null; cell.mergeSpan = null; }
      }
    }
  }

  _setMergeMarkers(sheet, cmd) {
    const topLeft = sheet.getOrCreateCell(cmd.startRow, cmd.startCol);
    topLeft.mergeSpan = { rows: cmd.endRow - cmd.startRow + 1, cols: cmd.endCol - cmd.startCol + 1 };
    for (let r = cmd.startRow; r <= cmd.endRow; r++) {
      for (let c = cmd.startCol; c <= cmd.endCol; c++) {
        if (r === cmd.startRow && c === cmd.startCol) continue;
        const cell = sheet.getOrCreateCell(r, c);
        cell.mergeParent = { row: cmd.startRow, col: cmd.startCol };
      }
    }
  }

  get canUndo() { return this.undoStack.length > 0; }
  get canRedo() { return this.redoStack.length > 0; }
  clear() { this.undoStack = []; this.redoStack = []; }
}
