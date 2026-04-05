import CellRange from '../core/CellRange.js';

export default class SelectionManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    // Anchor: the cell where selection started (stays fixed during Shift-extend)
    this.activeRow = 0;
    this.activeCol = 0;
    // Cursor: the extending end of the selection (moves with Shift+Arrow)
    this._cursorRow = 0;
    this._cursorCol = 0;
    this.selection = CellRange.fromCell(0, 0);
  }

  get range() { return this.selection; }
  _sheet() { return this.spreadsheet.activeSheet; }

  _clampRow(r) {
    const sheet = this._sheet();
    return sheet ? Math.max(0, Math.min(r, sheet.rowCount - 1)) : 0;
  }

  _clampCol(c) {
    const sheet = this._sheet();
    return sheet ? Math.max(0, Math.min(c, sheet.colCount - 1)) : 0;
  }

  _updateSelection() {
    this.selection = new CellRange(this.activeRow, this.activeCol, this._cursorRow, this._cursorCol);
    this.spreadsheet.emit('selectionChanged', this.selection);
  }

  _scrollToCursor() {
    this.spreadsheet.renderer.ensureCellVisible(this._cursorRow, this._cursorCol);
  }

  // ── Select a single cell (resets anchor + cursor) ──
  select(row, col) {
    const sheet = this._sheet();
    if (!sheet) return;
    row = this._clampRow(row);
    col = this._clampCol(col);

    const merge = sheet.getMergeAt(row, col);
    if (merge) { row = merge.startRow; col = merge.startCol; }

    this.activeRow = row;
    this.activeCol = col;
    this._cursorRow = row;
    this._cursorCol = col;
    this._updateSelection();
  }

  // ── Extend selection to a cell (anchor stays, cursor moves) ──
  extendTo(row, col) {
    const sheet = this._sheet();
    if (!sheet) return;
    this._cursorRow = this._clampRow(row);
    this._cursorCol = this._clampCol(col);
    this._updateSelection();
  }

  selectRange(range) {
    this.activeRow = range.startRow;
    this.activeCol = range.startCol;
    this._cursorRow = range.endRow;
    this._cursorCol = range.endCol;
    this.selection = range;
    this.spreadsheet.emit('selectionChanged', this.selection);
  }

  selectAll() {
    const sheet = this._sheet();
    if (!sheet) return;
    this._cursorRow = sheet.rowCount - 1;
    this._cursorCol = sheet.colCount - 1;
    this.selection = new CellRange(0, 0, this._cursorRow, this._cursorCol);
    this.spreadsheet.emit('selectionChanged', this.selection);
  }

  selectRow(row) {
    const sheet = this._sheet();
    if (!sheet) return;
    this.activeRow = row;
    this.activeCol = 0;
    this._cursorRow = row;
    this._cursorCol = sheet.colCount - 1;
    this._updateSelection();
  }

  selectCol(col) {
    const sheet = this._sheet();
    if (!sheet) return;
    this.activeRow = 0;
    this.activeCol = col;
    this._cursorRow = sheet.rowCount - 1;
    this._cursorCol = col;
    this._updateSelection();
  }

  // ── Move: navigate (no extend) or extend selection (Shift) ──
  move(dr, dc, extend = false) {
    const sheet = this._sheet();
    if (!sheet) return;

    if (extend) {
      // Shift+Arrow: move cursor (extending end), anchor stays
      let newRow = this._cursorRow + dr;
      let newCol = this._cursorCol + dc;

      // Skip hidden rows/cols
      while (newRow >= 0 && newRow < sheet.rowCount && !sheet.isRowVisible(newRow)) newRow += dr || 1;
      while (newCol >= 0 && newCol < sheet.colCount && !sheet.isColVisible(newCol)) newCol += dc || 1;

      this._cursorRow = this._clampRow(newRow);
      this._cursorCol = this._clampCol(newCol);
      this._updateSelection();
      this._scrollToCursor();
    } else {
      // Plain Arrow: collapse selection, move anchor + cursor together
      // If there's a multi-cell selection, first collapse to the edge in the direction
      if (!this.selection.isSingleCell) {
        if (dr < 0) { this.select(this.selection.startRow, this.selection.startCol); return; }
        if (dr > 0) { this.select(this.selection.endRow, this.selection.startCol); return; }
        if (dc < 0) { this.select(this.selection.startRow, this.selection.startCol); return; }
        if (dc > 0) { this.select(this.selection.startRow, this.selection.endCol); return; }
      }

      let newRow = this.activeRow + dr;
      let newCol = this.activeCol + dc;

      while (newRow >= 0 && newRow < sheet.rowCount && !sheet.isRowVisible(newRow)) newRow += dr || 1;
      while (newCol >= 0 && newCol < sheet.colCount && !sheet.isColVisible(newCol)) newCol += dc || 1;

      this.select(this._clampRow(newRow), this._clampCol(newCol));
    }
  }

  // ── Ctrl+Arrow: jump to edge of data region ──
  moveToEdge(dr, dc, extend = false) {
    const sheet = this._sheet();
    if (!sheet) return;
    // Start from cursor position when extending, anchor when not
    let row = extend ? this._cursorRow : this.activeRow;
    let col = extend ? this._cursorCol : this.activeCol;

    if (dc !== 0) {
      const startEmpty = !sheet.getCell(row, col) || sheet.getCell(row, col).isEmpty;
      col += dc;
      if (startEmpty) {
        while (col > 0 && col < sheet.colCount - 1) {
          if (sheet.getCell(row, col) && !sheet.getCell(row, col).isEmpty) break;
          col += dc;
        }
      } else {
        while (col > 0 && col < sheet.colCount - 1) {
          const next = sheet.getCell(row, col);
          if (!next || next.isEmpty) { col -= dc; break; }
          col += dc;
        }
      }
    }

    if (dr !== 0) {
      const startEmpty = !sheet.getCell(row, col) || sheet.getCell(row, col).isEmpty;
      row += dr;
      if (startEmpty) {
        while (row > 0 && row < sheet.rowCount - 1) {
          if (sheet.getCell(row, col) && !sheet.getCell(row, col).isEmpty) break;
          row += dr;
        }
      } else {
        while (row > 0 && row < sheet.rowCount - 1) {
          const next = sheet.getCell(row, col);
          if (!next || next.isEmpty) { row -= dr; break; }
          row += dr;
        }
      }
    }

    row = this._clampRow(row);
    col = this._clampCol(col);

    if (extend) {
      this.extendTo(row, col);
      this._scrollToCursor();
    } else {
      this.select(row, col);
    }
  }

  moveToStart(extend = false) {
    if (extend) { this.extendTo(0, 0); this._scrollToCursor(); }
    else this.select(0, 0);
  }

  moveToEnd(extend = false) {
    const sheet = this._sheet();
    if (!sheet) return;
    const used = sheet.getUsedRange();
    if (!used) return;
    if (extend) { this.extendTo(used.endRow, used.endCol); this._scrollToCursor(); }
    else this.select(used.endRow, used.endCol);
  }

  tabNext(reverse = false) {
    const sheet = this._sheet();
    if (!sheet) return;
    if (reverse) {
      if (this.activeCol > 0) this.select(this.activeRow, this.activeCol - 1);
      else if (this.activeRow > 0) this.select(this.activeRow - 1, sheet.colCount - 1);
    } else {
      if (this.activeCol < sheet.colCount - 1) this.select(this.activeRow, this.activeCol + 1);
      else if (this.activeRow < sheet.rowCount - 1) this.select(this.activeRow + 1, 0);
    }
  }

  pageMove(dr, extend = false) {
    const sheet = this._sheet();
    if (!sheet) return;
    const visibleRows = Math.floor(
      (this.spreadsheet.renderer.height - 60) / sheet.getRowHeight(this.activeRow)
    );
    this.move(dr * visibleRows, 0, extend);
  }
}
