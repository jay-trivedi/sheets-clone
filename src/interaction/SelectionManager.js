import CellRange from '../core/CellRange.js';

export default class SelectionManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.activeRow = 0;
    this.activeCol = 0;
    this.selection = CellRange.fromCell(0, 0);
    this.multiSelections = [];
  }

  get range() { return this.selection; }

  _sheet() { return this.spreadsheet.activeSheet; }

  select(row, col, extend = false) {
    const sheet = this._sheet();
    if (!sheet) return;
    row = Math.max(0, Math.min(row, sheet.rowCount - 1));
    col = Math.max(0, Math.min(col, sheet.colCount - 1));

    const merge = sheet.getMergeAt(row, col);
    if (merge) {
      row = merge.startRow;
      col = merge.startCol;
    }

    if (extend) {
      this.selection = new CellRange(this.activeRow, this.activeCol, row, col);
    } else {
      this.activeRow = row;
      this.activeCol = col;
      this.selection = CellRange.fromCell(row, col);
    }

    this.spreadsheet.emit('selectionChanged', this.selection);
  }

  selectRange(range) {
    this.selection = range;
    this.activeRow = range.startRow;
    this.activeCol = range.startCol;
    this.spreadsheet.emit('selectionChanged', this.selection);
  }

  selectAll() {
    const sheet = this._sheet();
    if (!sheet) return;
    this.selection = new CellRange(0, 0, sheet.rowCount - 1, sheet.colCount - 1);
    this.spreadsheet.emit('selectionChanged', this.selection);
  }

  selectRow(row) {
    const sheet = this._sheet();
    if (!sheet) return;
    this.activeRow = row;
    this.activeCol = 0;
    this.selection = new CellRange(row, 0, row, sheet.colCount - 1);
    this.spreadsheet.emit('selectionChanged', this.selection);
  }

  selectCol(col) {
    const sheet = this._sheet();
    if (!sheet) return;
    this.activeRow = 0;
    this.activeCol = col;
    this.selection = new CellRange(0, col, sheet.rowCount - 1, col);
    this.spreadsheet.emit('selectionChanged', this.selection);
  }

  extendTo(row, col) {
    const sheet = this._sheet();
    if (!sheet) return;
    row = Math.max(0, Math.min(row, sheet.rowCount - 1));
    col = Math.max(0, Math.min(col, sheet.colCount - 1));
    this.selection = new CellRange(this.activeRow, this.activeCol, row, col);
    this.spreadsheet.emit('selectionChanged', this.selection);
  }

  move(dr, dc, extend = false) {
    const sheet = this._sheet();
    if (!sheet) return;
    let newRow = this.activeRow + dr;
    let newCol = this.activeCol + dc;

    while (newRow >= 0 && newRow < sheet.rowCount && !sheet.isRowVisible(newRow)) newRow += dr || 1;
    while (newCol >= 0 && newCol < sheet.colCount && !sheet.isColVisible(newCol)) newCol += dc || 1;

    newRow = Math.max(0, Math.min(newRow, sheet.rowCount - 1));
    newCol = Math.max(0, Math.min(newCol, sheet.colCount - 1));

    if (extend) {
      this.extendTo(newRow, newCol);
    } else {
      this.select(newRow, newCol);
    }
  }

  moveToEdge(dr, dc, extend = false) {
    const sheet = this._sheet();
    if (!sheet) return;
    let row = this.activeRow;
    let col = this.activeCol;

    if (dc !== 0) {
      const startEmpty = !sheet.getCell(row, col) || sheet.getCell(row, col).isEmpty;
      col += dc;
      if (startEmpty) {
        while (col > 0 && col < sheet.colCount - 1) {
          const cell = sheet.getCell(row, col);
          if (cell && !cell.isEmpty) break;
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
          const cell = sheet.getCell(row, col);
          if (cell && !cell.isEmpty) break;
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

    row = Math.max(0, Math.min(row, sheet.rowCount - 1));
    col = Math.max(0, Math.min(col, sheet.colCount - 1));

    if (extend) {
      this.extendTo(row, col);
    } else {
      this.select(row, col);
    }
  }

  moveToStart(extend = false) {
    if (extend) {
      this.extendTo(0, 0);
    } else {
      this.select(0, 0);
    }
  }

  moveToEnd(extend = false) {
    const sheet = this._sheet();
    if (!sheet) return;
    const used = sheet.getUsedRange();
    if (!used) return;
    if (extend) {
      this.extendTo(used.endRow, used.endCol);
    } else {
      this.select(used.endRow, used.endCol);
    }
  }

  tabNext(reverse = false) {
    const sheet = this._sheet();
    if (!sheet) return;
    if (reverse) {
      if (this.activeCol > 0) {
        this.select(this.activeRow, this.activeCol - 1);
      } else if (this.activeRow > 0) {
        this.select(this.activeRow - 1, sheet.colCount - 1);
      }
    } else {
      if (this.activeCol < sheet.colCount - 1) {
        this.select(this.activeRow, this.activeCol + 1);
      } else if (this.activeRow < sheet.rowCount - 1) {
        this.select(this.activeRow + 1, 0);
      }
    }
  }

  pageMove(dr) {
    const sheet = this._sheet();
    if (!sheet) return;
    const visibleRows = Math.floor(
      (this.spreadsheet.renderer.height - 60) / sheet.getRowHeight(this.activeRow)
    );
    this.move(dr * visibleRows, 0);
  }
}
