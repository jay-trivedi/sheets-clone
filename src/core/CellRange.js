import { parseCellRef, cellRefToString, clamp } from '../utils/helpers.js';

export default class CellRange {
  constructor(startRow, startCol, endRow, endCol) {
    this.startRow = Math.min(startRow, endRow);
    this.startCol = Math.min(startCol, endCol);
    this.endRow = Math.max(startRow, endRow);
    this.endCol = Math.max(startCol, endCol);
  }

  get rowCount() { return this.endRow - this.startRow + 1; }
  get colCount() { return this.endCol - this.startCol + 1; }
  get cellCount() { return this.rowCount * this.colCount; }
  get isSingleCell() { return this.startRow === this.endRow && this.startCol === this.endCol; }

  contains(row, col) {
    return row >= this.startRow && row <= this.endRow &&
           col >= this.startCol && col <= this.endCol;
  }

  intersects(other) {
    return this.startRow <= other.endRow && this.endRow >= other.startRow &&
           this.startCol <= other.endCol && this.endCol >= other.startCol;
  }

  intersection(other) {
    if (!this.intersects(other)) return null;
    return new CellRange(
      Math.max(this.startRow, other.startRow),
      Math.max(this.startCol, other.startCol),
      Math.min(this.endRow, other.endRow),
      Math.min(this.endCol, other.endCol),
    );
  }

  union(other) {
    return new CellRange(
      Math.min(this.startRow, other.startRow),
      Math.min(this.startCol, other.startCol),
      Math.max(this.endRow, other.endRow),
      Math.max(this.endCol, other.endCol),
    );
  }

  equals(other) {
    return this.startRow === other.startRow && this.startCol === other.startCol &&
           this.endRow === other.endRow && this.endCol === other.endCol;
  }

  forEach(fn) {
    for (let r = this.startRow; r <= this.endRow; r++) {
      for (let c = this.startCol; c <= this.endCol; c++) {
        fn(r, c);
      }
    }
  }

  map(fn) {
    const result = [];
    for (let r = this.startRow; r <= this.endRow; r++) {
      const row = [];
      for (let c = this.startCol; c <= this.endCol; c++) {
        row.push(fn(r, c));
      }
      result.push(row);
    }
    return result;
  }

  *[Symbol.iterator]() {
    for (let r = this.startRow; r <= this.endRow; r++) {
      for (let c = this.startCol; c <= this.endCol; c++) {
        yield { row: r, col: c };
      }
    }
  }

  translate(dr, dc) {
    return new CellRange(
      this.startRow + dr, this.startCol + dc,
      this.endRow + dr, this.endCol + dc,
    );
  }

  expand(rows, cols) {
    return new CellRange(
      this.startRow - rows, this.startCol - cols,
      this.endRow + rows, this.endCol + cols,
    );
  }

  clamp(maxRow, maxCol) {
    return new CellRange(
      clamp(this.startRow, 0, maxRow),
      clamp(this.startCol, 0, maxCol),
      clamp(this.endRow, 0, maxRow),
      clamp(this.endCol, 0, maxCol),
    );
  }

  toString() {
    if (this.isSingleCell) {
      return cellRefToString(this.startCol, this.startRow);
    }
    return cellRefToString(this.startCol, this.startRow) + ':' +
           cellRefToString(this.endCol, this.endRow);
  }

  static fromString(ref) {
    const parts = ref.split(':');
    if (parts.length === 1) {
      const cell = parseCellRef(parts[0]);
      if (!cell) return null;
      return new CellRange(cell.row, cell.col, cell.row, cell.col);
    }
    const start = parseCellRef(parts[0]);
    const end = parseCellRef(parts[1]);
    if (!start || !end) return null;
    return new CellRange(start.row, start.col, end.row, end.col);
  }

  static fromCell(row, col) {
    return new CellRange(row, col, row, col);
  }

  static spanningCells(cells) {
    if (!cells.length) return null;
    let minR = Infinity, minC = Infinity, maxR = -Infinity, maxC = -Infinity;
    for (const { row, col } of cells) {
      if (row < minR) minR = row;
      if (col < minC) minC = col;
      if (row > maxR) maxR = row;
      if (col > maxC) maxC = col;
    }
    return new CellRange(minR, minC, maxR, maxC);
  }
}
