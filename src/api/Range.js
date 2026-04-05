import { parseCellRef, cellRefToString, indexToCol, colToIndex } from '../utils/helpers.js';
import CellRange from '../core/CellRange.js';

/**
 * Google Sheets-compatible Range class.
 * All setters return `this` for method chaining.
 * Uses 0-based indexing internally, 1-based in public API (matching Apps Script).
 */
export default class Range {
  constructor(sheet, startRow, startCol, numRows = 1, numCols = 1) {
    this._sheet = sheet;
    this._r = startRow;   // 0-based
    this._c = startCol;   // 0-based
    this._nr = numRows;
    this._nc = numCols;
  }

  // ── Metadata ──

  getRow() { return this._r + 1; }
  getColumn() { return this._c + 1; }
  getLastRow() { return this._r + this._nr; }
  getLastColumn() { return this._c + this._nc; }
  getNumRows() { return this._nr; }
  getNumColumns() { return this._nc; }
  getHeight() { return this._nr; }
  getWidth() { return this._nc; }
  getSheet() { return this._sheet; }

  getA1Notation() {
    if (this._nr === 1 && this._nc === 1) {
      return cellRefToString(this._c, this._r);
    }
    return cellRefToString(this._c, this._r) + ':' +
           cellRefToString(this._c + this._nc - 1, this._r + this._nr - 1);
  }

  // ── Cell Navigation ──

  getCell(row, column) {
    return new Range(this._sheet, this._r + row - 1, this._c + column - 1, 1, 1);
  }

  offset(rowOffset, colOffset, numRows, numCols) {
    return new Range(
      this._sheet,
      this._r + rowOffset,
      this._c + colOffset,
      numRows !== undefined ? numRows : this._nr,
      numCols !== undefined ? numCols : this._nc,
    );
  }

  // ── Values (singular/plural pattern) ──

  getValue() {
    return this._sheet.getCellValue(this._r, this._c);
  }

  getValues() {
    const result = [];
    for (let r = 0; r < this._nr; r++) {
      const row = [];
      for (let c = 0; c < this._nc; c++) {
        row.push(this._sheet.getCellValue(this._r + r, this._c + c) ?? '');
      }
      result.push(row);
    }
    return result;
  }

  setValue(value) {
    this._forEachCell((r, c) => this._sheet.setCellValue(r, c, value));
    return this;
  }

  setValues(values) {
    for (let r = 0; r < this._nr; r++) {
      for (let c = 0; c < this._nc; c++) {
        this._sheet.setCellValue(this._r + r, this._c + c, values[r][c]);
      }
    }
    return this;
  }

  // ── Display Values ──

  getDisplayValue() {
    const cell = this._sheet.getCell(this._r, this._c);
    return cell ? cell.displayValue : '';
  }

  getDisplayValues() {
    return this._mapCells((r, c) => {
      const cell = this._sheet.getCell(r, c);
      return cell ? cell.displayValue : '';
    });
  }

  // ── Formulas ──

  getFormula() {
    const cell = this._sheet.getCell(this._r, this._c);
    return cell && cell.formula ? cell.formula : '';
  }

  getFormulas() {
    return this._mapCells((r, c) => {
      const cell = this._sheet.getCell(r, c);
      return cell && cell.formula ? cell.formula : '';
    });
  }

  setFormula(formula) {
    this._sheet.setCellFormula(this._r, this._c, formula);
    return this;
  }

  setFormulas(formulas) {
    for (let r = 0; r < this._nr; r++) {
      for (let c = 0; c < this._nc; c++) {
        if (formulas[r][c]) {
          this._sheet.setCellFormula(this._r + r, this._c + c, formulas[r][c]);
        }
      }
    }
    return this;
  }

  // ── Background ──

  getBackground() {
    const cell = this._sheet.getCell(this._r, this._c);
    return cell && cell.style && cell.style.bgColor ? cell.style.bgColor : null;
  }

  getBackgrounds() {
    return this._mapCells((r, c) => {
      const cell = this._sheet.getCell(r, c);
      return cell && cell.style && cell.style.bgColor ? cell.style.bgColor : null;
    });
  }

  setBackground(color) {
    this._forEachCell((r, c) => this._sheet.setCellStyle(r, c, { bgColor: color }));
    return this;
  }

  setBackgrounds(colors) {
    this._forEachCellWithIndex((r, c, ri, ci) => {
      this._sheet.setCellStyle(r, c, { bgColor: colors[ri][ci] });
    });
    return this;
  }

  setBackgroundRGB(red, green, blue) {
    const hex = '#' + [red, green, blue].map(v => v.toString(16).padStart(2, '0')).join('');
    return this.setBackground(hex);
  }

  // ── Font Color ──

  getFontColor() {
    const cell = this._sheet.getCell(this._r, this._c);
    return cell && cell.style ? cell.style.textColor || '#000000' : '#000000';
  }

  getFontColors() {
    return this._mapCells((r, c) => {
      const cell = this._sheet.getCell(r, c);
      return cell && cell.style ? cell.style.textColor || '#000000' : '#000000';
    });
  }

  setFontColor(color) {
    this._forEachCell((r, c) => this._sheet.setCellStyle(r, c, { textColor: color }));
    return this;
  }

  setFontColors(colors) {
    this._forEachCellWithIndex((r, c, ri, ci) => {
      this._sheet.setCellStyle(r, c, { textColor: colors[ri][ci] });
    });
    return this;
  }

  // ── Font Family ──

  getFontFamily() {
    const cell = this._sheet.getCell(this._r, this._c);
    return cell && cell.style ? cell.style.fontFamily || 'Arial' : 'Arial';
  }

  setFontFamily(fontFamily) {
    this._forEachCell((r, c) => this._sheet.setCellStyle(r, c, { fontFamily }));
    return this;
  }

  setFontFamilies(families) {
    this._forEachCellWithIndex((r, c, ri, ci) => {
      this._sheet.setCellStyle(r, c, { fontFamily: families[ri][ci] });
    });
    return this;
  }

  // ── Font Size ──

  getFontSize() {
    const cell = this._sheet.getCell(this._r, this._c);
    return cell && cell.style ? cell.style.fontSize || 10 : 10;
  }

  setFontSize(size) {
    this._forEachCell((r, c) => this._sheet.setCellStyle(r, c, { fontSize: size }));
    return this;
  }

  setFontSizes(sizes) {
    this._forEachCellWithIndex((r, c, ri, ci) => {
      this._sheet.setCellStyle(r, c, { fontSize: sizes[ri][ci] });
    });
    return this;
  }

  // ── Font Weight (bold) ──

  getFontWeight() {
    const cell = this._sheet.getCell(this._r, this._c);
    return cell && cell.style && cell.style.bold ? 'bold' : 'normal';
  }

  getFontWeights() {
    return this._mapCells((r, c) => {
      const cell = this._sheet.getCell(r, c);
      return cell && cell.style && cell.style.bold ? 'bold' : 'normal';
    });
  }

  setFontWeight(fontWeight) {
    this._forEachCell((r, c) => this._sheet.setCellStyle(r, c, { bold: fontWeight === 'bold' }));
    return this;
  }

  setFontWeights(weights) {
    this._forEachCellWithIndex((r, c, ri, ci) => {
      this._sheet.setCellStyle(r, c, { bold: weights[ri][ci] === 'bold' });
    });
    return this;
  }

  // ── Font Style (italic) ──

  getFontStyle() {
    const cell = this._sheet.getCell(this._r, this._c);
    return cell && cell.style && cell.style.italic ? 'italic' : 'normal';
  }

  setFontStyle(fontStyle) {
    this._forEachCell((r, c) => this._sheet.setCellStyle(r, c, { italic: fontStyle === 'italic' }));
    return this;
  }

  // ── Font Line (underline/strikethrough) ──

  getFontLine() {
    const cell = this._sheet.getCell(this._r, this._c);
    if (!cell || !cell.style) return 'none';
    if (cell.style.underline) return 'underline';
    if (cell.style.strikethrough) return 'line-through';
    return 'none';
  }

  setFontLine(fontLine) {
    this._forEachCell((r, c) => this._sheet.setCellStyle(r, c, {
      underline: fontLine === 'underline',
      strikethrough: fontLine === 'line-through',
    }));
    return this;
  }

  // ── Alignment ──

  getHorizontalAlignment() {
    const cell = this._sheet.getCell(this._r, this._c);
    return cell && cell.style && cell.style.hAlign ? cell.style.hAlign : 'left';
  }

  setHorizontalAlignment(alignment) {
    this._forEachCell((r, c) => this._sheet.setCellStyle(r, c, { hAlign: alignment }));
    return this;
  }

  setHorizontalAlignments(alignments) {
    this._forEachCellWithIndex((r, c, ri, ci) => {
      this._sheet.setCellStyle(r, c, { hAlign: alignments[ri][ci] });
    });
    return this;
  }

  getVerticalAlignment() {
    const cell = this._sheet.getCell(this._r, this._c);
    return cell && cell.style && cell.style.vAlign ? cell.style.vAlign : 'bottom';
  }

  setVerticalAlignment(alignment) {
    this._forEachCell((r, c) => this._sheet.setCellStyle(r, c, { vAlign: alignment }));
    return this;
  }

  // ── Wrapping ──

  getWrap() {
    const cell = this._sheet.getCell(this._r, this._c);
    return cell && cell.style ? !!cell.style.wrap : false;
  }

  setWrap(wrap) {
    this._forEachCell((r, c) => this._sheet.setCellStyle(r, c, { wrap }));
    return this;
  }

  setWraps(wraps) {
    this._forEachCellWithIndex((r, c, ri, ci) => {
      this._sheet.setCellStyle(r, c, { wrap: wraps[ri][ci] });
    });
    return this;
  }

  // ── Number Format ──

  getNumberFormat() {
    const cell = this._sheet.getCell(this._r, this._c);
    return cell && cell.style ? cell.style.numberFormat || 'General' : 'General';
  }

  setNumberFormat(format) {
    this._forEachCell((r, c) => this._sheet.setCellStyle(r, c, { numberFormat: format }));
    return this;
  }

  setNumberFormats(formats) {
    this._forEachCellWithIndex((r, c, ri, ci) => {
      this._sheet.setCellStyle(r, c, { numberFormat: formats[ri][ci] });
    });
    return this;
  }

  // ── Borders ──

  setBorder(top, left, bottom, right, vertical, horizontal, color, style) {
    const border = color ? { color, width: style === 'SOLID_THICK' ? 3 : style === 'SOLID_MEDIUM' ? 2 : 1 } : { color: '#000', width: 1 };
    this._forEachCell((r, c) => {
      const s = {};
      if (top === true && r === this._r) s.borderTop = border;
      if (top === false && r === this._r) s.borderTop = null;
      if (bottom === true && r === this._r + this._nr - 1) s.borderBottom = border;
      if (bottom === false && r === this._r + this._nr - 1) s.borderBottom = null;
      if (left === true && c === this._c) s.borderLeft = border;
      if (left === false && c === this._c) s.borderLeft = null;
      if (right === true && c === this._c + this._nc - 1) s.borderRight = border;
      if (right === false && c === this._c + this._nc - 1) s.borderRight = null;
      if (vertical === true) { s.borderLeft = border; s.borderRight = border; }
      if (horizontal === true) { s.borderTop = border; s.borderBottom = border; }
      if (Object.keys(s).length) this._sheet.setCellStyle(r, c, s);
    });
    return this;
  }

  // ── Notes ──

  getNote() {
    const cell = this._sheet.getCell(this._r, this._c);
    return cell ? cell.comment || '' : '';
  }

  getNotes() {
    return this._mapCells((r, c) => {
      const cell = this._sheet.getCell(r, c);
      return cell ? cell.comment || '' : '';
    });
  }

  setNote(note) {
    this._forEachCell((r, c) => {
      const cell = this._sheet.getOrCreateCell(r, c);
      cell.comment = note || null;
    });
    return this;
  }

  clearNote() {
    return this.setNote(null);
  }

  // ── Merge ──

  merge() {
    this._sheet.addMerge(new CellRange(this._r, this._c, this._r + this._nr - 1, this._c + this._nc - 1));
    return this;
  }

  breakApart() {
    this._sheet.removeMerge(new CellRange(this._r, this._c, this._r + this._nr - 1, this._c + this._nc - 1));
    return this;
  }

  isPartOfMerge() {
    const cell = this._sheet.getCell(this._r, this._c);
    return cell ? !!(cell.mergeParent || cell.mergeSpan) : false;
  }

  // ── Clear ──

  clear(options) {
    if (!options) {
      this._forEachCell((r, c) => this._sheet.clearCell(r, c));
    } else if (options.contentsOnly) {
      this._forEachCell((r, c) => {
        const cell = this._sheet.getCell(r, c);
        if (cell) cell.clearContent();
      });
    } else if (options.formatOnly) {
      this._forEachCell((r, c) => {
        const cell = this._sheet.getCell(r, c);
        if (cell) cell.clearFormat();
      });
    }
    return this;
  }

  clearContent() { return this.clear({ contentsOnly: true }); }
  clearFormat() { return this.clear({ formatOnly: true }); }

  // ── Status ──

  isBlank() {
    for (let r = 0; r < this._nr; r++) {
      for (let c = 0; c < this._nc; c++) {
        const cell = this._sheet.getCell(this._r + r, this._c + c);
        if (cell && !cell.isEmpty) return false;
      }
    }
    return true;
  }

  // ── Sort ──

  sort(sortSpec) {
    // sortSpec: { column: 1-based, ascending: bool } or just column number
    const col = typeof sortSpec === 'number' ? sortSpec : sortSpec.column;
    const asc = typeof sortSpec === 'number' ? true : sortSpec.ascending !== false;
    const colIdx = col - 1;

    const rows = [];
    for (let r = 0; r < this._nr; r++) {
      const rowData = [];
      for (let c = 0; c < this._nc; c++) {
        const cell = this._sheet.getCell(this._r + r, this._c + c);
        rowData.push(cell ? cell.clone() : null);
      }
      const sortVal = this._sheet.getCellValue(this._r + r, this._c + colIdx);
      rows.push({ data: rowData, key: sortVal });
    }

    rows.sort((a, b) => {
      const va = a.key, vb = b.key;
      if (va == null && vb == null) return 0;
      if (va == null) return 1;
      if (vb == null) return -1;
      return asc ? (va < vb ? -1 : va > vb ? 1 : 0) : (va > vb ? -1 : va < vb ? 1 : 0);
    });

    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].data.length; c++) {
        if (rows[r].data[c]) this._sheet.setCell(this._r + r, this._c + c, rows[r].data[c]);
        else this._sheet.clearCell(this._r + r, this._c + c);
      }
    }
    return this;
  }

  // ── Checkboxes ──

  insertCheckboxes() {
    this._forEachCell((r, c) => {
      if (!this._sheet.getCellValue(r, c)) this._sheet.setCellValue(r, c, false);
    });
    return this;
  }

  check() {
    this._forEachCell((r, c) => this._sheet.setCellValue(r, c, true));
    return this;
  }

  uncheck() {
    this._forEachCell((r, c) => this._sheet.setCellValue(r, c, false));
    return this;
  }

  isChecked() {
    const v = this.getValue();
    if (v === true) return true;
    if (v === false) return false;
    return null;
  }

  // ── Internal helpers ──

  _forEachCell(fn) {
    for (let r = this._r; r < this._r + this._nr; r++) {
      for (let c = this._c; c < this._c + this._nc; c++) {
        fn(r, c);
      }
    }
  }

  _forEachCellWithIndex(fn) {
    for (let ri = 0; ri < this._nr; ri++) {
      for (let ci = 0; ci < this._nc; ci++) {
        fn(this._r + ri, this._c + ci, ri, ci);
      }
    }
  }

  _mapCells(fn) {
    const result = [];
    for (let r = 0; r < this._nr; r++) {
      const row = [];
      for (let c = 0; c < this._nc; c++) {
        row.push(fn(this._r + r, this._c + c));
      }
      result.push(row);
    }
    return result;
  }

  // ── Static: parse A1 notation ──

  static fromA1Notation(sheet, notation) {
    const parts = notation.split('!');
    const rangeStr = parts.length > 1 ? parts[1] : parts[0];

    const rangeParts = rangeStr.split(':');
    if (rangeParts.length === 1) {
      const ref = parseCellRef(rangeParts[0]);
      if (!ref) return null;
      return new Range(sheet, ref.row, ref.col, 1, 1);
    }

    const start = parseCellRef(rangeParts[0]);
    const end = parseCellRef(rangeParts[1]);
    if (!start || !end) return null;

    const r1 = Math.min(start.row, end.row);
    const c1 = Math.min(start.col, end.col);
    const r2 = Math.max(start.row, end.row);
    const c2 = Math.max(start.col, end.col);

    return new Range(sheet, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
  }
}
