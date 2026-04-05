import { CELL_TYPE, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE, DEFAULT_TEXT_COLOR, DEFAULT_BG_COLOR, H_ALIGN, V_ALIGN, NUMBER_FORMATS } from '../utils/constants.js';

export class CellStyle {
  constructor(props = {}) {
    this.bold = props.bold || false;
    this.italic = props.italic || false;
    this.underline = props.underline || false;
    this.strikethrough = props.strikethrough || false;
    this.fontFamily = props.fontFamily || DEFAULT_FONT_FAMILY;
    this.fontSize = props.fontSize || DEFAULT_FONT_SIZE;
    this.textColor = props.textColor || DEFAULT_TEXT_COLOR;
    this.bgColor = props.bgColor || null;
    this.hAlign = props.hAlign || null; // null = auto (left for text, right for numbers)
    this.vAlign = props.vAlign || V_ALIGN.BOTTOM;
    this.wrap = props.wrap || false;
    this.numberFormat = props.numberFormat || NUMBER_FORMATS.GENERAL;
    this.borderTop = props.borderTop || null;
    this.borderRight = props.borderRight || null;
    this.borderBottom = props.borderBottom || null;
    this.borderLeft = props.borderLeft || null;
    this.rotation = props.rotation || 0;
    this.indent = props.indent || 0;
  }

  clone() {
    return new CellStyle(this);
  }

  toJSON() {
    const out = {};
    if (this.bold) out.bold = true;
    if (this.italic) out.italic = true;
    if (this.underline) out.underline = true;
    if (this.strikethrough) out.strikethrough = true;
    if (this.fontFamily !== DEFAULT_FONT_FAMILY) out.fontFamily = this.fontFamily;
    if (this.fontSize !== DEFAULT_FONT_SIZE) out.fontSize = this.fontSize;
    if (this.textColor !== DEFAULT_TEXT_COLOR) out.textColor = this.textColor;
    if (this.bgColor) out.bgColor = this.bgColor;
    if (this.hAlign) out.hAlign = this.hAlign;
    if (this.vAlign !== V_ALIGN.BOTTOM) out.vAlign = this.vAlign;
    if (this.wrap) out.wrap = true;
    if (this.numberFormat !== NUMBER_FORMATS.GENERAL) out.numberFormat = this.numberFormat;
    if (this.borderTop) out.borderTop = this.borderTop;
    if (this.borderRight) out.borderRight = this.borderRight;
    if (this.borderBottom) out.borderBottom = this.borderBottom;
    if (this.borderLeft) out.borderLeft = this.borderLeft;
    if (this.rotation) out.rotation = this.rotation;
    if (this.indent) out.indent = this.indent;
    return out;
  }

  getEffectiveHAlign(cellType) {
    if (this.hAlign) return this.hAlign;
    return (cellType === CELL_TYPE.NUMBER || cellType === CELL_TYPE.FORMULA)
      ? H_ALIGN.RIGHT : H_ALIGN.LEFT;
  }

  getFont() {
    let font = '';
    if (this.italic) font += 'italic ';
    if (this.bold) font += 'bold ';
    font += this.fontSize + 'pt ' + this.fontFamily;
    return font;
  }
}

export default class Cell {
  constructor() {
    this.rawValue = null;
    this.computedValue = null;
    this.formula = null;
    this.type = CELL_TYPE.EMPTY;
    this.style = null;
    this.comment = null;
    this.validation = null;
    this.mergeParent = null;   // { row, col } if this cell is part of a merge (not the top-left)
    this.mergeSpan = null;     // { rows, cols } if this is the top-left cell of a merge
  }

  get displayValue() {
    if (this.type === CELL_TYPE.EMPTY) return '';
    // Formula cells always show computed value, never the raw formula text
    if (this.formula) {
      if (this.computedValue === null || this.computedValue === undefined) return '';
      return String(this.computedValue);
    }
    const val = this.computedValue !== null ? this.computedValue : this.rawValue;
    if (val === null || val === undefined) return '';
    return String(val);
  }

  get effectiveType() {
    if (this.formula) {
      if (this.computedValue === null || this.computedValue === undefined) return CELL_TYPE.EMPTY;
      if (typeof this.computedValue === 'number') return CELL_TYPE.NUMBER;
      if (typeof this.computedValue === 'boolean') return CELL_TYPE.BOOLEAN;
      if (typeof this.computedValue === 'string' && this.computedValue.startsWith('#')) return CELL_TYPE.ERROR;
      return CELL_TYPE.STRING;
    }
    return this.type;
  }

  get isEmpty() {
    return this.type === CELL_TYPE.EMPTY && !this.formula;
  }

  getStyle() {
    return this.style || new CellStyle();
  }

  setStyle(props) {
    if (!this.style) this.style = new CellStyle();
    Object.assign(this.style, props);
  }

  setValue(value) {
    this.formula = null;
    if (value === null || value === undefined || value === '') {
      this.rawValue = null;
      this.computedValue = null;
      this.type = CELL_TYPE.EMPTY;
    } else if (typeof value === 'number') {
      this.rawValue = value;
      this.computedValue = value;
      this.type = CELL_TYPE.NUMBER;
    } else if (typeof value === 'boolean') {
      this.rawValue = value;
      this.computedValue = value;
      this.type = CELL_TYPE.BOOLEAN;
    } else {
      const str = String(value);
      if (str.startsWith('=')) {
        this.formula = str;
        this.rawValue = str;
        this.type = CELL_TYPE.FORMULA;
      } else {
        const num = Number(str);
        if (str.trim() !== '' && !isNaN(num)) {
          this.rawValue = str;
          this.computedValue = num;
          this.type = CELL_TYPE.NUMBER;
        } else if (str.toUpperCase() === 'TRUE') {
          this.rawValue = str;
          this.computedValue = true;
          this.type = CELL_TYPE.BOOLEAN;
        } else if (str.toUpperCase() === 'FALSE') {
          this.rawValue = str;
          this.computedValue = false;
          this.type = CELL_TYPE.BOOLEAN;
        } else {
          this.rawValue = str;
          this.computedValue = str;
          this.type = CELL_TYPE.STRING;
        }
      }
    }
  }

  setFormula(formula) {
    this.formula = formula.startsWith('=') ? formula : '=' + formula;
    this.rawValue = this.formula;
    this.type = CELL_TYPE.FORMULA;
  }

  clear() {
    this.rawValue = null;
    this.computedValue = null;
    this.formula = null;
    this.type = CELL_TYPE.EMPTY;
  }

  clearContent() {
    this.rawValue = null;
    this.computedValue = null;
    this.formula = null;
    this.type = CELL_TYPE.EMPTY;
  }

  clearFormat() {
    this.style = null;
  }

  clone() {
    const cell = new Cell();
    cell.rawValue = this.rawValue;
    cell.computedValue = this.computedValue;
    cell.formula = this.formula;
    cell.type = this.type;
    cell.style = this.style ? this.style.clone() : null;
    cell.comment = this.comment;
    cell.validation = this.validation;
    cell.mergeParent = this.mergeParent ? { ...this.mergeParent } : null;
    cell.mergeSpan = this.mergeSpan ? { ...this.mergeSpan } : null;
    return cell;
  }

  toJSON() {
    const out = {};
    if (this.formula) {
      out.formula = this.formula;
    } else if (this.rawValue !== null) {
      out.value = this.rawValue;
    }
    if (this.style) {
      const styleJSON = this.style.toJSON();
      if (Object.keys(styleJSON).length > 0) out.style = styleJSON;
    }
    if (this.comment) out.comment = this.comment;
    if (this.mergeSpan) out.mergeSpan = this.mergeSpan;
    return out;
  }

  static fromJSON(data) {
    const cell = new Cell();
    if (data.formula) {
      cell.setFormula(data.formula);
    } else if (data.value !== undefined) {
      cell.setValue(data.value);
    }
    if (data.style) cell.style = new CellStyle(data.style);
    if (data.comment) cell.comment = data.comment;
    if (data.mergeSpan) cell.mergeSpan = data.mergeSpan;
    return cell;
  }
}
