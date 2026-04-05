import { el } from '../utils/helpers.js';

export default class StatusBar {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.element = null;
  }

  init(container) {
    this.element = el('div', { className: 'sheets-status-bar' });
    this._content = el('div', { className: 'sheets-status-content' });
    this.element.appendChild(this._content);
    container.appendChild(this.element);

    this.spreadsheet.on('selectionChanged', () => this.update());
  }

  update() {
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    const sel = ss.selection;
    if (!sheet || !sel || !this._content) return;

    // Collect numeric values from selection
    const nums = [];
    let count = 0;
    for (let r = sel.startRow; r <= sel.endRow; r++) {
      for (let c = sel.startCol; c <= sel.endCol; c++) {
        const val = sheet.getCellValue(r, c);
        if (val !== null && val !== undefined && val !== '') {
          count++;
          const n = typeof val === 'number' ? val : parseFloat(val);
          if (!isNaN(n)) nums.push(n);
        }
      }
    }

    if (nums.length === 0 && count === 0) {
      this._content.textContent = '';
      return;
    }

    const parts = [];
    if (count > 1) parts.push(`Count: ${count}`);
    if (nums.length > 0) {
      const sum = nums.reduce((a, b) => a + b, 0);
      parts.push(`Sum: ${this._fmt(sum)}`);
      parts.push(`Avg: ${this._fmt(sum / nums.length)}`);
      if (nums.length > 1) {
        parts.push(`Min: ${this._fmt(Math.min(...nums))}`);
        parts.push(`Max: ${this._fmt(Math.max(...nums))}`);
      }
    }

    this._content.textContent = parts.join('    ');
  }

  _fmt(n) {
    if (Number.isInteger(n)) return n.toLocaleString();
    return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  destroy() {
    if (this.element && this.element.parentElement) {
      this.element.parentElement.removeChild(this.element);
    }
  }
}
