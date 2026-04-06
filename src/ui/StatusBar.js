import { el } from '../utils/helpers.js';

/**
 * Status bar — sits inside the sheet tabs bar (right side).
 * Shows one metric at a time (Sum by default), click to cycle.
 * Click the label to see a dropdown with all metrics.
 */
export default class StatusBar {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.element = null;
    this._metrics = ['Sum', 'Avg', 'Count', 'Min', 'Max'];
    this._activeMetric = 'Sum';
    this._allValues = {};
  }

  init(container) {
    this.element = el('div', { className: 'sheets-status-content' });
    this.element.addEventListener('click', (e) => this._showPicker(e));

    const tabsBar = container.querySelector('.sheets-sheet-tabs');
    if (tabsBar) {
      tabsBar.appendChild(this.element);
    } else {
      const bar = el('div', { className: 'sheets-status-bar' });
      bar.appendChild(this.element);
      container.appendChild(bar);
    }

    this.spreadsheet.on('selectionChanged', () => this.update());
  }

  update() {
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    const sel = ss.selection;
    if (!sheet || !sel || !this.element) return;

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

    if (nums.length === 0 && count <= 1) {
      this.element.textContent = '';
      this._allValues = {};
      return;
    }

    const sum = nums.length > 0 ? nums.reduce((a, b) => a + b, 0) : 0;
    this._allValues = {
      Sum: nums.length > 0 ? this._fmt(sum) : '-',
      Avg: nums.length > 0 ? this._fmt(sum / nums.length) : '-',
      Count: String(count),
      Min: nums.length > 0 ? this._fmt(Math.min(...nums)) : '-',
      Max: nums.length > 0 ? this._fmt(Math.max(...nums)) : '-',
    };

    const val = this._allValues[this._activeMetric] || '-';
    this.element.textContent = `${this._activeMetric}: ${val}`;
  }

  _showPicker(e) {
    if (Object.keys(this._allValues).length === 0) return;

    // Remove existing picker
    const existing = document.querySelector('.sheets-status-picker');
    if (existing) { existing.remove(); return; }

    const picker = el('div', { className: 'sheets-status-picker' });
    const rect = this.element.getBoundingClientRect();
    picker.style.position = 'fixed';
    picker.style.left = rect.left + 'px';
    picker.style.bottom = (window.innerHeight - rect.top + 4) + 'px';

    for (const metric of this._metrics) {
      const val = this._allValues[metric] || '-';
      const item = el('div', {
        className: 'sheets-status-picker-item' + (metric === this._activeMetric ? ' active' : ''),
      });
      item.innerHTML = `<span class="sheets-sp-label">${metric}</span><span class="sheets-sp-value">${val}</span>`;
      item.addEventListener('click', (ev) => {
        ev.stopPropagation();
        this._activeMetric = metric;
        this.update();
        picker.remove();
      });
      picker.appendChild(item);
    }

    document.body.appendChild(picker);

    const close = (ev) => {
      if (!picker.contains(ev.target) && ev.target !== this.element) {
        picker.remove();
        document.removeEventListener('mousedown', close);
      }
    };
    setTimeout(() => document.addEventListener('mousedown', close), 0);
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
