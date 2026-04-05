import { el } from '../utils/helpers.js';
import { cellKey, keyToRC } from '../utils/helpers.js';

export default class FindReplace {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.element = null;
    this.visible = false;
    this._matches = [];
    this._matchIndex = -1;
  }

  init(container) {
    this.container = container;
  }

  show(showReplace = false) {
    this.hide();
    this.visible = true;

    this.element = el('div', { className: 'sheets-find-replace' });

    const findRow = el('div', { className: 'sheets-find-row' });
    const findInput = el('input', {
      type: 'text',
      className: 'sheets-find-input',
      placeholder: 'Find',
    });
    this._findInput = findInput;

    const findPrev = el('button', { className: 'sheets-find-btn', onClick: () => this._findPrev() }, '▲');
    const findNext = el('button', { className: 'sheets-find-btn', onClick: () => this._findNext() }, '▼');
    const closeBtn = el('button', { className: 'sheets-find-btn', onClick: () => this.hide() }, '✕');
    this._countLabel = el('span', { className: 'sheets-find-count' }, '');

    findInput.addEventListener('input', () => this._search());
    findInput.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') {
        e.preventDefault();
        if (e.shiftKey) this._findPrev(); else this._findNext();
      }
      if (e.key === 'Escape') { e.preventDefault(); this.hide(); }
    });

    findRow.appendChild(findInput);
    findRow.appendChild(this._countLabel);
    findRow.appendChild(findPrev);
    findRow.appendChild(findNext);
    findRow.appendChild(closeBtn);
    this.element.appendChild(findRow);

    if (showReplace) {
      const replaceRow = el('div', { className: 'sheets-find-row' });
      const replaceInput = el('input', {
        type: 'text',
        className: 'sheets-find-input',
        placeholder: 'Replace with',
      });
      this._replaceInput = replaceInput;

      replaceInput.addEventListener('keydown', (e) => {
        e.stopPropagation();
        if (e.key === 'Escape') { e.preventDefault(); this.hide(); }
      });

      const replaceBtn = el('button', { className: 'sheets-find-btn', onClick: () => this._replace() }, 'Replace');
      const replaceAllBtn = el('button', { className: 'sheets-find-btn', onClick: () => this._replaceAll() }, 'All');

      replaceRow.appendChild(replaceInput);
      replaceRow.appendChild(replaceBtn);
      replaceRow.appendChild(replaceAllBtn);
      this.element.appendChild(replaceRow);
    }

    this.container.appendChild(this.element);
    findInput.focus();
  }

  hide() {
    if (this.element) {
      this.element.remove();
      this.element = null;
    }
    this.visible = false;
    this._matches = [];
    this._matchIndex = -1;
  }

  _search() {
    const query = this._findInput.value.toLowerCase();
    this._matches = [];
    this._matchIndex = -1;

    if (!query) {
      this._countLabel.textContent = '';
      return;
    }

    const sheet = this.spreadsheet.activeSheet;
    for (const [key, cell] of sheet.cells) {
      const display = cell.displayValue.toLowerCase();
      if (display.includes(query)) {
        const { row, col } = keyToRC(key);
        this._matches.push({ row, col });
      }
    }

    this._countLabel.textContent = this._matches.length + ' found';
    if (this._matches.length > 0) {
      this._matchIndex = 0;
      this._goToMatch();
    }
  }

  _findNext() {
    if (this._matches.length === 0) return;
    this._matchIndex = (this._matchIndex + 1) % this._matches.length;
    this._goToMatch();
  }

  _findPrev() {
    if (this._matches.length === 0) return;
    this._matchIndex = (this._matchIndex - 1 + this._matches.length) % this._matches.length;
    this._goToMatch();
  }

  _goToMatch() {
    const match = this._matches[this._matchIndex];
    if (!match) return;
    this.spreadsheet.selectionManager.select(match.row, match.col);
    this.spreadsheet.renderer.ensureCellVisible(match.row, match.col);
    this.spreadsheet.render();
    this._countLabel.textContent = `${this._matchIndex + 1} / ${this._matches.length}`;
  }

  _replace() {
    if (this._matchIndex < 0 || !this._replaceInput) return;
    const match = this._matches[this._matchIndex];
    if (!match) return;

    const sheet = this.spreadsheet.activeSheet;
    const cell = sheet.getCell(match.row, match.col);
    if (!cell) return;

    const query = this._findInput.value;
    const replacement = this._replaceInput.value;
    const newVal = cell.displayValue.split(query).join(replacement);
    sheet.setCellValue(match.row, match.col, newVal);

    this._search();
    this.spreadsheet.recalculate();
    this.spreadsheet.render();
  }

  _replaceAll() {
    if (!this._replaceInput) return;
    const query = this._findInput.value;
    const replacement = this._replaceInput.value;
    if (!query) return;

    const sheet = this.spreadsheet.activeSheet;
    for (const match of this._matches) {
      const cell = sheet.getCell(match.row, match.col);
      if (!cell) continue;
      const newVal = cell.displayValue.split(query).join(replacement);
      sheet.setCellValue(match.row, match.col, newVal);
    }

    this._search();
    this.spreadsheet.recalculate();
    this.spreadsheet.render();
  }

  destroy() {
    this.hide();
  }
}
