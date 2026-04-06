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

    this.element.innerHTML = `
      <div class="sheets-fr-row">
        <input type="text" class="sheets-fr-input" placeholder="Find" autofocus />
        <span class="sheets-fr-count"></span>
        <button class="sheets-fr-btn" title="Previous (Shift+Enter)">▲</button>
        <button class="sheets-fr-btn" title="Next (Enter)">▼</button>
        <button class="sheets-fr-btn sheets-fr-close" title="Close (Escape)">✕</button>
      </div>
      ${showReplace ? `
      <div class="sheets-fr-row">
        <input type="text" class="sheets-fr-replace-input" placeholder="Replace with" />
        <button class="sheets-fr-btn sheets-fr-replace-btn">Replace</button>
        <button class="sheets-fr-btn sheets-fr-replace-all-btn">All</button>
      </div>` : ''}
      <div class="sheets-fr-options">
        <label><input type="checkbox" class="sheets-fr-case" /> Match case</label>
        <label><input type="checkbox" class="sheets-fr-regex" /> Regex</label>
        <label><input type="checkbox" class="sheets-fr-entire" /> Entire cell</label>
        <select class="sheets-fr-scope">
          <option value="sheet">This sheet</option>
          <option value="all">All sheets</option>
        </select>
      </div>
    `;

    this._findInput = this.element.querySelector('.sheets-fr-input');
    this._replaceInput = this.element.querySelector('.sheets-fr-replace-input');
    this._countLabel = this.element.querySelector('.sheets-fr-count');
    this._caseCheck = this.element.querySelector('.sheets-fr-case');
    this._regexCheck = this.element.querySelector('.sheets-fr-regex');
    this._entireCheck = this.element.querySelector('.sheets-fr-entire');
    this._scopeSelect = this.element.querySelector('.sheets-fr-scope');

    const btns = this.element.querySelectorAll('.sheets-fr-btn');
    btns[0].addEventListener('click', () => this._findPrev());
    btns[1].addEventListener('click', () => this._findNext());
    btns[2].addEventListener('click', () => this.hide());

    if (showReplace) {
      this.element.querySelector('.sheets-fr-replace-btn').addEventListener('click', () => this._replace());
      this.element.querySelector('.sheets-fr-replace-all-btn').addEventListener('click', () => this._replaceAll());
      this._replaceInput.addEventListener('keydown', (e) => {
        e.stopPropagation();
        if (e.key === 'Escape') { e.preventDefault(); this.hide(); }
      });
    }

    this._findInput.addEventListener('input', () => this._search());
    this._findInput.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') {
        e.preventDefault();
        if (e.shiftKey) this._findPrev(); else this._findNext();
      }
      if (e.key === 'Escape') { e.preventDefault(); this.hide(); }
    });

    // Options trigger re-search
    this._caseCheck.addEventListener('change', () => this._search());
    this._regexCheck.addEventListener('change', () => this._search());
    this._entireCheck.addEventListener('change', () => this._search());
    this._scopeSelect.addEventListener('change', () => this._search());

    this.container.appendChild(this.element);
    this._findInput.focus();

    // Pre-fill with current selection text
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    if (sheet) {
      const cell = sheet.getCell(ss.activeRow, ss.activeCol);
      if (cell && cell.displayValue && !cell.displayValue.startsWith('=')) {
        this._findInput.value = cell.displayValue;
        this._search();
      }
    }
  }

  hide() {
    if (this.element) {
      this.element.remove();
      this.element = null;
    }
    this.visible = false;
    this._matches = [];
    this._matchIndex = -1;
    // Refocus container
    if (this.spreadsheet.container) this.spreadsheet.container.focus();
  }

  _search() {
    const query = this._findInput.value;
    this._matches = [];
    this._matchIndex = -1;

    if (!query) {
      this._countLabel.textContent = '';
      return;
    }

    const matchCase = this._caseCheck.checked;
    const useRegex = this._regexCheck.checked;
    const entireCell = this._entireCheck.checked;
    const scope = this._scopeSelect.value;

    let re;
    try {
      if (useRegex) {
        re = new RegExp(query, matchCase ? '' : 'i');
      } else {
        const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        re = entireCell
          ? new RegExp('^' + escaped + '$', matchCase ? '' : 'i')
          : new RegExp(escaped, matchCase ? '' : 'i');
      }
    } catch (e) {
      this._countLabel.textContent = 'Invalid regex';
      return;
    }

    const sheets = scope === 'all' ? this.spreadsheet.sheets : [this.spreadsheet.activeSheet];

    for (const sheet of sheets) {
      if (!sheet) continue;
      for (const [key, cell] of sheet.cells) {
        const display = cell.displayValue;
        if (display && re.test(display)) {
          const { row, col } = keyToRC(key);
          this._matches.push({ row, col, sheetId: sheet.id, sheetName: sheet.name });
        }
      }
    }

    this._countLabel.textContent = this._matches.length > 0
      ? `${this._matches.length} found`
      : 'No results';

    if (this._matches.length > 0) {
      this._matchIndex = 0;
      this._goToMatch();
    }
  }

  _findNext() {
    if (this._matches.length === 0) { this._search(); return; }
    this._matchIndex = (this._matchIndex + 1) % this._matches.length;
    this._goToMatch();
  }

  _findPrev() {
    if (this._matches.length === 0) { this._search(); return; }
    this._matchIndex = (this._matchIndex - 1 + this._matches.length) % this._matches.length;
    this._goToMatch();
  }

  _goToMatch() {
    const match = this._matches[this._matchIndex];
    if (!match) return;

    const ss = this.spreadsheet;

    // Switch sheet if needed
    if (match.sheetId !== ss.activeSheet?.id) {
      ss.setActiveSheet(match.sheetId);
    }

    ss.selectionManager.select(match.row, match.col);
    ss.renderer.ensureCellVisible(match.row, match.col);
    ss.render();
    this._countLabel.textContent = `${this._matchIndex + 1} / ${this._matches.length}`;
  }

  _replace() {
    if (this._matchIndex < 0 || !this._replaceInput) return;
    const match = this._matches[this._matchIndex];
    if (!match) return;

    const ss = this.spreadsheet;
    const sheet = ss.getSheetById(match.sheetId) || ss.activeSheet;
    if (!sheet) return;
    const cell = sheet.getCell(match.row, match.col);
    if (!cell) return;

    const query = this._findInput.value;
    const replacement = this._replaceInput.value;
    const matchCase = this._caseCheck.checked;
    const useRegex = this._regexCheck.checked;

    let re;
    if (useRegex) {
      re = new RegExp(query, matchCase ? 'g' : 'gi');
    } else {
      const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      re = new RegExp(escaped, matchCase ? 'g' : 'gi');
    }

    const newVal = cell.displayValue.replace(re, replacement);
    sheet.setCellValue(match.row, match.col, newVal);

    ss.recalculate();
    ss.render();
    this._search(); // Re-search to update matches
  }

  _replaceAll() {
    if (!this._replaceInput) return;
    const query = this._findInput.value;
    const replacement = this._replaceInput.value;
    if (!query) return;

    const matchCase = this._caseCheck.checked;
    const useRegex = this._regexCheck.checked;

    let re;
    try {
      if (useRegex) {
        re = new RegExp(query, matchCase ? 'g' : 'gi');
      } else {
        const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        re = new RegExp(escaped, matchCase ? 'g' : 'gi');
      }
    } catch (e) { return; }

    const ss = this.spreadsheet;
    let count = 0;

    for (const match of [...this._matches]) {
      const sheet = ss.getSheetById(match.sheetId) || ss.activeSheet;
      if (!sheet) continue;
      const cell = sheet.getCell(match.row, match.col);
      if (!cell) continue;
      const newVal = cell.displayValue.replace(re, replacement);
      if (newVal !== cell.displayValue) {
        sheet.setCellValue(match.row, match.col, newVal);
        count++;
      }
    }

    ss.recalculate();
    ss.render();
    this._countLabel.textContent = `Replaced ${count}`;
    this._matches = [];
    this._matchIndex = -1;
  }

  destroy() {
    this.hide();
  }
}
