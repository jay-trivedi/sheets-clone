import { el, indexToCol, parseCellRef } from '../utils/helpers.js';

export default class FormulaBar {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.element = null;
    this.nameBox = null;
    this.input = null;
    this._editing = false;
  }

  init(container) {
    this.element = el('div', { className: 'sheets-formula-bar' });

    // Name box (cell reference)
    this.nameBox = el('input', {
      className: 'sheets-name-box',
      type: 'text',
      value: 'A1',
    });
    this.nameBox.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        e.stopPropagation(); // Prevent KeyboardHandler from seeing Enter
        this._navigateTo(this.nameBox.value);
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        this.update();
        if (this.spreadsheet.container) this.spreadsheet.container.focus();
      }
    });

    // fx label
    const fxLabel = el('div', { className: 'sheets-fx-label' }, 'fx');

    // Formula input
    this.input = el('input', {
      className: 'sheets-formula-input',
      type: 'text',
    });

    this.input.addEventListener('focus', () => {
      if (!this._editing) {
        const ss = this.spreadsheet;
        const sheet = ss.activeSheet;
        if (!sheet) return;
        const cell = sheet.getCell(ss.selectionManager.activeRow, ss.selectionManager.activeCol);
        const val = cell ? (cell.formula || cell.displayValue) : '';
        if (ss.editor) {
          ss.editor.begin(val, true);
        }
      }
    });

    this.input.addEventListener('input', () => {
      if (this.spreadsheet.editor && this.spreadsheet.editor.isActive) {
        this.spreadsheet.editor.textarea.value = this.input.value;
      }
    });

    this.input.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') {
        if (this.spreadsheet.editor && this.spreadsheet.editor.isActive) {
          this.spreadsheet.editor.textarea.value = this.input.value;
          this.spreadsheet.editor.commit();
        }
        this.spreadsheet.container.focus();
        e.preventDefault();
      } else if (e.key === 'Escape') {
        if (this.spreadsheet.editor && this.spreadsheet.editor.isActive) {
          this.spreadsheet.editor.cancel();
        }
        this.spreadsheet.container.focus();
        e.preventDefault();
      }
    });

    // Formula bar color overlay (for colored cell references)
    const inputWrap = el('div', { className: 'sheets-formula-input-wrap' });
    this.inputOverlay = el('div', { className: 'sheets-formula-input-overlay' });
    inputWrap.appendChild(this.input);
    inputWrap.appendChild(this.inputOverlay);

    this.element.appendChild(this.nameBox);
    this.element.appendChild(fxLabel);
    this.element.appendChild(inputWrap);
    container.appendChild(this.element);

    // Listen for selection changes
    this.spreadsheet.on('selectionChanged', () => this.update());
  }

  update() {
    if (!this.nameBox || !this.input) return;
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    if (!sheet) return;
    const sel = ss.selectionManager;
    const row = sel.activeRow;
    const col = sel.activeCol;

    this.nameBox.value = indexToCol(col) + (row + 1);

    const cell = sheet.getCell(row, col);
    if (cell) {
      this.input.value = cell.formula || cell.displayValue;
    } else {
      this.input.value = '';
    }
  }

  setValue(val) {
    if (!this.input) return;
    this.input.value = val;
    this.updateColorOverlay(val);
  }

  updateColorOverlay(text) {
    if (!this.inputOverlay) return;
    const fh = this.spreadsheet.formulaHelper;
    if (fh && text && text.startsWith('=')) {
      this.input.style.color = 'transparent';
      this.input.style.caretColor = '#000';
      this.inputOverlay.style.display = 'block';
      this.inputOverlay.innerHTML = fh.getColoredFormulaHTML(text);
    } else {
      this.input.style.color = '';
      this.input.style.caretColor = '';
      this.inputOverlay.style.display = 'none';
      this.inputOverlay.innerHTML = '';
    }
  }

  setEditing(editing) {
    this._editing = editing;
    if (!this.input) return;
    if (editing) {
      this.input.classList.add('editing');
    } else {
      this.input.classList.remove('editing');
      this.input.style.color = '';
      this.input.style.caretColor = '';
      if (this.inputOverlay) { this.inputOverlay.style.display = 'none'; this.inputOverlay.innerHTML = ''; }
      this.update();
    }
  }

  _navigateTo(ref) {
    const parsed = parseCellRef(ref.toUpperCase());
    if (parsed) {
      this.spreadsheet.selectionManager.select(parsed.row, parsed.col);
      this.spreadsheet.renderer.ensureCellVisible(parsed.row, parsed.col);
      this.spreadsheet.render();
      // Refocus container so keyboard shortcuts work
      if (this.spreadsheet.container) this.spreadsheet.container.focus();
    }
  }

  destroy() {
    if (this.element && this.element.parentElement) {
      this.element.parentElement.removeChild(this.element);
    }
  }
}
