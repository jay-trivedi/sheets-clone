import { ROW_HEADER_WIDTH, COL_HEADER_HEIGHT, DEFAULT_FONT_FAMILY } from '../utils/constants.js';
import { el } from '../utils/helpers.js';

export default class Editor {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.isActive = false;
    this.editRow = -1;
    this.editCol = -1;
    this.textarea = null;
  }

  init(container) {
    this.textarea = el('textarea', {
      className: 'sheets-cell-editor',
      style: {
        position: 'absolute',
        display: 'none',
        zIndex: '10',
        border: '2px solid #1a73e8',
        outline: 'none',
        resize: 'none',
        overflow: 'hidden',
        padding: '2px 4px',
        fontFamily: DEFAULT_FONT_FAMILY,
        fontSize: '10pt',
        lineHeight: '1.2',
        background: '#fff',
        boxSizing: 'border-box',
        whiteSpace: 'pre',
        minWidth: '60px',
      },
    });

    this.textarea.addEventListener('input', () => this._onInput());
    this.textarea.addEventListener('keydown', (e) => this._onKeyDown(e));

    container.appendChild(this.textarea);
  }

  get isFormulaMode() {
    return this.isActive && this.textarea.value.startsWith('=');
  }

  begin(initialValue = '', cursorMode = false) {
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    if (!sheet) return;
    const sel = ss.selectionManager;
    const renderer = ss.renderer;

    this.editRow = sel.activeRow;
    this.editCol = sel.activeCol;
    this.isActive = true;

    const rect = renderer._getCellRect(sheet, this.editRow, this.editCol);
    const cell = sheet.getCell(this.editRow, this.editCol);
    let w = rect.w, h = rect.h;
    if (cell && cell.mergeSpan) {
      for (let c = this.editCol + 1; c < this.editCol + cell.mergeSpan.cols; c++) w += sheet.getColWidth(c);
      for (let r = this.editRow + 1; r < this.editRow + cell.mergeSpan.rows; r++) h += sheet.getRowHeight(r);
    }

    const style = cell ? cell.getStyle() : null;

    this.textarea.style.display = 'block';
    this.textarea.style.left = rect.x + 'px';
    this.textarea.style.top = rect.y + 'px';
    this.textarea.style.width = Math.max(w, 60) + 'px';
    this.textarea.style.height = h + 'px';
    this.textarea.style.minHeight = h + 'px';

    if (style) {
      this.textarea.style.fontFamily = style.fontFamily;
      this.textarea.style.fontSize = style.fontSize + 'pt';
      this.textarea.style.fontWeight = style.bold ? 'bold' : 'normal';
      this.textarea.style.fontStyle = style.italic ? 'italic' : 'normal';
      this.textarea.style.color = style.textColor;
      this.textarea.style.background = style.bgColor || '#fff';
    } else {
      this.textarea.style.fontWeight = 'normal';
      this.textarea.style.fontStyle = 'normal';
      this.textarea.style.color = '#000';
      this.textarea.style.background = '#fff';
    }

    this.textarea.value = initialValue;
    this.textarea.focus();

    if (cursorMode) {
      this.textarea.selectionStart = this.textarea.value.length;
      this.textarea.selectionEnd = this.textarea.value.length;
    } else {
      this.textarea.select();
    }

    if (ss.formulaBar) {
      ss.formulaBar.setValue(initialValue);
      ss.formulaBar.setEditing(true);
    }

    this._notifyFormulaHelper();
  }

  commit() {
    if (!this.isActive) return;

    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    if (!sheet) { this._close(); return; }
    const value = this.textarea.value;

    const oldCell = sheet.getCell(this.editRow, this.editCol);
    const oldValue = oldCell ? oldCell.rawValue : null;
    const oldFormula = oldCell ? oldCell.formula : null;

    sheet.setCellValue(this.editRow, this.editCol, value);

    ss.commandManager.execute({
      type: 'setCellValue',
      row: this.editRow,
      col: this.editCol,
      oldValue,
      oldFormula,
      newValue: value,
    });

    this._close();
    ss.recalculate();
    ss.render();
  }

  cancel() {
    if (!this.isActive) return;
    this._close();
    this.spreadsheet.render();
  }

  _close() {
    this.isActive = false;
    this.textarea.style.display = 'none';
    this.textarea.value = '';
    this.editRow = -1;
    this.editCol = -1;

    const ss = this.spreadsheet;
    if (ss.formulaBar) ss.formulaBar.setEditing(false);
    if (ss.formulaHelper) ss.formulaHelper.hide();

    // Return focus to container so keyboard shortcuts work
    if (ss.container) ss.container.focus();
  }

  _onInput() {
    const ctx = this.spreadsheet.renderer.ctx;
    if (ctx) {
      ctx.font = this.textarea.style.fontSize + ' ' + this.textarea.style.fontFamily;
      const textW = ctx.measureText(this.textarea.value + 'W').width;
      const minW = parseInt(this.textarea.style.width) || 60;
      this.textarea.style.width = Math.max(minW, textW + 20) + 'px';
    }

    this.textarea.style.height = 'auto';
    this.textarea.style.height = Math.max(parseInt(this.textarea.style.minHeight) || 25, this.textarea.scrollHeight) + 'px';

    if (this.spreadsheet.formulaBar) {
      this.spreadsheet.formulaBar.setValue(this.textarea.value);
    }

    this._notifyFormulaHelper();
  }

  _notifyFormulaHelper() {
    const fh = this.spreadsheet.formulaHelper;
    if (!fh) return;

    const text = this.textarea.value;
    const cursorPos = this.textarea.selectionStart;
    const taRect = this.textarea.getBoundingClientRect();
    const containerRect = this.textarea.parentElement.getBoundingClientRect();

    fh.onInput(text, cursorPos, {
      left: taRect.left - containerRect.left,
      top: taRect.top - containerRect.top,
      bottom: taRect.bottom - containerRect.top,
      width: taRect.width,
    });

    this.spreadsheet.renderer.requestRender();
  }

  _onKeyDown(e) {
    // IMPORTANT: Editor owns ALL keyboard input when active.
    // We stopPropagation so KeyboardHandler doesn't interfere.
    e.stopPropagation();

    const ss = this.spreadsheet;
    const sel = ss.selectionManager;
    const fh = ss.formulaHelper;

    // ── Formula helper gets first crack (autocomplete, point-mode arrows) ──
    if (fh && this.isFormulaMode) {
      const handled = fh.handlePointModeKey(e, this.textarea);
      if (handled) return;
    }

    // ── F4: cycle cell reference absolute/relative ($A$1 → A$1 → $A1 → A1) ──
    if (e.key === 'F4' && this.isFormulaMode) {
      e.preventDefault();
      this._cycleAbsoluteRef();
      return;
    }

    // ── Escape: cancel editing ──
    if (e.key === 'Escape') {
      e.preventDefault();
      this.cancel();
      return;
    }

    // ── Enter: commit and move down (Shift: up) ──
    if (e.key === 'Enter' && !e.altKey) {
      e.preventDefault();
      this.commit();
      sel.move(e.shiftKey ? -1 : 1, 0);
      ss.renderer.ensureCellVisible(sel.activeRow, sel.activeCol);
      ss.render();
      return;
    }

    // ── Tab: commit and move right (Shift: left) ──
    if (e.key === 'Tab') {
      e.preventDefault();
      this.commit();
      sel.tabNext(e.shiftKey);
      ss.renderer.ensureCellVisible(sel.activeRow, sel.activeCol);
      ss.render();
      return;
    }

    // All other keys: let textarea handle normally (typing, cursor movement,
    // Ctrl+Z native undo, Ctrl+A select all in textarea, etc.)
  }

  _cycleAbsoluteRef() {
    const ta = this.textarea;
    const text = ta.value;
    const cursor = ta.selectionStart;

    // Find the cell reference (or range) surrounding the cursor
    // Walk left to find start of ref token
    let start = cursor;
    while (start > 0 && /[A-Za-z0-9$:]/.test(text[start - 1])) start--;
    // Walk right to find end
    let end = cursor;
    while (end < text.length && /[A-Za-z0-9$:]/.test(text[end])) end++;

    const token = text.substring(start, end);
    if (!token) return;

    // Handle range (e.g. A1:B2) or single ref (e.g. A1)
    const cycled = token.split(':').map(ref => this._cycleOneRef(ref)).join(':');
    if (cycled === token) return;

    ta.value = text.substring(0, start) + cycled + text.substring(end);
    // Keep cursor within the new token
    const newEnd = start + cycled.length;
    ta.selectionStart = start;
    ta.selectionEnd = newEnd;
    ta.dispatchEvent(new Event('input'));
  }

  _cycleOneRef(ref) {
    // Parse: optional $ before col letters, optional $ before row digits
    const m = ref.match(/^(\$?)([A-Za-z]+)(\$?)(\d+)$/);
    if (!m) return ref;

    const colAbs = m[1] === '$';
    const col = m[2];
    const rowAbs = m[3] === '$';
    const row = m[4];

    // Cycle: A1 → $A$1 → A$1 → $A1 → A1
    if (!colAbs && !rowAbs) return '$' + col + '$' + row;   // A1 → $A$1
    if (colAbs && rowAbs)   return col + '$' + row;          // $A$1 → A$1
    if (!colAbs && rowAbs)  return '$' + col + row;          // A$1 → $A1
    return col + row;                                         // $A1 → A1
  }

  updatePosition() {
    if (!this.isActive) return;
    const sheet = this.spreadsheet.activeSheet;
    if (!sheet) return;
    const renderer = this.spreadsheet.renderer;
    const rect = renderer._getCellRect(sheet, this.editRow, this.editCol);
    this.textarea.style.left = rect.x + 'px';
    this.textarea.style.top = rect.y + 'px';
  }

  destroy() {
    if (this.textarea && this.textarea.parentElement) {
      this.textarea.parentElement.removeChild(this.textarea);
    }
  }
}
