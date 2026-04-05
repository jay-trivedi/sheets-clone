import { ROW_HEADER_WIDTH, COL_HEADER_HEIGHT, DEFAULT_FONT_FAMILY } from '../utils/constants.js';
import { el, cellRefToString } from '../utils/helpers.js';

/*
 * Cell editing state machine (matches Google Sheets behavior):
 *
 * READY  → default, no editing. Arrow keys navigate. Typing starts ENTER mode.
 * ENTER  → new input (cell content cleared). For formulas, arrows → POINT mode.
 *          For non-formulas, arrows commit + navigate.
 * EDIT   → editing existing content with cursor. Arrows move text cursor.
 *          F2 toggles to POINT mode (for formulas).
 * POINT  → formula reference selection. Arrows insert cell refs.
 *          F2 toggles back to EDIT mode.
 */

export const MODE = {
  READY: 'ready',
  ENTER: 'enter',
  EDIT: 'edit',
  POINT: 'point',
};

export default class Editor {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.mode = MODE.READY;
    this.editRow = -1;
    this.editCol = -1;
    this.textarea = null;
    this._tabStartCol = -1; // column where Tab-based entry started
  }

  get isActive() {
    return this.mode !== MODE.READY;
  }

  get isFormulaMode() {
    return this.isActive && this.textarea.value.startsWith('=');
  }

  init(container) {
    // Color overlay — renders colored reference text on top of transparent textarea text
    this.colorOverlay = el('div', {
      className: 'sheets-cell-editor-overlay',
      style: {
        position: 'absolute',
        display: 'none',
        zIndex: '11',
        pointerEvents: 'none',
        padding: '2px 4px',
        border: '2px solid transparent',
        fontFamily: DEFAULT_FONT_FAMILY,
        fontSize: '10pt',
        lineHeight: '1.2',
        boxSizing: 'border-box',
        whiteSpace: 'pre',
        minWidth: '60px',
        overflow: 'hidden',
      },
    });
    container.appendChild(this.colorOverlay);

    this.textarea = el('textarea', {
      className: 'sheets-cell-editor',
      style: {
        position: 'absolute',
        display: 'none',
        zIndex: '12',
        border: '2px solid #3271ea',
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

  // ── Enter ENTER mode (typing replaces cell content) ──
  beginEnter(initialChar = '') {
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    if (!sheet) return;

    this.mode = MODE.ENTER; // set mode BEFORE _openEditor so _syncOverlay sees it
    this._openEditor(sheet, '');
    this.textarea.value = initialChar;
    this.textarea.selectionStart = initialChar.length;
    this.textarea.selectionEnd = initialChar.length;

    this._syncOverlay(); // re-sync after setting value
    this._notifyFormulaHelper();
  }

  // ── Enter EDIT mode (edit existing content with cursor) ──
  beginEdit(cursorAtEnd = true) {
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    if (!sheet) return;

    const cell = sheet.getCell(ss.selectionManager.activeRow, ss.selectionManager.activeCol);
    const val = cell ? (cell.formula || cell.displayValue) : '';

    this.mode = MODE.EDIT; // set mode BEFORE _openEditor so _syncOverlay sees it
    this._openEditor(sheet, val);

    if (cursorAtEnd) {
      this.textarea.selectionStart = this.textarea.value.length;
      this.textarea.selectionEnd = this.textarea.value.length;
    }

    this._notifyFormulaHelper();
  }

  // ── Legacy API: begin(value, cursorMode) for backward compat ──
  begin(initialValue = '', cursorMode = false) {
    if (cursorMode) {
      const ss = this.spreadsheet;
      const sheet = ss.activeSheet;
      if (!sheet) return;
      this.mode = MODE.EDIT;
      this._openEditor(sheet, initialValue);
      this.textarea.selectionStart = this.textarea.value.length;
      this.textarea.selectionEnd = this.textarea.value.length;
      this._notifyFormulaHelper();
    } else {
      // Typing starts ENTER mode
      this.beginEnter(initialValue);
    }
  }

  _openEditor(sheet, value) {
    const ss = this.spreadsheet;
    const sel = ss.selectionManager;

    this.editRow = sel.activeRow;
    this.editCol = sel.activeCol;

    const rect = ss.renderer._getCellRect(sheet, this.editRow, this.editCol);
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

    this.textarea.value = value;
    this.textarea.focus();

    // Show color overlay (synced position/size)
    this._syncOverlay();

    if (ss.formulaBar) {
      ss.formulaBar.setValue(value);
      ss.formulaBar.setEditing(true);
    }

    ss.emit('editingStarted', {
      cell: this.editRow >= 0 ? cellRefToString(this.editCol, this.editRow) : null,
      mode: this.mode,
      sheet: ss.activeSheet,
    });
  }

  // ── Switch to POINT mode (from ENTER or EDIT, for formulas) ──
  enterPointMode() {
    if (!this.isFormulaMode) return;
    this.mode = MODE.POINT;
    const fh = this.spreadsheet.formulaHelper;
    if (fh) {
      fh.pointMode = true;
      fh.pointAnchor = null;
      fh.pointCurrent = null;
    }
  }

  // ── Switch to EDIT mode (from POINT, toggle with F2) ──
  enterEditMode() {
    this.mode = MODE.EDIT;
    const fh = this.spreadsheet.formulaHelper;
    if (fh) {
      fh.pointMode = false;
      fh.pointAnchor = null;
      fh.pointCurrent = null;
    }
  }

  commit() {
    if (!this.isActive) return;

    const ss = this.spreadsheet;
    // Use the editing sheet, not the currently viewed sheet (may differ during cross-sheet formula editing)
    const editSheetId = ss._formulaEditSheetId || ss.activeSheet?.id;
    const sheet = ss.getSheetById(editSheetId) || ss.activeSheet;
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

    // Emit edit event (Google Sheets onEdit compatible)
    ss.emit('edit', {
      range: sheet.getRange(this.editRow + 1, this.editCol + 1),
      source: ss,
      value: value,
      oldValue: oldValue,
      sheet: sheet,
    });
  }

  cancel() {
    if (!this.isActive) return;
    this._close();
    this.spreadsheet.render();
  }

  _close() {
    this.mode = MODE.READY;
    this.textarea.style.display = 'none';
    this.colorOverlay.style.display = 'none';
    this.textarea.value = '';
    this.editRow = -1;
    this.editCol = -1;

    const ss = this.spreadsheet;

    // If we switched sheets during formula editing, switch back to the editing sheet
    if (ss._formulaEditSheetId) {
      const editSheet = ss.getSheetById(ss._formulaEditSheetId);
      if (editSheet) {
        ss.activeSheetIndex = ss.sheets.indexOf(editSheet);
        if (ss.sheetTabs) ss.sheetTabs.update();
      }
      ss._formulaEditSheetId = null;
    }

    if (ss.formulaBar) ss.formulaBar.setEditing(false);
    if (ss.formulaHelper) ss.formulaHelper.hide();

    ss.emit('editingStopped', {
      cell: this.editRow >= 0 ? cellRefToString(this.editCol, this.editRow) : null,
      committed: this.mode !== MODE.READY, // true if commit, false if cancel
    });

    if (ss.container) ss.container.focus();
  }

  _onInput() {
    // Auto-expand width
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
      this.spreadsheet.formulaBar.updateColorOverlay(this.textarea.value);
    }

    this._syncOverlay();
    this._notifyFormulaHelper();
  }

  _syncOverlay() {
    const ov = this.colorOverlay;
    const ta = this.textarea;
    // Match position and size exactly
    ov.style.left = ta.style.left;
    ov.style.top = ta.style.top;
    ov.style.width = ta.style.width;
    ov.style.height = ta.style.height;
    ov.style.minHeight = ta.style.minHeight;
    ov.style.fontFamily = ta.style.fontFamily;
    ov.style.fontSize = ta.style.fontSize;
    ov.style.fontWeight = ta.style.fontWeight;
    ov.style.fontStyle = ta.style.fontStyle;

    const fh = this.spreadsheet.formulaHelper;
    if (fh && this.isFormulaMode) {
      // Formula mode: transparent textarea text + colored overlay on top
      ta.style.color = 'transparent';
      ta.style.caretColor = '#000';
      ta.style.background = 'transparent';
      ov.style.display = 'block';
      ov.style.background = '#fff';
      ov.innerHTML = fh.getColoredFormulaHTML(ta.value);
    } else {
      // Normal mode: no overlay, normal text color
      const sheet = this.spreadsheet.activeSheet;
      const cell = sheet ? sheet.getCell(this.editRow, this.editCol) : null;
      const style = cell ? cell.getStyle() : null;
      ta.style.color = (style && style.textColor) || '#000';
      ta.style.caretColor = '';
      ta.style.background = (style && style.bgColor) || '#fff';
      ov.style.display = 'none';
    }
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
    e.stopPropagation();

    const ss = this.spreadsheet;
    const sel = ss.selectionManager;
    const fh = ss.formulaHelper;

    // ── Autocomplete always gets first crack ──
    if (fh && this.isFormulaMode) {
      if (fh.handleAutocompleteKey(e, this.textarea)) return;
    }

    // ── F4: cycle cell reference absolute/relative ──
    if (e.key === 'F4' && this.isFormulaMode) {
      e.preventDefault();
      this._cycleAbsoluteRef();
      return;
    }

    // ── F2: toggle between EDIT and POINT modes ──
    if (e.key === 'F2') {
      e.preventDefault();
      if (this.mode === MODE.EDIT && this.isFormulaMode) {
        this.enterPointMode();
      } else if (this.mode === MODE.POINT) {
        this.enterEditMode();
      } else if (this.mode === MODE.ENTER) {
        this.mode = MODE.EDIT; // switch to edit (cursor moves text, not grid)
      }
      return;
    }

    // ── Escape: cancel editing ──
    if (e.key === 'Escape') {
      e.preventDefault();
      this.cancel();
      return;
    }

    // ── Enter ──
    if (e.key === 'Enter' && !e.altKey) {
      e.preventDefault();
      if (e.ctrlKey || e.metaKey) {
        // Ctrl+Enter: commit, stay in same cell
        this.commit();
      } else {
        this.commit();
        sel.move(e.shiftKey ? -1 : 1, 0);
        // Tab-start column: only on Enter (down), not Shift+Enter (up)
        if (!e.shiftKey && this._tabStartCol >= 0) {
          sel.select(sel.activeRow, this._tabStartCol);
          this._tabStartCol = -1;
        }
      }
      ss.renderer.ensureCellVisible(sel.activeRow, sel.activeCol);
      ss.render();
      return;
    }

    // ── Tab: commit and move right/left ──
    if (e.key === 'Tab') {
      e.preventDefault();
      if (this._tabStartCol < 0) this._tabStartCol = this.editCol; // remember starting column
      this.commit();
      sel.tabNext(e.shiftKey);
      ss.renderer.ensureCellVisible(sel.activeRow, sel.activeCol);
      ss.render();
      return;
    }

    // ── Alt+Enter: insert newline ──
    if (e.key === 'Enter' && e.altKey) {
      // Let textarea handle it naturally (inserts newline)
      return;
    }

    // ── Arrow keys: behavior depends on mode ──
    if (e.key === 'ArrowUp' || e.key === 'ArrowDown' || e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
      if (this.mode === MODE.POINT) {
        // POINT mode: arrows insert/extend cell references
        e.preventDefault();
        if (fh) fh.handlePointArrow(e, this.textarea, ss);
        return;
      }

      if (this.mode === MODE.ENTER && !this.isFormulaMode) {
        // ENTER mode (non-formula): arrows commit + navigate
        e.preventDefault();
        this.commit();
        sel.move(
          e.key === 'ArrowDown' ? 1 : e.key === 'ArrowUp' ? -1 : 0,
          e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0,
          e.shiftKey,
        );
        ss.renderer.ensureCellVisible(sel.activeRow, sel.activeCol);
        ss.render();
        return;
      }

      if (this.mode === MODE.ENTER && this.isFormulaMode) {
        // ENTER mode (formula): check if we should enter POINT mode
        const text = this.textarea.value;
        const cursor = this.textarea.selectionStart;
        const before = text.substring(0, cursor).trimEnd();
        const lastChar = before.slice(-1);
        const TRIGGERS = '=+-*/^(,;<>&!';
        if (TRIGGERS.includes(lastChar)) {
          e.preventDefault();
          this.enterPointMode();
          if (fh) fh.handlePointArrow(e, this.textarea, ss);
          return;
        }
        // Otherwise, let textarea handle cursor movement
        return;
      }

      // EDIT mode: let textarea handle arrow keys (cursor movement)
      // No preventDefault — browser handles text cursor
      return;
    }

    // ── Typing an operator in POINT mode: finalize ref, back to ENTER ──
    if (this.mode === MODE.POINT && e.key.length === 1) {
      const OPERATORS = '+-*/^(,;<>&!=)';
      if (OPERATORS.includes(e.key)) {
        this.mode = MODE.ENTER;
        if (fh) {
          fh.pointAnchor = null;
          fh.pointCurrent = null;
        }
        // Let the character be typed normally
        return;
      }
      // Non-operator character: exit point mode, switch to ENTER
      this.mode = MODE.ENTER;
      if (fh) {
        fh.pointMode = false;
        fh.pointAnchor = null;
        fh.pointCurrent = null;
      }
      return;
    }

    // All other keys: let textarea handle normally
  }

  _cycleAbsoluteRef() {
    const ta = this.textarea;
    const text = ta.value;
    const cursor = ta.selectionStart;

    let start = cursor;
    while (start > 0 && /[A-Za-z0-9$:]/.test(text[start - 1])) start--;
    let end = cursor;
    while (end < text.length && /[A-Za-z0-9$:]/.test(text[end])) end++;

    const token = text.substring(start, end);
    if (!token) return;

    const cycled = token.split(':').map(ref => this._cycleOneRef(ref)).join(':');
    if (cycled === token) return;

    ta.value = text.substring(0, start) + cycled + text.substring(end);
    const newEnd = start + cycled.length;
    ta.selectionStart = start;
    ta.selectionEnd = newEnd;
    ta.dispatchEvent(new Event('input'));
  }

  _cycleOneRef(ref) {
    const m = ref.match(/^(\$?)([A-Za-z]+)(\$?)(\d+)$/);
    if (!m) return ref;
    const [, colAbs, col, rowAbs, row] = m;
    if (!colAbs && !rowAbs) return '$' + col + '$' + row;
    if (colAbs && rowAbs) return col + '$' + row;
    if (!colAbs && rowAbs) return '$' + col + row;
    return col + row;
  }

  updatePosition() {
    if (!this.isActive) return;
    const sheet = this.spreadsheet.activeSheet;
    if (!sheet) return;
    const rect = this.spreadsheet.renderer._getCellRect(sheet, this.editRow, this.editCol);
    this.textarea.style.left = rect.x + 'px';
    this.textarea.style.top = rect.y + 'px';
  }

  destroy() {
    if (this.textarea && this.textarea.parentElement) {
      this.textarea.parentElement.removeChild(this.textarea);
    }
  }
}
