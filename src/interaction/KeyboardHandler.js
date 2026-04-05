import { KEY } from '../utils/constants.js';

export default class KeyboardHandler {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this._onKeyDown = this._onKeyDown.bind(this);
  }

  init(container) {
    this.container = container;
    container.addEventListener('keydown', this._onKeyDown);
  }

  destroy() {
    this.container.removeEventListener('keydown', this._onKeyDown);
  }

  _onKeyDown(e) {
    const ss = this.spreadsheet;
    const sel = ss.selectionManager;
    const editor = ss.editor;

    // When editor is active, it handles its own keys via stopPropagation.
    // This block is a safety net only — normally we never reach here while editing.
    if (editor && editor.isActive) return;

    const ctrl = e.ctrlKey || e.metaKey;
    const shift = e.shiftKey;

    // ── Navigation keys ──
    switch (e.key) {
      case KEY.ARROW_UP:
        e.preventDefault();
        if (ctrl) sel.moveToEdge(-1, 0, shift);
        else sel.move(-1, 0, shift);
        break;

      case KEY.ARROW_DOWN:
        e.preventDefault();
        if (ctrl) sel.moveToEdge(1, 0, shift);
        else sel.move(1, 0, shift);
        break;

      case KEY.ARROW_LEFT:
        e.preventDefault();
        if (ctrl) sel.moveToEdge(0, -1, shift);
        else sel.move(0, -1, shift);
        break;

      case KEY.ARROW_RIGHT:
        e.preventDefault();
        if (ctrl) sel.moveToEdge(0, 1, shift);
        else sel.move(0, 1, shift);
        break;

      case KEY.HOME:
        e.preventDefault();
        if (ctrl) sel.moveToStart(shift);
        else { if (shift) sel.extendTo(sel.activeRow, 0); else sel.select(sel.activeRow, 0); }
        break;

      case KEY.END:
        e.preventDefault();
        if (ctrl) sel.moveToEnd(shift);
        break;

      case KEY.PAGE_UP:
        e.preventDefault();
        sel.pageMove(-1);
        break;

      case KEY.PAGE_DOWN:
        e.preventDefault();
        sel.pageMove(1);
        break;

      case KEY.TAB:
        e.preventDefault();
        sel.tabNext(shift);
        break;

      // ── Enter: edit current cell (show existing value) ──
      case KEY.ENTER:
        e.preventDefault();
        if (editor) {
          const sheet = ss.activeSheet;
          const cell = sheet ? sheet.getCell(sel.activeRow, sel.activeCol) : null;
          const val = cell ? (cell.formula || cell.displayValue) : '';
          editor.begin(val, true);
        }
        break;

      // ── F2: edit current cell (cursor mode) ──
      case KEY.F2:
        e.preventDefault();
        if (editor) {
          const sheet = ss.activeSheet;
          const cell = sheet ? sheet.getCell(sel.activeRow, sel.activeCol) : null;
          const val = cell ? (cell.formula || cell.displayValue) : '';
          editor.begin(val, true);
        }
        break;

      // ── Delete / Backspace: clear selection ──
      case KEY.DELETE:
      case KEY.BACKSPACE:
        e.preventDefault();
        ss.deleteSelection();
        break;

      // ── Escape: clear copy indicator ──
      case KEY.ESCAPE:
        e.preventDefault();
        ss.copyRange = null;
        ss.render();
        break;

      // ── Space: select row/column/all ──
      case KEY.SPACE:
        if (ctrl && shift) { e.preventDefault(); sel.selectAll(); }
        else if (ctrl) { e.preventDefault(); sel.selectCol(sel.activeCol); }
        else if (shift) { e.preventDefault(); sel.selectRow(sel.activeRow); }
        break;

      default:
        // ── Ctrl shortcuts ──
        if (ctrl) {
          switch (e.key.toLowerCase()) {
            case 'c': e.preventDefault(); ss.copy(); break;
            case 'x': e.preventDefault(); ss.cut(); break;
            case 'v': break; // Let paste event handle it
            case 'z': e.preventDefault(); if (shift) ss.redo(); else ss.undo(); break;
            case 'y': e.preventDefault(); ss.redo(); break;
            case 'a': e.preventDefault(); sel.selectAll(); break;
            case 'b': e.preventDefault(); ss.toggleBold(); break;
            case 'i': e.preventDefault(); ss.toggleItalic(); break;
            case 'u': e.preventDefault(); ss.toggleUnderline(); break;
            case 'f': e.preventDefault(); ss.showFindReplace(); break;
            case 'h': e.preventDefault(); ss.showFindReplace(true); break;
            case 's': e.preventDefault(); ss.emit('save'); break;
          }
          return;
        }

        // ── Typing a printable character starts editing ──
        if (e.key.length === 1 && !ctrl && !e.altKey && editor) {
          editor.begin(e.key);
        }
        break;
    }

    // Ensure active cell is visible after any navigation
    ss.renderer.ensureCellVisible(sel.activeRow, sel.activeCol);
    ss.render();
  }
}
