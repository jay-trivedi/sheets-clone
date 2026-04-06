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
    if (editor && editor.isActive) return;

    // ── READY mode: no editing active ──

    const ctrl = e.ctrlKey || e.metaKey;
    const shift = e.shiftKey;

    switch (e.key) {
      // ── Navigation ──
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
        sel.pageMove(-1, shift);
        break;

      case KEY.PAGE_DOWN:
        e.preventDefault();
        sel.pageMove(1, shift);
        break;

      case KEY.TAB:
        e.preventDefault();
        sel.tabNext(shift);
        break;

      // ── Enter: start editing. Shift+Enter: move up. ──
      case KEY.ENTER:
        e.preventDefault();
        if (shift) {
          sel.move(-1, 0);
        } else if (editor) {
          editor.beginEdit(true);
        }
        break;

      // ── F2: start EDIT mode ──
      case KEY.F2:
        e.preventDefault();
        if (editor) editor.beginEdit(true);
        break;

      // ── Delete: clear cell contents ──
      case KEY.DELETE:
        e.preventDefault();
        ss.deleteSelection();
        break;

      // ── Backspace: clear cell and enter EDIT mode ──
      case KEY.BACKSPACE:
        e.preventDefault();
        ss.deleteSelection();
        if (editor) editor.beginEdit(true);
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
            case 'v':
              if (shift) { e.preventDefault(); ss.pasteValuesOnly(); }
              break; // Regular Ctrl+V: let paste event handle it
            case 'z': e.preventDefault(); if (shift) ss.redo(); else ss.undo(); break;
            case 'y': e.preventDefault(); ss.redo(); break;
            case 'a': e.preventDefault(); sel.selectAll(); break;
            case 'b': e.preventDefault(); ss.toggleBold(); break;
            case 'i': e.preventDefault(); ss.toggleItalic(); break;
            case 'u': e.preventDefault(); ss.toggleUnderline(); break;
            case 'f': e.preventDefault(); ss.showFindReplace(); break;
            case 'h': e.preventDefault(); ss.showFindReplace(true); break;
            case 's': e.preventDefault(); ss.emit('save'); break;
            case 'd': e.preventDefault(); ss.fillDown(); break;
            case 'r': e.preventDefault(); ss.fillRight(); break;
            case ';': e.preventDefault();
              if (shift) { if (ss.editor) ss.editor.beginEnter(new Date().toLocaleTimeString()); }
              else { if (ss.editor) ss.editor.beginEnter(new Date().toLocaleDateString()); }
              break;
            case '\\': e.preventDefault(); ss.clearFormatting(); break;
            case '`': e.preventDefault(); ss.toggleShowFormulas(); break;
          }
          return;
        }

        // ── Typing a printable character: ENTER mode (replaces cell content) ──
        if (e.key.length === 1 && !ctrl && !e.altKey && editor) {
          e.preventDefault();
          editor.beginEnter(e.key);
        }
        break;
    }

    // Ensure cursor (extending end or active cell) is visible after navigation
    ss.renderer.ensureCellVisible(sel._cursorRow, sel._cursorCol);
    ss.render();
  }
}
