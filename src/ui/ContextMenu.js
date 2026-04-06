import { el } from '../utils/helpers.js';

export default class ContextMenu {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.menu = null;
    this._closeHandler = null;
  }

  show(x, y, cell, header) {
    this.hide();

    const ss = this.spreadsheet;
    const sel = ss.selection;

    const items = [];

    if (cell) {
      items.push({ label: 'Cut', shortcut: 'Ctrl+X', action: () => ss.cut() });
      items.push({ label: 'Copy', shortcut: 'Ctrl+C', action: () => ss.copy() });
      items.push({ label: 'Paste', shortcut: 'Ctrl+V', action: () => ss.paste() });
      items.push({ label: 'Paste special', children: [
        { label: 'Paste values only', shortcut: 'Ctrl+Shift+V', action: () => ss.pasteValuesOnly() },
        { label: 'Paste format only', action: () => ss.pasteFormatOnly() },
        { label: 'Paste transpose', action: () => ss.pasteTranspose() },
      ]});
      items.push(null);

      items.push({ label: 'Insert row above', action: () => ss.insertRowAbove() });
      items.push({ label: 'Insert row below', action: () => ss.insertRowBelow() });
      items.push({ label: 'Insert column left', action: () => ss.insertColLeft() });
      items.push({ label: 'Insert column right', action: () => ss.insertColRight() });
      items.push(null);

      items.push({ label: 'Delete row', action: () => ss.deleteSelectedRows() });
      items.push({ label: 'Delete column', action: () => ss.deleteSelectedCols() });
      items.push(null);

      items.push({ label: 'Clear contents', action: () => ss.deleteSelection() });
      items.push({ label: 'Clear formatting', action: () => ss.clearFormatting() });
      items.push(null);

      if (sel && !sel.isSingleCell) {
        items.push({ label: 'Merge cells', action: () => ss.mergeSelection() });
        items.push({ label: 'Unmerge cells', action: () => ss.unmergeSelection() });
        items.push(null);
      }

      items.push({ label: 'Hide rows', action: () => ss.hideSelectedRows() });
      items.push({ label: 'Hide columns', action: () => ss.hideSelectedCols() });
      items.push(null);

      items.push({ label: 'Sort A → Z', action: () => ss.sortAscending() });
      items.push({ label: 'Sort Z → A', action: () => ss.sortDescending() });
      items.push(null);

      items.push({ label: 'Data validation...', action: () => ss.showDataValidation() });
      items.push({ label: 'Conditional formatting...', action: () => ss.showConditionalFormatting() });
      items.push(null);

      items.push({ label: 'Freeze rows', action: () => ss.freezeRows(cell.row + 1) });
      items.push({ label: 'Freeze columns', action: () => ss.freezeCols(cell.col + 1) });
      items.push({ label: 'Unfreeze', action: () => { ss.freezeRows(0); ss.freezeCols(0); } });
    }

    if (header) {
      if (header.type === 'col') {
        items.push({ label: 'Insert column left', action: () => { ss.activeSheet.insertCols(header.index); ss.recalculate(); ss.render(); } });
        items.push({ label: 'Insert column right', action: () => { ss.activeSheet.insertCols(header.index + 1); ss.recalculate(); ss.render(); } });
        items.push({ label: 'Delete column', action: () => { ss.activeSheet.deleteCols(header.index); ss.recalculate(); ss.render(); } });
        items.push({ label: 'Clear column', action: () => {
          const sheet = ss.activeSheet;
          for (let r = 0; r < sheet.rowCount; r++) { const c = sheet.getCell(r, header.index); if (c) c.clearContent(); }
          ss.render();
        }});
        items.push(null);
        items.push({ label: 'Hide column', action: () => { ss.activeSheet.hiddenCols.add(header.index); ss.render(); } });
        items.push({ label: 'Unhide columns', action: () => { ss.unhideCols(Math.max(0, header.index - 1), header.index + 1); } });
        items.push(null);
        items.push({ label: 'Resize column...', action: () => {
          const w = prompt('Column width:', ss.activeSheet.getColWidth(header.index));
          if (w) ss.activeSheet.setColWidth(header.index, parseInt(w));
          ss.render();
        }});
        items.push(null);
        items.push({ label: 'Sort sheet A → Z', action: () => ss.sortAscending() });
        items.push({ label: 'Sort sheet Z → A', action: () => ss.sortDescending() });
      }
      if (header.type === 'row') {
        items.push({ label: 'Insert row above', action: () => { ss.activeSheet.insertRows(header.index); ss.recalculate(); ss.render(); } });
        items.push({ label: 'Insert row below', action: () => { ss.activeSheet.insertRows(header.index + 1); ss.recalculate(); ss.render(); } });
        items.push({ label: 'Delete row', action: () => { ss.activeSheet.deleteRows(header.index); ss.recalculate(); ss.render(); } });
        items.push({ label: 'Clear row', action: () => {
          const sheet = ss.activeSheet;
          for (let c = 0; c < sheet.colCount; c++) { const cell = sheet.getCell(header.index, c); if (cell) cell.clearContent(); }
          ss.render();
        }});
        items.push(null);
        items.push({ label: 'Hide row', action: () => { ss.activeSheet.hiddenRows.add(header.index); ss.render(); } });
        items.push({ label: 'Unhide rows', action: () => { ss.unhideRows(Math.max(0, header.index - 1), header.index + 1); } });
        items.push(null);
        items.push({ label: 'Resize row...', action: () => {
          const h = prompt('Row height:', ss.activeSheet.getRowHeight(header.index));
          if (h) ss.activeSheet.setRowHeight(header.index, parseInt(h));
          ss.render();
        }});
      }
    }

    if (items.length === 0) return;

    this.menu = el('div', { className: 'sheets-context-menu' });
    this.menu.style.left = x + 'px';
    this.menu.style.top = y + 'px';

    for (const item of items) {
      if (item === null) {
        this.menu.appendChild(el('div', { className: 'sheets-context-sep' }));
        continue;
      }

      const row = el('div', { className: 'sheets-context-item' + (item.children ? ' sheets-has-submenu' : '') });
      row.appendChild(el('span', {}, item.label));
      if (item.shortcut) {
        row.appendChild(el('span', { className: 'sheets-context-shortcut' }, item.shortcut));
      }
      if (item.children) {
        row.appendChild(el('span', { className: 'sheets-context-shortcut' }, '▸'));
        const sub = el('div', { className: 'sheets-context-submenu' });
        for (const child of item.children) {
          const subRow = el('div', { className: 'sheets-context-item' });
          subRow.appendChild(el('span', {}, child.label));
          if (child.shortcut) subRow.appendChild(el('span', { className: 'sheets-context-shortcut' }, child.shortcut));
          subRow.addEventListener('click', (e) => { e.stopPropagation(); this.hide(); child.action(); });
          sub.appendChild(subRow);
        }
        row.appendChild(sub);
      } else {
        row.addEventListener('click', () => { this.hide(); item.action(); });
      }
      this.menu.appendChild(row);
    }

    document.body.appendChild(this.menu);

    // Adjust position to stay within viewport
    requestAnimationFrame(() => {
      if (!this.menu) return;
      const rect = this.menu.getBoundingClientRect();
      if (rect.right > window.innerWidth) {
        this.menu.style.left = (x - rect.width) + 'px';
      }
      if (rect.bottom > window.innerHeight) {
        this.menu.style.top = (y - rect.height) + 'px';
      }
    });

    this._closeHandler = (e) => {
      if (this.menu && !this.menu.contains(e.target)) {
        this.hide();
      }
    };
    setTimeout(() => document.addEventListener('mousedown', this._closeHandler), 0);
  }

  hide() {
    if (this.menu) {
      this.menu.remove();
      this.menu = null;
    }
    if (this._closeHandler) {
      document.removeEventListener('mousedown', this._closeHandler);
      this._closeHandler = null;
    }
  }

  destroy() {
    this.hide();
  }
}
