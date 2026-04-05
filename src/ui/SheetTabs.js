import { el } from '../utils/helpers.js';

export default class SheetTabs {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.element = null;
    this.tabsContainer = null;
  }

  init(container) {
    this.element = el('div', { className: 'sheets-sheet-tabs' });

    // Add sheet button
    const addBtn = el('button', {
      className: 'sheets-add-sheet-btn',
      title: 'Add sheet',
      onClick: () => this.spreadsheet.addSheet(),
    }, '+');

    this.tabsContainer = el('div', { className: 'sheets-tabs-container' });

    this.element.appendChild(addBtn);
    this.element.appendChild(this.tabsContainer);
    container.appendChild(this.element);

    this.update();
  }

  update() {
    if (!this.tabsContainer) return;
    const ss = this.spreadsheet;
    this.tabsContainer.innerHTML = '';

    for (const sheet of ss.sheets) {
      const tab = el('div', {
        className: 'sheets-tab' + (sheet === ss.activeSheet ? ' active' : ''),
      });

      const label = el('span', {
        className: 'sheets-tab-label',
      }, sheet.name);

      // Double-click to rename
      label.addEventListener('dblclick', (e) => {
        e.stopPropagation();
        this._startRename(tab, label, sheet);
      });

      tab.appendChild(label);

      // Click to switch
      tab.addEventListener('click', (e) => {
        if (ss.editor && ss.editor.isActive && ss.editor.isFormulaMode) {
          // During formula editing: switch visible sheet but keep editor open
          // Store which sheet we're viewing for cross-sheet ref insertion
          ss._formulaEditSheetId = ss.activeSheet.id;
          ss.activeSheetIndex = ss.sheets.findIndex(s => s.id === sheet.id);
          ss.renderer.scrollTo(0, 0);
          ss.render();
          this.update();
          return;
        }
        ss.setActiveSheet(sheet.id);
        this.update();
      });

      // Right-click for context menu
      tab.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        this._showTabMenu(e.clientX, e.clientY, sheet);
      });

      this.tabsContainer.appendChild(tab);
    }
  }

  _startRename(tab, label, sheet) {
    const input = el('input', {
      type: 'text',
      className: 'sheets-tab-rename',
      value: sheet.name,
    });

    const finish = () => {
      const newName = input.value.trim();
      if (newName && newName !== sheet.name) {
        sheet.name = newName;
      }
      this.update();
    };

    input.addEventListener('blur', finish);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); finish(); }
      if (e.key === 'Escape') { e.preventDefault(); this.update(); }
      e.stopPropagation();
    });

    label.replaceWith(input);
    input.focus();
    input.select();
  }

  _showTabMenu(x, y, sheet) {
    const ss = this.spreadsheet;
    const self = this;

    const menu = el('div', { className: 'sheets-context-menu', style: { left: x + 'px', top: y + 'px' } });

    const closeMenu = () => {
      if (menu.parentElement) menu.remove();
      document.removeEventListener('mousedown', onOutsideClick);
    };

    const onOutsideClick = (e) => {
      if (!menu.contains(e.target)) closeMenu();
    };

    const items = [
      { label: 'Rename', action: () => {
        closeMenu();
        // Use requestAnimationFrame to wait for DOM update after menu closes
        requestAnimationFrame(() => {
          if (!self.tabsContainer) return;
          const tab = self.tabsContainer.querySelector('.active');
          if (tab) {
            const label = tab.querySelector('.sheets-tab-label');
            if (label) self._startRename(tab, label, sheet);
          }
        });
      }},
      { label: 'Duplicate', action: () => { closeMenu(); ss.duplicateSheet(sheet.id); } },
      { label: 'Delete', action: () => { closeMenu(); if (ss.sheets.length > 1) ss.deleteSheet(sheet.id); } },
    ];

    for (const item of items) {
      const menuItem = el('div', {
        className: 'sheets-context-item',
        onClick: () => item.action(),
      }, item.label);
      menu.appendChild(menuItem);
    }

    document.body.appendChild(menu);
    setTimeout(() => document.addEventListener('mousedown', onOutsideClick), 0);
  }

  destroy() {
    if (this.element && this.element.parentElement) {
      this.element.parentElement.removeChild(this.element);
    }
  }
}
