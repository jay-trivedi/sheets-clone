import { el } from '../utils/helpers.js';

export default class SheetTabs {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.element = null;
    this.tabsContainer = null;
    this._dragTab = null;
    this._dragIdx = -1;
  }

  init(container) {
    this.element = el('div', { className: 'sheets-sheet-tabs' });

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

    for (let i = 0; i < ss.sheets.length; i++) {
      const sheet = ss.sheets[i];
      const isActive = sheet === ss.activeSheet;

      const tab = el('div', {
        className: 'sheets-tab' + (isActive ? ' active' : ''),
      });

      // Tab color indicator
      if (sheet._tabColor) {
        tab.style.borderBottomColor = sheet._tabColor;
        tab.style.borderBottomWidth = '3px';
      }

      const label = el('span', { className: 'sheets-tab-label' }, sheet.name);

      // Double-click to rename
      label.addEventListener('dblclick', (e) => {
        e.stopPropagation();
        this._startRename(tab, label, sheet);
      });

      tab.appendChild(label);

      // Click to switch
      tab.addEventListener('click', (e) => {
        if (ss.editor && ss.editor.isActive && ss.editor.isFormulaMode) {
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

      // Right-click context menu
      tab.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        this._showTabMenu(e.clientX, e.clientY, sheet);
      });

      // Drag to reorder
      tab.draggable = true;
      tab.addEventListener('dragstart', (e) => {
        this._dragIdx = i;
        this._dragTab = tab;
        tab.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
      });
      tab.addEventListener('dragend', () => {
        tab.classList.remove('dragging');
        this._dragTab = null;
        this._dragIdx = -1;
      });
      tab.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        tab.classList.add('drag-over');
      });
      tab.addEventListener('dragleave', () => {
        tab.classList.remove('drag-over');
      });
      tab.addEventListener('drop', (e) => {
        e.preventDefault();
        tab.classList.remove('drag-over');
        if (this._dragIdx >= 0 && this._dragIdx !== i) {
          const moved = ss.sheets.splice(this._dragIdx, 1)[0];
          ss.sheets.splice(i, 0, moved);
          ss.activeSheetIndex = ss.sheets.findIndex(s => s === ss.activeSheet);
          this.update();
        }
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

    const menu = el('div', { className: 'sheets-context-menu' });

    // Position — ensure it doesn't go off-screen (especially at bottom)
    menu.style.position = 'fixed';
    menu.style.left = x + 'px';
    // Position ABOVE the click point since tabs are at the bottom
    menu.style.top = 'auto';
    menu.style.bottom = (window.innerHeight - y) + 'px';

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
        requestAnimationFrame(() => {
          if (!this.tabsContainer) return;
          const tab = this.tabsContainer.querySelector('.active');
          if (tab) {
            const label = tab.querySelector('.sheets-tab-label');
            if (label) this._startRename(tab, label, sheet);
          }
        });
      }},
      { label: 'Duplicate', action: () => { closeMenu(); ss.duplicateSheet(sheet.id); } },
      { label: 'Delete', action: () => { closeMenu(); if (ss.sheets.length > 1) ss.deleteSheet(sheet.id); } },
      null,
      { label: sheet._hidden ? 'Show' : 'Hide', action: () => {
        closeMenu();
        if (sheet._hidden) { sheet._hidden = false; }
        else if (ss.sheets.filter(s => !s._hidden).length > 1) { sheet._hidden = true; }
        this.update();
      }},
      null,
      { label: 'Change color', action: () => {
        closeMenu();
        const color = prompt('Tab color (hex):', sheet._tabColor || '#4285f4');
        if (color) { sheet._tabColor = color; this.update(); }
      }},
      { label: 'Clear color', action: () => { closeMenu(); sheet._tabColor = null; this.update(); } },
      null,
      { label: 'Move left', action: () => {
        closeMenu();
        const idx = ss.sheets.indexOf(sheet);
        if (idx > 0) { ss.sheets.splice(idx, 1); ss.sheets.splice(idx - 1, 0, sheet); ss.activeSheetIndex = ss.sheets.indexOf(ss.activeSheet); this.update(); }
      }},
      { label: 'Move right', action: () => {
        closeMenu();
        const idx = ss.sheets.indexOf(sheet);
        if (idx < ss.sheets.length - 1) { ss.sheets.splice(idx, 1); ss.sheets.splice(idx + 1, 0, sheet); ss.activeSheetIndex = ss.sheets.indexOf(ss.activeSheet); this.update(); }
      }},
    ];

    for (const item of items) {
      if (item === null) {
        menu.appendChild(el('div', { className: 'sheets-context-sep' }));
        continue;
      }
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
