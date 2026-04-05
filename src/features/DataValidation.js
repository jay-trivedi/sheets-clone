import { el } from '../utils/helpers.js';

/**
 * Data Validation Manager
 * Supports: list of items (dropdown), number ranges, text rules, checkboxes.
 * Renders dropdown arrows in cells and shows dropdown list on click.
 */
export default class DataValidationManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this._activeDropdown = null;
  }

  // ── API: set validation on a range ──

  setValidation(sheet, startRow, startCol, endRow, endCol, rule) {
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const cell = sheet.getOrCreateCell(r, c);
        cell.validation = rule;
      }
    }
    this.spreadsheet.render();
  }

  clearValidation(sheet, startRow, startCol, endRow, endCol) {
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const cell = sheet.getCell(r, c);
        if (cell) cell.validation = null;
      }
    }
  }

  // ── Validation rules ──

  static listRule(items, showDropdown = true) {
    return { type: 'list', items: Array.isArray(items) ? items : items.split(',').map(s => s.trim()), showDropdown };
  }

  static numberRule(op, value, value2) {
    return { type: 'number', op, value, value2 }; // op: 'between', 'greaterThan', 'lessThan', etc.
  }

  static textRule(op, value) {
    return { type: 'text', op, value }; // op: 'contains', 'equals', 'isEmail', 'isUrl'
  }

  static checkboxRule(checkedVal = true, uncheckedVal = false) {
    return { type: 'checkbox', checkedVal, uncheckedVal };
  }

  // ── Validate a value against a rule ──

  validate(value, rule) {
    if (!rule) return true;
    switch (rule.type) {
      case 'list':
        return rule.items.includes(String(value));
      case 'number': {
        const n = typeof value === 'number' ? value : parseFloat(value);
        if (isNaN(n)) return false;
        switch (rule.op) {
          case 'between': return n >= rule.value && n <= rule.value2;
          case 'greaterThan': return n > rule.value;
          case 'lessThan': return n < rule.value;
          case 'equalTo': return n === rule.value;
          default: return true;
        }
      }
      case 'text': {
        const s = String(value || '');
        switch (rule.op) {
          case 'contains': return s.includes(rule.value);
          case 'equals': return s === rule.value;
          case 'isEmail': return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
          case 'isUrl': return /^https?:\/\/.+/.test(s);
          default: return true;
        }
      }
      case 'checkbox':
        return value === rule.checkedVal || value === rule.uncheckedVal;
      default:
        return true;
    }
  }

  // ── Render dropdown arrow in cells that have list validation ──

  drawDropdownArrow(ctx, rect, cell) {
    if (!cell || !cell.validation || cell.validation.type !== 'list' || !cell.validation.showDropdown) return false;
    const ax = rect.x + rect.w - 16;
    const ay = rect.y + (rect.h - 8) / 2;
    ctx.fillStyle = '#666';
    ctx.beginPath();
    ctx.moveTo(ax + 2, ay + 2);
    ctx.lineTo(ax + 10, ay + 2);
    ctx.lineTo(ax + 6, ay + 7);
    ctx.closePath();
    ctx.fill();
    return true;
  }

  isDropdownArrow(x, rect, cell) {
    if (!cell || !cell.validation || cell.validation.type !== 'list') return false;
    return x >= rect.x + rect.w - 18;
  }

  // ── Show dropdown list ──

  showDropdown(row, col) {
    this.hideDropdown();
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    if (!sheet) return;
    const cell = sheet.getCell(row, col);
    if (!cell || !cell.validation || cell.validation.type !== 'list') return;

    const rect = ss.renderer._getCellRect(sheet, row, col);
    const containerRect = ss.container.getBoundingClientRect();

    const dropdown = el('div', { className: 'sheets-validation-dropdown' });
    dropdown.style.position = 'fixed';
    dropdown.style.left = (containerRect.left + rect.x) + 'px';
    dropdown.style.top = (containerRect.top + rect.y + rect.h) + 'px';
    dropdown.style.minWidth = rect.w + 'px';

    const currentVal = sheet.getCellValue(row, col);

    for (const item of cell.validation.items) {
      const opt = el('div', { className: 'sheets-validation-option' + (String(item) === String(currentVal) ? ' selected' : '') });
      opt.textContent = item;
      opt.addEventListener('click', () => {
        sheet.setCellValue(row, col, item);
        ss.recalculate();
        ss.render();
        ss.emit('edit', {
          range: sheet.getRange(row + 1, col + 1),
          source: ss, value: item, oldValue: currentVal, sheet,
        });
        this.hideDropdown();
      });
      dropdown.appendChild(opt);
    }

    document.body.appendChild(dropdown);
    this._activeDropdown = dropdown;

    setTimeout(() => {
      const close = (e) => {
        if (!dropdown.contains(e.target)) { this.hideDropdown(); document.removeEventListener('mousedown', close); }
      };
      document.addEventListener('mousedown', close);
    }, 0);
  }

  hideDropdown() {
    if (this._activeDropdown) {
      this._activeDropdown.remove();
      this._activeDropdown = null;
    }
  }

  // ── Dialog to create validation rules ──

  showDialog() {
    const ss = this.spreadsheet;
    const sel = ss.selection;
    if (!sel) return;

    const overlay = el('div', { className: 'sheets-modal-overlay' });
    const dialog = el('div', { className: 'sheets-dv-dialog' });

    dialog.innerHTML = `
      <h3>Data validation</h3>
      <div class="sheets-dv-field">
        <label>Cell range</label>
        <input type="text" class="sheets-dv-range" value="${sel.toString()}" />
      </div>
      <div class="sheets-dv-field">
        <label>Criteria</label>
        <select class="sheets-dv-type">
          <option value="list">List of items</option>
          <option value="number">Number</option>
          <option value="text">Text</option>
          <option value="checkbox">Checkbox</option>
        </select>
      </div>
      <div class="sheets-dv-field sheets-dv-items-field">
        <label>Items (comma-separated)</label>
        <input type="text" class="sheets-dv-items" placeholder="Yes, No, Maybe" />
      </div>
      <div class="sheets-dv-field sheets-dv-number-field" style="display:none">
        <label>Condition</label>
        <select class="sheets-dv-number-op">
          <option value="between">Between</option>
          <option value="greaterThan">Greater than</option>
          <option value="lessThan">Less than</option>
          <option value="equalTo">Equal to</option>
        </select>
        <input type="number" class="sheets-dv-num1" placeholder="Min" />
        <input type="number" class="sheets-dv-num2" placeholder="Max" />
      </div>
      <div class="sheets-dv-actions">
        <button class="sheets-dv-remove">Remove validation</button>
        <span style="flex:1"></span>
        <button class="sheets-dv-cancel">Cancel</button>
        <button class="sheets-dv-save">Save</button>
      </div>
    `;

    const typeSelect = dialog.querySelector('.sheets-dv-type');
    const itemsField = dialog.querySelector('.sheets-dv-items-field');
    const numberField = dialog.querySelector('.sheets-dv-number-field');

    typeSelect.addEventListener('change', () => {
      itemsField.style.display = typeSelect.value === 'list' ? '' : 'none';
      numberField.style.display = typeSelect.value === 'number' ? '' : 'none';
    });

    dialog.querySelector('.sheets-dv-cancel').addEventListener('click', () => overlay.remove());
    dialog.querySelector('.sheets-dv-remove').addEventListener('click', () => {
      this.clearValidation(ss.activeSheet, sel.startRow, sel.startCol, sel.endRow, sel.endCol);
      ss.render();
      overlay.remove();
    });

    dialog.querySelector('.sheets-dv-save').addEventListener('click', () => {
      const type = typeSelect.value;
      let rule;
      if (type === 'list') {
        rule = DataValidationManager.listRule(dialog.querySelector('.sheets-dv-items').value);
      } else if (type === 'number') {
        const op = dialog.querySelector('.sheets-dv-number-op').value;
        const v1 = parseFloat(dialog.querySelector('.sheets-dv-num1').value);
        const v2 = parseFloat(dialog.querySelector('.sheets-dv-num2').value);
        rule = DataValidationManager.numberRule(op, v1, v2);
      } else if (type === 'checkbox') {
        rule = DataValidationManager.checkboxRule();
      } else {
        rule = DataValidationManager.textRule('contains', '');
      }
      this.setValidation(ss.activeSheet, sel.startRow, sel.startCol, sel.endRow, sel.endCol, rule);
      overlay.remove();
    });

    overlay.appendChild(dialog);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });
    document.body.appendChild(overlay);
  }

  destroy() {
    this.hideDropdown();
  }
}
