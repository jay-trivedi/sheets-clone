import { el } from '../utils/helpers.js';
import CellRange from '../core/CellRange.js';

/**
 * Named Ranges Manager
 * Allows defining aliases for ranges (e.g., "Revenue" for Sheet1!B2:B100).
 * Named ranges can be used in formulas.
 */
export default class NamedRangesManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.ranges = new Map(); // name -> { sheetId, range: CellRange }
  }

  // ── API ──

  define(name, sheetId, range) {
    this.ranges.set(name, { sheetId, range });
  }

  remove(name) {
    this.ranges.delete(name);
  }

  get(name) {
    return this.ranges.get(name) || null;
  }

  getAll() {
    return [...this.ranges.entries()].map(([name, { sheetId, range }]) => {
      const sheet = this.spreadsheet.getSheetById(sheetId);
      return { name, sheetName: sheet ? sheet.name : '?', range: range.toString() };
    });
  }

  // Resolve a name to a Range API object
  resolve(name) {
    const entry = this.ranges.get(name);
    if (!entry) return null;
    const sheet = this.spreadsheet.getSheetById(entry.sheetId);
    if (!sheet) return null;
    return sheet.getRange(entry.range.startRow + 1, entry.range.startCol + 1,
      entry.range.rowCount, entry.range.colCount);
  }

  // ── Dialog ──

  showDialog() {
    const ss = this.spreadsheet;
    const overlay = el('div', { className: 'sheets-modal-overlay' });
    const dialog = el('div', { className: 'sheets-nr-dialog' });

    const refresh = () => {
      const list = dialog.querySelector('.sheets-nr-list');
      list.innerHTML = '';
      for (const { name, sheetName, range } of this.getAll()) {
        const row = el('div', { className: 'sheets-nr-row' });
        row.innerHTML = `<span class="sheets-nr-name">${name}</span>
          <span class="sheets-nr-ref">${sheetName}!${range}</span>
          <button class="sheets-nr-delete" title="Delete">✕</button>`;
        row.querySelector('.sheets-nr-delete').addEventListener('click', () => {
          this.remove(name);
          refresh();
        });
        list.appendChild(row);
      }
      if (this.getAll().length === 0) {
        list.innerHTML = '<div style="color:#888;padding:8px;font-size:12px">No named ranges defined</div>';
      }
    };

    dialog.innerHTML = `
      <h3>Named ranges</h3>
      <div class="sheets-nr-add">
        <input type="text" class="sheets-nr-input" placeholder="Name" />
        <input type="text" class="sheets-nr-range-input" placeholder="Range (e.g. A1:B10)" value="${ss.selection ? ss.selection.toString() : 'A1'}" />
        <button class="sheets-nr-add-btn">Add</button>
      </div>
      <div class="sheets-nr-list"></div>
      <div class="sheets-nr-actions">
        <button class="sheets-nr-done">Done</button>
      </div>
    `;

    dialog.querySelector('.sheets-nr-add-btn').addEventListener('click', () => {
      const name = dialog.querySelector('.sheets-nr-input').value.trim();
      const rangeStr = dialog.querySelector('.sheets-nr-range-input').value.trim();
      if (!name || !rangeStr) return;
      const range = CellRange.fromString(rangeStr);
      if (!range) return;
      this.define(name, ss.activeSheet.id, range);
      dialog.querySelector('.sheets-nr-input').value = '';
      refresh();
    });

    dialog.querySelector('.sheets-nr-done').addEventListener('click', () => overlay.remove());

    overlay.appendChild(dialog);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });
    document.body.appendChild(overlay);
    refresh();
  }

  // ── Serialization ──

  toJSON() {
    const out = {};
    for (const [name, { sheetId, range }] of this.ranges) {
      out[name] = { sheetId, range: range.toString() };
    }
    return out;
  }

  fromJSON(data) {
    this.ranges.clear();
    for (const [name, { sheetId, range }] of Object.entries(data)) {
      const r = CellRange.fromString(range);
      if (r) this.ranges.set(name, { sheetId, range: r });
    }
  }

  destroy() {}
}
