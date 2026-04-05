import { el } from '../utils/helpers.js';

/**
 * Conditional Formatting UI
 * Dialog to create, edit, and manage conditional format rules.
 */
export default class ConditionalFormatUI {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
  }

  showDialog() {
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    if (!sheet) return;
    const sel = ss.selection;

    const overlay = el('div', { className: 'sheets-modal-overlay' });
    const dialog = el('div', { className: 'sheets-cf-dialog' });

    dialog.innerHTML = `
      <h3>Conditional format rules</h3>
      <div class="sheets-cf-rules"></div>
      <button class="sheets-cf-add">+ Add another rule</button>
      <div class="sheets-cf-new" style="display:none">
        <div class="sheets-cf-field">
          <label>Apply to range</label>
          <input type="text" class="sheets-cf-range" value="${sel ? sel.toString() : 'A1:Z100'}" />
        </div>
        <div class="sheets-cf-field">
          <label>Format cells if...</label>
          <select class="sheets-cf-condition">
            <option value="greaterThan">Greater than</option>
            <option value="lessThan">Less than</option>
            <option value="equalTo">Equal to</option>
            <option value="between">Between</option>
            <option value="text_contains">Text contains</option>
            <option value="isEmpty">Is empty</option>
            <option value="isNotEmpty">Is not empty</option>
          </select>
        </div>
        <div class="sheets-cf-field sheets-cf-value-field">
          <label>Value</label>
          <input type="text" class="sheets-cf-value" placeholder="Value" />
          <input type="text" class="sheets-cf-value2" placeholder="Max" style="display:none" />
        </div>
        <div class="sheets-cf-field">
          <label>Formatting</label>
          <div class="sheets-cf-format-preview">
            <input type="color" class="sheets-cf-bg" value="#c6efce" title="Background" />
            <input type="color" class="sheets-cf-fg" value="#006100" title="Text color" />
            <label><input type="checkbox" class="sheets-cf-bold" /> Bold</label>
          </div>
        </div>
        <div class="sheets-cf-new-actions">
          <button class="sheets-cf-cancel-new">Cancel</button>
          <button class="sheets-cf-save-new">Done</button>
        </div>
      </div>
      <div class="sheets-cf-actions">
        <button class="sheets-cf-close">Close</button>
      </div>
    `;

    const rulesDiv = dialog.querySelector('.sheets-cf-rules');
    const newSection = dialog.querySelector('.sheets-cf-new');
    const condSelect = dialog.querySelector('.sheets-cf-condition');
    const val2 = dialog.querySelector('.sheets-cf-value2');

    condSelect.addEventListener('change', () => {
      val2.style.display = condSelect.value === 'between' ? '' : 'none';
      dialog.querySelector('.sheets-cf-value-field').style.display =
        ['isEmpty', 'isNotEmpty'].includes(condSelect.value) ? 'none' : '';
    });

    const refreshRules = () => {
      rulesDiv.innerHTML = '';
      for (let i = 0; i < sheet.conditionalFormats.length; i++) {
        const cf = sheet.conditionalFormats[i];
        const row = el('div', { className: 'sheets-cf-rule-row' });
        const swatch = `<span style="display:inline-block;width:16px;height:16px;background:${cf.style?.bgColor || '#eee'};border:1px solid #ccc;border-radius:2px;vertical-align:middle;margin-right:6px"></span>`;
        row.innerHTML = `${swatch} <span>${cf.type} ${cf.value || cf.min || ''}</span> on ${cf.range}
          <button class="sheets-cf-rule-delete" title="Delete">✕</button>`;
        row.querySelector('.sheets-cf-rule-delete').addEventListener('click', () => {
          sheet.conditionalFormats.splice(i, 1);
          ss.render();
          refreshRules();
        });
        rulesDiv.appendChild(row);
      }
      if (sheet.conditionalFormats.length === 0) {
        rulesDiv.innerHTML = '<div style="color:#888;padding:8px;font-size:12px">No rules yet</div>';
      }
    };

    dialog.querySelector('.sheets-cf-add').addEventListener('click', () => {
      newSection.style.display = '';
    });

    dialog.querySelector('.sheets-cf-cancel-new').addEventListener('click', () => {
      newSection.style.display = 'none';
    });

    dialog.querySelector('.sheets-cf-save-new').addEventListener('click', () => {
      const rangeStr = dialog.querySelector('.sheets-cf-range').value;
      const type = condSelect.value;
      const value = dialog.querySelector('.sheets-cf-value').value;
      const value2 = dialog.querySelector('.sheets-cf-value2').value;
      const bgColor = dialog.querySelector('.sheets-cf-bg').value;
      const textColor = dialog.querySelector('.sheets-cf-fg').value;
      const bold = dialog.querySelector('.sheets-cf-bold').checked;

      const rule = { range: rangeStr, type, style: { bgColor, textColor } };
      if (bold) rule.style.bold = true;

      if (type === 'between') {
        rule.min = parseFloat(value);
        rule.max = parseFloat(value2);
      } else if (['greaterThan', 'lessThan', 'equalTo'].includes(type)) {
        rule.value = parseFloat(value) || value;
      } else if (type === 'text_contains') {
        rule.value = value;
      }

      sheet.conditionalFormats.push(rule);
      ss.render();
      newSection.style.display = 'none';
      refreshRules();
    });

    dialog.querySelector('.sheets-cf-close').addEventListener('click', () => overlay.remove());

    overlay.appendChild(dialog);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });
    document.body.appendChild(overlay);
    refreshRules();
  }

  destroy() {}
}
