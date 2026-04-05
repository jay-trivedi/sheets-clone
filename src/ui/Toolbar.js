import { el } from '../utils/helpers.js';
import { H_ALIGN, NUMBER_FORMATS } from '../utils/constants.js';

export default class Toolbar {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.element = null;
  }

  init(container) {
    this.element = el('div', { className: 'sheets-toolbar' });

    const groups = [
      this._buildUndoGroup(),
      this._buildFormatGroup(),
      this._buildFontGroup(),
      this._buildAlignGroup(),
      this._buildBorderGroup(),
      this._buildColorGroup(),
      this._buildMergeGroup(),
      this._buildNumberGroup(),
      this._buildInsertGroup(),
    ];

    for (let i = 0; i < groups.length; i++) {
      this.element.appendChild(groups[i]);
      if (i < groups.length - 1) {
        this.element.appendChild(el('div', { className: 'sheets-toolbar-sep' }));
      }
    }

    container.appendChild(this.element);
  }

  _btn(title, html, onClick) {
    const btn = el('button', {
      className: 'sheets-toolbar-btn',
      title,
      onClick,
    });
    btn.innerHTML = html;
    return btn;
  }

  _select(title, options, onChange) {
    const select = el('select', { className: 'sheets-toolbar-select', title });
    for (const [label, value] of options) {
      const opt = el('option', { value }, label);
      select.appendChild(opt);
    }
    select.addEventListener('change', () => onChange(select.value));
    return select;
  }

  _buildUndoGroup() {
    const ss = this.spreadsheet;
    const group = el('div', { className: 'sheets-toolbar-group' });
    group.appendChild(this._btn('Undo (Ctrl+Z)', `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 10h10a5 5 0 0 1 0 10H12"/><path d="M3 10l4-4"/><path d="M3 10l4 4"/></svg>`, () => ss.undo()));
    group.appendChild(this._btn('Redo (Ctrl+Y)', `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 10H11a5 5 0 0 0 0 10h1"/><path d="M21 10l-4-4"/><path d="M21 10l-4 4"/></svg>`, () => ss.redo()));
    return group;
  }

  _buildFormatGroup() {
    const ss = this.spreadsheet;
    const group = el('div', { className: 'sheets-toolbar-group' });
    group.appendChild(this._btn('Bold (Ctrl+B)', '<b>B</b>', () => ss.toggleBold()));
    group.appendChild(this._btn('Italic (Ctrl+I)', '<i>I</i>', () => ss.toggleItalic()));
    group.appendChild(this._btn('Underline (Ctrl+U)', '<u>U</u>', () => ss.toggleUnderline()));
    group.appendChild(this._btn('Strikethrough', '<s>S</s>', () => ss.toggleStrikethrough()));
    return group;
  }

  _buildFontGroup() {
    const ss = this.spreadsheet;
    const group = el('div', { className: 'sheets-toolbar-group' });

    const fontFamilies = [
      ['Arial', 'Arial'], ['Helvetica', 'Helvetica'], ['Times New Roman', 'Times New Roman'],
      ['Courier New', 'Courier New'], ['Georgia', 'Georgia'], ['Verdana', 'Verdana'],
      ['Trebuchet MS', 'Trebuchet MS'], ['Comic Sans MS', 'Comic Sans MS'],
    ];
    group.appendChild(this._select('Font family', fontFamilies, (v) => ss.setStyle({ fontFamily: v })));

    const fontSizes = [
      ['6', '6'], ['7', '7'], ['8', '8'], ['9', '9'], ['10', '10'], ['11', '11'], ['12', '12'],
      ['14', '14'], ['16', '16'], ['18', '18'], ['20', '20'], ['24', '24'], ['28', '28'],
      ['36', '36'], ['48', '48'], ['72', '72'],
    ];
    const sizeSelect = this._select('Font size', fontSizes, (v) => ss.setStyle({ fontSize: parseInt(v) }));
    sizeSelect.value = '10';
    group.appendChild(sizeSelect);

    return group;
  }

  _buildAlignGroup() {
    const ss = this.spreadsheet;
    const group = el('div', { className: 'sheets-toolbar-group' });
    group.appendChild(this._btn('Align left', `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="15" y2="12"/><line x1="3" y1="18" x2="18" y2="18"/></svg>`, () => ss.setStyle({ hAlign: H_ALIGN.LEFT })));
    group.appendChild(this._btn('Align center', `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="3" y1="6" x2="21" y2="6"/><line x1="6" y1="12" x2="18" y2="12"/><line x1="4" y1="18" x2="20" y2="18"/></svg>`, () => ss.setStyle({ hAlign: H_ALIGN.CENTER })));
    group.appendChild(this._btn('Align right', `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="3" y1="6" x2="21" y2="6"/><line x1="9" y1="12" x2="21" y2="12"/><line x1="6" y1="18" x2="21" y2="18"/></svg>`, () => ss.setStyle({ hAlign: H_ALIGN.RIGHT })));
    group.appendChild(this._btn('Wrap text', `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18"/><path d="M3 12h15a3 3 0 1 1 0 6h-4"/><path d="M16 16l-2 2 2 2"/><path d="M3 18h7"/></svg>`, () => ss.toggleWrap()));
    return group;
  }

  _buildBorderGroup() {
    const ss = this.spreadsheet;
    const group = el('div', { className: 'sheets-toolbar-group' });
    group.appendChild(this._btn('Borders', `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="1"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="12" y1="3" x2="12" y2="21"/></svg>`,
      () => ss.setAllBorders({ color: '#000', width: 1 })));
    group.appendChild(this._btn('Clear borders', `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="1" stroke-dasharray="3 3"/></svg>`,
      () => ss.clearBorders()));
    return group;
  }

  _buildColorGroup() {
    const ss = this.spreadsheet;
    const group = el('div', { className: 'sheets-toolbar-group' });

    // Google Sheets color palette
    const PALETTE = [
      '#000000','#434343','#666666','#999999','#b7b7b7','#cccccc','#d9d9d9','#efefef','#f3f3f3','#ffffff',
      '#980000','#ff0000','#ff9900','#ffff00','#00ff00','#00ffff','#4a86e8','#0000ff','#9900ff','#ff00ff',
      '#e6b8af','#f4cccc','#fce5cd','#fff2cc','#d9ead3','#d0e0e3','#c9daf8','#cfe2f3','#d9d2e9','#ead1dc',
      '#dd7e6b','#ea9999','#f9cb9c','#ffe599','#b6d7a8','#a2c4c9','#a4c2f4','#9fc5e8','#b4a7d6','#d5a6bd',
      '#cc4125','#e06666','#f6b26b','#ffd966','#93c47d','#76a5af','#6d9eeb','#6fa8dc','#8e7cc3','#c27ba0',
      '#a61c00','#cc0000','#e69138','#f1c232','#6aa84f','#45818e','#3c78d8','#3d85c6','#674ea7','#a64d79',
      '#85200c','#990000','#b45f06','#bf9000','#38761d','#134f5c','#1155cc','#0b5394','#351c75','#741b47',
      '#5b0f00','#660000','#783f04','#7f6000','#274e13','#0c343d','#1c4587','#073763','#20124d','#4c1130',
    ];

    function buildColorPicker(title, icon, styleProp, defaultColor) {
      const wrapper = el('div', { className: 'sheets-color-picker-wrap' });

      // Main button
      const btn = el('div', { className: 'sheets-color-btn', title });
      btn.innerHTML = icon;
      const colorLine = el('div', { className: 'sheets-color-indicator' });
      colorLine.style.background = defaultColor;
      btn.appendChild(colorLine);
      wrapper.appendChild(btn);

      // Dropdown arrow
      const arrow = el('div', { className: 'sheets-color-arrow', title });
      arrow.innerHTML = '&#9662;';
      wrapper.appendChild(arrow);

      // Dropdown panel
      const dropdown = el('div', { className: 'sheets-color-dropdown' });
      dropdown.style.display = 'none';

      // Reset button
      const resetBtn = el('div', { className: 'sheets-color-reset' });
      resetBtn.textContent = styleProp === 'bgColor' ? 'None' : 'Automatic';
      resetBtn.addEventListener('click', () => {
        ss.setStyle({ [styleProp]: styleProp === 'bgColor' ? null : '#000000' });
        colorLine.style.background = defaultColor;
        dropdown.style.display = 'none';
      });
      dropdown.appendChild(resetBtn);

      // Color grid
      const grid = el('div', { className: 'sheets-color-grid' });
      for (const color of PALETTE) {
        const swatch = el('div', { className: 'sheets-color-swatch' });
        swatch.style.background = color;
        if (color === '#ffffff') swatch.style.border = '1px solid #ddd';
        swatch.title = color;
        swatch.addEventListener('click', () => {
          ss.setStyle({ [styleProp]: color });
          colorLine.style.background = color;
          dropdown.style.display = 'none';
        });
        grid.appendChild(swatch);
      }
      dropdown.appendChild(grid);

      // Custom color button
      const customBtn = el('div', { className: 'sheets-color-custom' });
      customBtn.textContent = 'Custom...';
      const customInput = el('input', { type: 'color', value: defaultColor,
        style: { position: 'absolute', opacity: '0', width: '0', height: '0', pointerEvents: 'none' } });
      customInput.addEventListener('input', () => {
        ss.setStyle({ [styleProp]: customInput.value });
        colorLine.style.background = customInput.value;
        dropdown.style.display = 'none';
      });
      customBtn.appendChild(customInput);
      customBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        customInput.click();
      });
      dropdown.appendChild(customBtn);

      // Append dropdown to body so toolbar overflow:hidden doesn't clip it
      dropdown.style.position = 'fixed';
      document.body.appendChild(dropdown);

      function openDropdown() {
        document.querySelectorAll('.sheets-color-dropdown').forEach(d => d.style.display = 'none');
        const rect = arrow.getBoundingClientRect();
        dropdown.style.left = rect.left + 'px';
        dropdown.style.top = rect.bottom + 2 + 'px';
        dropdown.style.display = 'block';
        // Adjust if off-screen
        requestAnimationFrame(() => {
          const dr = dropdown.getBoundingClientRect();
          if (dr.right > window.innerWidth) dropdown.style.left = (window.innerWidth - dr.width - 8) + 'px';
          if (dr.bottom > window.innerHeight) dropdown.style.top = (rect.top - dr.height - 2) + 'px';
        });
        const close = (ev) => {
          if (!dropdown.contains(ev.target) && !arrow.contains(ev.target)) {
            dropdown.style.display = 'none';
            document.removeEventListener('mousedown', close);
          }
        };
        setTimeout(() => document.addEventListener('mousedown', close), 0);
      }

      // Click main button: apply last-used color
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        ss.setStyle({ [styleProp]: colorLine.style.background || defaultColor });
      });

      // Click arrow: toggle dropdown
      arrow.addEventListener('click', (e) => {
        e.stopPropagation();
        if (dropdown.style.display !== 'none') {
          dropdown.style.display = 'none';
        } else {
          openDropdown();
        }
      });

      return wrapper;
    }

    group.appendChild(buildColorPicker(
      'Text color', '<span style="font-weight:bold;font-size:14px;">A</span>',
      'textColor', '#000000'
    ));
    group.appendChild(buildColorPicker(
      'Fill color', `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2L2 22h20L12 2z"/></svg>`,
      'bgColor', '#ffffff'
    ));

    return group;
  }

  _buildMergeGroup() {
    const ss = this.spreadsheet;
    const group = el('div', { className: 'sheets-toolbar-group' });
    group.appendChild(this._btn('Merge cells', `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="1"/><path d="M9 12h6m-3-3v6"/></svg>`, () => ss.mergeSelection()));
    group.appendChild(this._btn('Unmerge cells', `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="1"/><path d="M9 12h6"/></svg>`, () => ss.unmergeSelection()));
    return group;
  }

  _buildNumberGroup() {
    const ss = this.spreadsheet;
    const group = el('div', { className: 'sheets-toolbar-group' });

    const formats = [
      ['Auto', 'General'],
      ['Number', '#,##0.00'],
      ['Integer', '#,##0'],
      ['Currency', '$#,##0.00'],
      ['Percent', '0.00%'],
      ['Date', 'M/d/yyyy'],
      ['Time', 'h:mm:ss AM/PM'],
      ['Scientific', '0.00E+0'],
      ['Text', '@'],
    ];

    group.appendChild(this._select('Number format', formats, (v) => ss.setStyle({ numberFormat: v })));
    group.appendChild(this._btn('Decrease decimals', '.0', () => ss.adjustDecimals(-1)));
    group.appendChild(this._btn('Increase decimals', '.00', () => ss.adjustDecimals(1)));
    return group;
  }

  _buildInsertGroup() {
    const ss = this.spreadsheet;
    const group = el('div', { className: 'sheets-toolbar-group' });
    group.appendChild(this._btn('Insert row above', `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>R`, () => ss.insertRowAbove()));
    group.appendChild(this._btn('Insert column left', `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>C`, () => ss.insertColLeft()));
    group.appendChild(this._btn('Delete row', `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="5" y1="12" x2="19" y2="12"/></svg>R`, () => ss.deleteSelectedRows()));
    group.appendChild(this._btn('Delete column', `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="5" y1="12" x2="19" y2="12"/></svg>C`, () => ss.deleteSelectedCols()));
    return group;
  }

  destroy() {
    if (this.element && this.element.parentElement) {
      this.element.parentElement.removeChild(this.element);
    }
  }
}
