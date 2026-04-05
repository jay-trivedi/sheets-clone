import { el } from '../utils/helpers.js';
import { H_ALIGN, V_ALIGN, NUMBER_FORMATS } from '../utils/constants.js';

// Google Sheets toolbar — exact order, left to right
export default class Toolbar {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.element = null;
    this._lastTextColor = '#000000';
    this._lastFillColor = '#ffff00';
  }

  init(container) {
    this.element = el('div', { className: 'sheets-toolbar' });
    const ss = this.spreadsheet;

    // 1-2: Undo / Redo
    this._addBtn('Undo (Ctrl+Z)', ICONS.undo, () => ss.undo());
    this._addBtn('Redo (Ctrl+Y)', ICONS.redo, () => ss.redo());
    this._addBtn('Print (Ctrl+P)', ICONS.print, () => window.print());
    this._sep();

    // 3b: Format painter
    this._addBtn('Paint format', ICONS.formatPainter, () => ss.startFormatPainter());
    this._sep();

    // 4-8: Number formatting
    this._addBtn('Format as currency', '<span style="font-size:13px;font-weight:500">$</span>', () => ss.setStyle({ numberFormat: '$#,##0.00' }));
    this._addBtn('Format as percent', '<span style="font-size:13px;font-weight:500">%</span>', () => ss.setStyle({ numberFormat: '0.00%' }));
    this._addBtn('Decrease decimal places', ICONS.decimalDec, () => ss.adjustDecimals(-1));
    this._addBtn('Increase decimal places', ICONS.decimalInc, () => ss.adjustDecimals(1));
    this._addDropdown('More formats', '123▾', this._buildFormatsMenu());
    this._sep();

    // 10: Font family
    const fontFamilies = [
      'Arial','Helvetica','Times New Roman','Courier New','Georgia',
      'Verdana','Trebuchet MS','Comic Sans MS','Impact','Lucida Console',
    ];
    this._addSelect('Font', fontFamilies.map(f => [f, f]), 'Arial', (v) => ss.setStyle({ fontFamily: v }), 'sheets-font-select');
    this._sep();

    // 12: Font size
    const sizes = ['6','7','8','9','10','11','12','14','16','18','20','24','28','36','48','72'];
    this._addSelect('Font size', sizes.map(s => [s, s]), '10', (v) => ss.setStyle({ fontSize: parseInt(v) }), 'sheets-size-select');
    this._sep();

    // 14-18: Bold, Italic, Underline, Strikethrough, Text color
    this._addBtn('Bold (Ctrl+B)', '<b>B</b>', () => ss.toggleBold(), 'sheets-btn-bold');
    this._addBtn('Italic (Ctrl+I)', '<i>I</i>', () => ss.toggleItalic(), 'sheets-btn-italic');
    this._addBtn('Underline (Ctrl+U)', '<u>U</u>', () => ss.toggleUnderline());
    this._addBtn('Strikethrough', '<s>S</s>', () => ss.toggleStrikethrough());
    this._addColorPicker('Text color', '<span style="font-weight:bold;font-size:14px">A</span>', 'textColor', '#000000');
    this._sep();

    // 20: Fill color
    this._addColorPicker('Fill color', ICONS.fillBucket, 'bgColor', '#ffffff');
    this._sep();

    // 22-23: Borders, Merge
    this._addDropdown('Borders', ICONS.borders, this._buildBordersMenu());
    this._addDropdown('Merge cells', ICONS.merge, this._buildMergeMenu());
    this._sep();

    // 25-28: Align H, Align V, Wrapping, Rotation
    this._addDropdown('Horizontal align', ICONS.alignLeft, this._buildHAlignMenu());
    this._addDropdown('Vertical align', ICONS.valignMiddle, this._buildVAlignMenu());
    this._addDropdown('Text wrapping', ICONS.wrapClip, this._buildWrapMenu());
    this._addDropdown('Text rotation', ICONS.textRotation, this._buildRotationMenu());
    this._sep();

    // 30-32: Link, Comment, Chart
    this._addBtn('Insert link', ICONS.link, () => this._insertLink());
    this._addBtn('Insert comment', ICONS.comment, () => this._insertComment());
    this._addBtn('Insert chart', ICONS.chart, () => this._toast('Charts coming soon'));
    this._sep();

    // 34: Filter
    this._addBtn('Create a filter', ICONS.filter, () => ss.toggleFilter());
    this._sep();

    // 36: Functions
    this._addDropdown('Functions', ICONS.functions, this._buildFunctionsMenu());

    container.appendChild(this.element);
  }

  // ── Button helpers ──

  _addBtn(title, html, onClick, className) {
    const btn = el('button', { className: 'sheets-toolbar-btn' + (className ? ' ' + className : ''), title, onClick });
    btn.innerHTML = html;
    this.element.appendChild(btn);
    return btn;
  }

  _sep() {
    this.element.appendChild(el('div', { className: 'sheets-toolbar-sep' }));
  }

  _addSelect(title, options, defaultVal, onChange, className) {
    const select = el('select', { className: 'sheets-toolbar-select' + (className ? ' ' + className : ''), title });
    for (const [label, value] of options) {
      select.appendChild(el('option', { value }, label));
    }
    select.value = defaultVal;
    select.addEventListener('change', () => onChange(select.value));
    this.element.appendChild(select);
    return select;
  }

  _addDropdown(title, iconHtml, menuItems) {
    const wrap = el('div', { className: 'sheets-toolbar-dropdown-wrap' });
    const btn = el('button', { className: 'sheets-toolbar-btn sheets-has-dropdown', title });
    btn.innerHTML = iconHtml + '<span class="sheets-dropdown-caret">&#9662;</span>';
    wrap.appendChild(btn);

    const menu = el('div', { className: 'sheets-toolbar-menu' });
    for (const item of menuItems) {
      if (item === null) {
        menu.appendChild(el('div', { className: 'sheets-toolbar-menu-sep' }));
        continue;
      }
      const row = el('div', { className: 'sheets-toolbar-menu-item' });
      if (item.icon) {
        const iconSpan = el('span', { className: 'sheets-menu-icon' });
        iconSpan.innerHTML = item.icon;
        row.appendChild(iconSpan);
      }
      row.appendChild(el('span', {}, item.label));
      row.addEventListener('click', () => { item.action(); menu.style.display = 'none'; });
      menu.appendChild(row);
    }
    document.body.appendChild(menu);
    menu.style.display = 'none';
    menu.style.position = 'fixed';

    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      document.querySelectorAll('.sheets-toolbar-menu').forEach(m => m.style.display = 'none');
      const rect = btn.getBoundingClientRect();
      menu.style.left = rect.left + 'px';
      menu.style.top = rect.bottom + 2 + 'px';
      menu.style.display = 'block';
      requestAnimationFrame(() => {
        const mr = menu.getBoundingClientRect();
        if (mr.right > window.innerWidth) menu.style.left = (window.innerWidth - mr.width - 8) + 'px';
      });
      setTimeout(() => {
        const close = (ev) => { if (!menu.contains(ev.target)) { menu.style.display = 'none'; document.removeEventListener('mousedown', close); } };
        document.addEventListener('mousedown', close);
      }, 0);
    });

    this.element.appendChild(wrap);
    return wrap;
  }

  _addColorPicker(title, iconHtml, styleProp, defaultColor) {
    const ss = this.spreadsheet;
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

    const wrap = el('div', { className: 'sheets-color-picker-wrap' });
    const btn = el('div', { className: 'sheets-color-btn', title });
    btn.innerHTML = iconHtml;
    const colorLine = el('div', { className: 'sheets-color-indicator' });
    colorLine.style.background = defaultColor;
    btn.appendChild(colorLine);
    wrap.appendChild(btn);

    const arrow = el('div', { className: 'sheets-color-arrow', title });
    arrow.innerHTML = '&#9662;';
    wrap.appendChild(arrow);

    const dropdown = el('div', { className: 'sheets-color-dropdown' });
    dropdown.style.display = 'none';
    dropdown.style.position = 'fixed';

    const resetBtn = el('div', { className: 'sheets-color-reset' });
    resetBtn.textContent = styleProp === 'bgColor' ? 'None' : 'Automatic';
    resetBtn.addEventListener('click', () => {
      ss.setStyle({ [styleProp]: styleProp === 'bgColor' ? null : '#000000' });
      colorLine.style.background = defaultColor;
      dropdown.style.display = 'none';
    });
    dropdown.appendChild(resetBtn);

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
    customBtn.addEventListener('click', (e) => { e.stopPropagation(); customInput.click(); });
    dropdown.appendChild(customBtn);
    document.body.appendChild(dropdown);

    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      ss.setStyle({ [styleProp]: colorLine.style.background || defaultColor });
    });

    arrow.addEventListener('click', (e) => {
      e.stopPropagation();
      if (dropdown.style.display !== 'none') { dropdown.style.display = 'none'; return; }
      document.querySelectorAll('.sheets-color-dropdown').forEach(d => d.style.display = 'none');
      const rect = arrow.getBoundingClientRect();
      dropdown.style.left = rect.left + 'px';
      dropdown.style.top = rect.bottom + 2 + 'px';
      dropdown.style.display = 'block';
      requestAnimationFrame(() => {
        const dr = dropdown.getBoundingClientRect();
        if (dr.right > window.innerWidth) dropdown.style.left = (window.innerWidth - dr.width - 8) + 'px';
      });
      setTimeout(() => {
        const close = (ev) => { if (!dropdown.contains(ev.target) && !arrow.contains(ev.target)) { dropdown.style.display = 'none'; document.removeEventListener('mousedown', close); } };
        document.addEventListener('mousedown', close);
      }, 0);
    });

    this.element.appendChild(wrap);
  }

  // ── Dropdown menus ──

  _buildFormatsMenu() {
    const ss = this.spreadsheet;
    return [
      { label: 'Automatic', action: () => ss.setStyle({ numberFormat: 'General' }) },
      null,
      { label: 'Number (1,000.12)', action: () => ss.setStyle({ numberFormat: '#,##0.00' }) },
      { label: 'Integer (1,000)', action: () => ss.setStyle({ numberFormat: '#,##0' }) },
      { label: 'Currency ($1,000.00)', action: () => ss.setStyle({ numberFormat: '$#,##0.00' }) },
      { label: 'Accounting', action: () => ss.setStyle({ numberFormat: '_($* #,##0.00_)' }) },
      null,
      { label: 'Percent (10.00%)', action: () => ss.setStyle({ numberFormat: '0.00%' }) },
      { label: 'Scientific (1.00E+3)', action: () => ss.setStyle({ numberFormat: '0.00E+0' }) },
      null,
      { label: 'Date (M/d/yyyy)', action: () => ss.setStyle({ numberFormat: 'M/d/yyyy' }) },
      { label: 'Date (MMMM d, yyyy)', action: () => ss.setStyle({ numberFormat: 'MMMM d, yyyy' }) },
      { label: 'Time (h:mm AM/PM)', action: () => ss.setStyle({ numberFormat: 'h:mm AM/PM' }) },
      { label: 'Date time', action: () => ss.setStyle({ numberFormat: 'M/d/yyyy h:mm' }) },
      null,
      { label: 'Plain text', action: () => ss.setStyle({ numberFormat: '@' }) },
    ];
  }

  _buildBordersMenu() {
    const ss = this.spreadsheet;
    const b = (color = '#000', width = 1) => ({ color, width });
    return [
      { icon: ICONS.borderAll, label: 'All borders', action: () => ss.setAllBorders(b()) },
      { icon: ICONS.borderInner, label: 'Inner borders', action: () => ss.setStyle({ borderBottom: b(), borderRight: b() }) },
      { icon: ICONS.borderOuter, label: 'Outer borders', action: () => {
        const sel = ss.selection; if (!sel) return;
        const sheet = ss.activeSheet; if (!sheet) return;
        for (let r = sel.startRow; r <= sel.endRow; r++) {
          for (let c = sel.startCol; c <= sel.endCol; c++) {
            const s = {};
            if (r === sel.startRow) s.borderTop = b();
            if (r === sel.endRow) s.borderBottom = b();
            if (c === sel.startCol) s.borderLeft = b();
            if (c === sel.endCol) s.borderRight = b();
            if (Object.keys(s).length) sheet.setCellStyle(r, c, s);
          }
        }
        ss.render();
      }},
      null,
      { icon: ICONS.borderBottom, label: 'Bottom border', action: () => ss.setStyle({ borderBottom: b() }) },
      { icon: ICONS.borderTop, label: 'Top border', action: () => ss.setStyle({ borderTop: b() }) },
      { icon: ICONS.borderLeft, label: 'Left border', action: () => ss.setStyle({ borderLeft: b() }) },
      { icon: ICONS.borderRight, label: 'Right border', action: () => ss.setStyle({ borderRight: b() }) },
      null,
      { icon: ICONS.borderClear, label: 'Clear borders', action: () => ss.clearBorders() },
    ];
  }

  _buildMergeMenu() {
    const ss = this.spreadsheet;
    return [
      { label: 'Merge all', action: () => ss.mergeSelection() },
      { label: 'Unmerge', action: () => ss.unmergeSelection() },
    ];
  }

  _buildHAlignMenu() {
    const ss = this.spreadsheet;
    return [
      { icon: ICONS.alignLeft, label: 'Left', action: () => ss.setStyle({ hAlign: H_ALIGN.LEFT }) },
      { icon: ICONS.alignCenter, label: 'Center', action: () => ss.setStyle({ hAlign: H_ALIGN.CENTER }) },
      { icon: ICONS.alignRight, label: 'Right', action: () => ss.setStyle({ hAlign: H_ALIGN.RIGHT }) },
    ];
  }

  _buildVAlignMenu() {
    const ss = this.spreadsheet;
    return [
      { icon: ICONS.valignTop, label: 'Top', action: () => ss.setStyle({ vAlign: V_ALIGN.TOP }) },
      { icon: ICONS.valignMiddle, label: 'Middle', action: () => ss.setStyle({ vAlign: V_ALIGN.MIDDLE }) },
      { icon: ICONS.valignBottom, label: 'Bottom', action: () => ss.setStyle({ vAlign: V_ALIGN.BOTTOM }) },
    ];
  }

  _buildWrapMenu() {
    const ss = this.spreadsheet;
    return [
      { icon: ICONS.wrapOverflow, label: 'Overflow', action: () => ss.setStyle({ wrap: false }) },
      { icon: ICONS.wrapWrap, label: 'Wrap', action: () => ss.setStyle({ wrap: true }) },
      { icon: ICONS.wrapClip, label: 'Clip', action: () => ss.setStyle({ wrap: false }) },
    ];
  }

  _buildRotationMenu() {
    const ss = this.spreadsheet;
    return [
      { label: 'None', action: () => ss.setStyle({ rotation: 0 }) },
      { label: 'Tilt up', action: () => ss.setStyle({ rotation: 45 }) },
      { label: 'Tilt down', action: () => ss.setStyle({ rotation: -45 }) },
      { label: 'Rotate up', action: () => ss.setStyle({ rotation: 90 }) },
      { label: 'Rotate down', action: () => ss.setStyle({ rotation: -90 }) },
    ];
  }

  _buildFunctionsMenu() {
    const ss = this.spreadsheet;
    const insert = (fn) => {
      if (ss.editor) { ss.editor.beginEnter('=' + fn + '('); }
    };
    return [
      { label: 'SUM', action: () => insert('SUM') },
      { label: 'AVERAGE', action: () => insert('AVERAGE') },
      { label: 'COUNT', action: () => insert('COUNT') },
      { label: 'MAX', action: () => insert('MAX') },
      { label: 'MIN', action: () => insert('MIN') },
    ];
  }

  _insertLink() {
    const url = prompt('Enter URL:');
    if (url) {
      const ss = this.spreadsheet;
      const sheet = ss.activeSheet;
      if (!sheet) return;
      const cell = sheet.getOrCreateCell(ss.activeRow, ss.activeCol);
      if (!cell.rawValue) sheet.setCellValue(ss.activeRow, ss.activeCol, url);
      cell.comment = url; // store URL as comment for now
      ss.render();
    }
  }

  _insertComment() {
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    if (!sheet) return;
    const cell = sheet.getOrCreateCell(ss.activeRow, ss.activeCol);
    const existing = cell.comment || '';
    const comment = prompt('Comment:', existing);
    if (comment !== null) {
      cell.comment = comment || null;
      ss.render();
    }
  }

  _toast(msg) {
    const t = document.getElementById('toast');
    if (t) { t.textContent = msg; t.classList.add('show'); setTimeout(() => t.classList.remove('show'), 2000); }
  }

  destroy() {
    if (this.element && this.element.parentElement) {
      this.element.parentElement.removeChild(this.element);
    }
  }
}

// ── SVG Icons (Google Sheets style, 16x16) ──
const ICONS = {
  undo: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 10h10a5 5 0 0 1 0 10H12"/><path d="M3 10l4-4M3 10l4 4"/></svg>`,
  redo: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 10H11a5 5 0 0 0 0 10h1"/><path d="M21 10l-4-4M21 10l-4 4"/></svg>`,
  print: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 9V2h12v7"/><path d="M6 18H4a2 2 0 01-2-2v-5a2 2 0 012-2h16a2 2 0 012 2v5a2 2 0 01-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>`,
  decimalDec: `<svg width="16" height="16" viewBox="0 0 24 24"><text x="1" y="15" font-size="11" fill="currentColor">.0</text><path d="M17 8l4 4-4 4" fill="none" stroke="currentColor" stroke-width="1.5"/></svg>`,
  decimalInc: `<svg width="16" height="16" viewBox="0 0 24 24"><text x="0" y="15" font-size="10" fill="currentColor">.00</text><path d="M19 8l-4 4 4 4" fill="none" stroke="currentColor" stroke-width="1.5"/></svg>`,
  fillBucket: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M16.56 8.94L7.62 0 6.21 1.41l2.38 2.38-5.15 5.15a1.49 1.49 0 000 2.12l5.5 5.5c.29.29.68.44 1.06.44s.77-.15 1.06-.44l5.5-5.5c.59-.58.59-1.53 0-2.12zM5.21 10L10 5.21 14.79 10H5.21zM19 11.5s-2 2.17-2 3.5c0 1.1.9 2 2 2s2-.9 2-2c0-1.33-2-3.5-2-3.5z"/></svg>`,
  borders: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M3 3v18h18V3H3zm4 16H5v-2h2v2zm0-4H5v-2h2v2zm0-4H5V9h2v2zm0-4H5V5h2v2zm4 12H9v-2h2v2zm0-4H9v-2h2v2zm0-4H9V9h2v2zm0-4H9V5h2v2zm8 12h-6v-2h2v-2h-2v-2h2v-2h-2V9h2V7h-2V5h6v14z"/></svg>`,
  borderAll: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M3 3v18h18V3H3zm8 16H5v-6h6v6zm0-8H5V5h6v6zm8 8h-6v-6h6v6zm0-8h-6V5h6v6z"/></svg>`,
  borderInner: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M3 21h2V3H3v18zm4-10h4V3H7v8zm6-8v8h4V3h-4zm6 0v18h2V3h-2zM7 13v8h4v-8H7zm6 0v8h4v-8h-4zM3 21h18v-2H3v2zM3 3v2h18V3H3z" opacity=".5"/><path d="M11 3h2v18h-2zM3 11h18v2H3z"/></svg>`,
  borderOuter: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M3 3v18h18V3H3zm16 16H5V5h14v14z"/></svg>`,
  borderBottom: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M3 19h18v2H3z"/><path d="M5 13h2v2H5zm4 0h2v2H9zm4 0h2v2h-2zm4 0h2v2h-2zm-8-4h2v2H9zm4 0h2v2h-2zM5 5h2v2H5zm4 0h2v2H9zm4 0h2v2h-2zm4 0h2v2h-2zm0 4h2v2h-2zM5 9h2v2H5z" opacity=".3"/></svg>`,
  borderTop: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M3 3h18v2H3z"/><path d="M5 13h2v2H5zm4 0h2v2H9zm4 0h2v2h-2zm4 0h2v2h-2zm-8-4h2v2H9zm4 0h2v2h-2zM5 17h2v2H5zm4 0h2v2H9zm4 0h2v2h-2zm4 0h2v2h-2zm0-8h2v2h-2zM5 9h2v2H5z" opacity=".3"/></svg>`,
  borderLeft: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M3 3h2v18H3z"/><path d="M9 13h2v2H9zm4 0h2v2h-2zm4 0h2v2h-2zm-4-4h2v2h-2zM9 5h2v2H9zm4 0h2v2h-2zm4 0h2v2h-2zm0 4h2v2h-2zm0 8h2v2h-2zM9 17h2v2H9zm4 0h2v2h-2zM9 9h2v2H9z" opacity=".3"/></svg>`,
  borderRight: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M19 3h2v18h-2z"/><path d="M5 13h2v2H5zm4 0h2v2H9zm4 0h2v2h-2zm-8-4h2v2H5zm4 0h2v2H9zm4 0h2v2h-2zM5 5h2v2H5zm4 0h2v2H9zm4 0h2v2h-2zM5 17h2v2H5zm4 0h2v2H9zm4 0h2v2h-2zm-4-8h2v2H9z" opacity=".3"/></svg>`,
  borderClear: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" opacity=".5"><path d="M5 13h2v2H5zm4 0h2v2H9zm4 0h2v2h-2zm4 0h2v2h-2zm-8-4h2v2H9zm4 0h2v2h-2zM5 5h2v2H5zm4 0h2v2H9zm4 0h2v2h-2zm4 0h2v2h-2zm0 4h2v2h-2zM5 9h2v2H5zm0 8h2v2H5zm4 0h2v2H9zm4 0h2v2h-2zm4 0h2v2h-2z"/></svg>`,
  merge: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="1"/><path d="M9 12h6m-3-3l3 3-3 3"/></svg>`,
  alignLeft: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M3 3h18v2H3zm0 4h12v2H3zm0 4h18v2H3zm0 4h12v2H3zm0 4h18v2H3z"/></svg>`,
  alignCenter: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M3 3h18v2H3zm3 4h12v2H6zM3 11h18v2H3zm3 4h12v2H6zM3 19h18v2H3z"/></svg>`,
  alignRight: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M3 3h18v2H3zm6 4h12v2H9zM3 11h18v2H3zm6 4h12v2H9zM3 19h18v2H3z"/></svg>`,
  valignTop: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M3 3h18v2H3zm5 4h8v2H8zm0 4h8v2H8z"/></svg>`,
  valignMiddle: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M8 8h8v2H8zm0 6h8v2H8zm-5-3h18v2H3z"/></svg>`,
  valignBottom: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M8 13h8v2H8zm0 4h8v2H8zM3 19h18v2H3z"/></svg>`,
  wrapOverflow: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M3 12h18M3 18h18"/></svg>`,
  wrapWrap: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M3 12h15a3 3 0 110 6h-4"/><path d="M16 16l-2 2 2 2"/><path d="M3 18h7"/></svg>`,
  wrapClip: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M3 12h12M3 18h12"/><path d="M17 10v8" stroke-dasharray="2 2"/></svg>`,
  textRotation: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M13 6.99l-6 6 1.41 1.41L13 9.82V22h2V6.99h-2zM16.71 4.29l-1.42 1.42L16.59 7H7v2h9.59l-1.3 1.29 1.42 1.42 3.29-3.29-3.29-3.13z" transform="scale(0.75) translate(4,4)"/></svg>`,
  link: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10 13a5 5 0 007.54.54l3-3a5 5 0 00-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 00-7.54-.54l-3 3a5 5 0 007.07 7.07l1.71-1.71"/></svg>`,
  comment: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"/></svg>`,
  chart: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M7 17v-4m4 4V9m4 8v-6m4 6v-2"/></svg>`,
  filter: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 3H2l8 9.46V19l4 2v-8.54L22 3z"/></svg>`,
  formatPainter: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M18 4V3c0-.55-.45-1-1-1H5c-.55 0-1 .45-1 1v4c0 .55.45 1 1 1h12c.55 0 1-.45 1-1V6h1v4H9v11c0 .55.45 1 1 1h2c.55 0 1-.45 1-1v-9h8V4h-3z"/></svg>`,
  functions: `<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M5 3h14v2H5zm0 16h14v2H5zm4-6.5L5.5 8H8l2 3 2-3h2.5L11 12.5 14.5 17H12l-2-3-2 3H5.5l3.5-4.5z" transform="scale(0.85) translate(2,2)"/></svg>`,
};
