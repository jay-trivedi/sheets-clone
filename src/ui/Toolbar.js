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

    // Text color
    const textColorBtn = el('div', { className: 'sheets-color-btn', title: 'Text color' });
    textColorBtn.innerHTML = '<span style="font-weight:bold">A</span>';
    const textColorLine = el('div', { style: { height: '3px', background: '#000', marginTop: '1px' } });
    textColorBtn.appendChild(textColorLine);
    const textInput = el('input', { type: 'color', value: '#000000', style: { position: 'absolute', opacity: '0', width: '0', height: '0' } });
    textInput.addEventListener('input', () => {
      ss.setStyle({ textColor: textInput.value });
      textColorLine.style.background = textInput.value;
    });
    textColorBtn.appendChild(textInput);
    textColorBtn.addEventListener('click', () => textInput.click());
    group.appendChild(textColorBtn);

    // Background color
    const bgColorBtn = el('div', { className: 'sheets-color-btn', title: 'Fill color' });
    bgColorBtn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M19 11h-6V5a1 1 0 0 0-2 0v6H5a1 1 0 0 0 0 2h6v6a1 1 0 0 0 2 0v-6h6a1 1 0 0 0 0-2z"/></svg>`;
    const bgColorLine = el('div', { style: { height: '3px', background: '#fff', border: '1px solid #ccc', marginTop: '1px' } });
    bgColorBtn.appendChild(bgColorLine);
    const bgInput = el('input', { type: 'color', value: '#ffffff', style: { position: 'absolute', opacity: '0', width: '0', height: '0' } });
    bgInput.addEventListener('input', () => {
      ss.setStyle({ bgColor: bgInput.value });
      bgColorLine.style.background = bgInput.value;
    });
    bgColorBtn.appendChild(bgInput);
    bgColorBtn.addEventListener('click', () => bgInput.click());
    group.appendChild(bgColorBtn);

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
