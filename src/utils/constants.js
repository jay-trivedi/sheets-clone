export const DEFAULT_COL_WIDTH = 100;
export const DEFAULT_ROW_HEIGHT = 25;
export const ROW_HEADER_WIDTH = 50;
export const COL_HEADER_HEIGHT = 25;
export const MIN_COL_WIDTH = 20;
export const MIN_ROW_HEIGHT = 15;
export const MAX_COLS = 26 * 27; // A-ZZ = 702 columns
export const MAX_ROWS = 100000;
export const DEFAULT_ROWS = 1000;
export const DEFAULT_COLS = 26;
export const SCROLLBAR_SIZE = 14;

export const CELL_PADDING = 4;
export const CELL_BORDER_COLOR = '#dadada';
export const HEADER_BG_COLOR = '#efeded';
export const HEADER_BORDER_COLOR = '#c4c7c5';
export const HEADER_TEXT_COLOR = '#1f1f1f';
export const SELECTION_BORDER_COLOR = '#3271ea';
export const SELECTION_BG_COLOR = 'rgba(50, 113, 234, 0.08)';
export const SELECTION_HEADER_BG = '#d3e3fd';
export const FILL_HANDLE_SIZE = 7;
export const FILL_HANDLE_COLOR = '#3271ea';
export const GRID_LINE_WIDTH = 1;

export const DEFAULT_FONT_FAMILY = 'Arial';
export const DEFAULT_FONT_SIZE = 10;
export const DEFAULT_TEXT_COLOR = '#1f1f1f';
export const DEFAULT_BG_COLOR = '#ffffff';

export const TOOLBAR_HEIGHT = 40;
export const FORMULA_BAR_HEIGHT = 30;
export const SHEET_TAB_HEIGHT = 30;
export const STATUS_BAR_HEIGHT = 24;

export const DOUBLE_CLICK_MS = 300;
export const AUTOSCROLL_SPEED = 15;
export const AUTOSCROLL_INTERVAL = 50;

export const FREEZE_LINE_COLOR = '#bababa';
export const FREEZE_LINE_WIDTH = 2;

export const COLUMN_LETTERS = (() => {
  const letters = [];
  for (let i = 0; i < MAX_COLS; i++) {
    let s = '';
    let n = i;
    while (n >= 0) {
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26) - 1;
    }
    letters.push(s);
  }
  return letters;
})();

export const KEY = {
  ENTER: 'Enter',
  TAB: 'Tab',
  ESCAPE: 'Escape',
  DELETE: 'Delete',
  BACKSPACE: 'Backspace',
  ARROW_UP: 'ArrowUp',
  ARROW_DOWN: 'ArrowDown',
  ARROW_LEFT: 'ArrowLeft',
  ARROW_RIGHT: 'ArrowRight',
  HOME: 'Home',
  END: 'End',
  PAGE_UP: 'PageUp',
  PAGE_DOWN: 'PageDown',
  SPACE: ' ',
  F2: 'F2',
};

export const CELL_TYPE = {
  EMPTY: 0,
  STRING: 1,
  NUMBER: 2,
  BOOLEAN: 3,
  ERROR: 4,
  FORMULA: 5,
};

export const ERROR_TYPE = {
  NULL: '#NULL!',
  DIV0: '#DIV/0!',
  VALUE: '#VALUE!',
  REF: '#REF!',
  NAME: '#NAME?',
  NUM: '#NUM!',
  NA: '#N/A',
  CIRCULAR: '#CIRCULAR!',
};

export const H_ALIGN = { LEFT: 'left', CENTER: 'center', RIGHT: 'right' };
export const V_ALIGN = { TOP: 'top', MIDDLE: 'middle', BOTTOM: 'bottom' };

export const BORDER_STYLE = {
  NONE: 0,
  THIN: 1,
  MEDIUM: 2,
  THICK: 3,
  DASHED: 4,
  DOTTED: 5,
};

export const NUMBER_FORMATS = {
  GENERAL: 'General',
  NUMBER: '#,##0.00',
  INTEGER: '#,##0',
  CURRENCY: '$#,##0.00',
  PERCENT: '0.00%',
  SCIENTIFIC: '0.00E+0',
  DATE_SHORT: 'M/d/yyyy',
  DATE_LONG: 'MMMM d, yyyy',
  TIME: 'h:mm:ss AM/PM',
  DATETIME: 'M/d/yyyy h:mm',
  TEXT: '@',
  ACCOUNTING: '_($* #,##0.00_)',
};
