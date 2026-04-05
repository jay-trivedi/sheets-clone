import { COLUMN_LETTERS, MAX_COLS } from './constants.js';

export function colToIndex(col) {
  let index = 0;
  for (let i = 0; i < col.length; i++) {
    index = index * 26 + (col.charCodeAt(i) - 64);
  }
  return index - 1;
}

export function indexToCol(index) {
  return COLUMN_LETTERS[index] || '';
}

export function parseCellRef(ref) {
  const match = ref.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/i);
  if (!match) return null;
  return {
    colAbs: match[1] === '$',
    col: colToIndex(match[2].toUpperCase()),
    rowAbs: match[3] === '$',
    row: parseInt(match[4], 10) - 1,
  };
}

export function cellRefToString(col, row, colAbs = false, rowAbs = false) {
  return (colAbs ? '$' : '') + indexToCol(col) + (rowAbs ? '$' : '') + (row + 1);
}

export function parseRangeRef(ref) {
  const parts = ref.split(':');
  if (parts.length !== 2) return null;
  const start = parseCellRef(parts[0]);
  const end = parseCellRef(parts[1]);
  if (!start || !end) return null;
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

export function rangeToString(startCol, startRow, endCol, endRow) {
  return cellRefToString(startCol, startRow) + ':' + cellRefToString(endCol, endRow);
}

export function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

export function deepClone(obj) {
  if (obj === null || typeof obj !== 'object') return obj;
  if (Array.isArray(obj)) return obj.map(deepClone);
  const clone = {};
  for (const key in obj) {
    if (Object.prototype.hasOwnProperty.call(obj, key)) {
      clone[key] = deepClone(obj[key]);
    }
  }
  return clone;
}

export function debounce(fn, ms) {
  let timer;
  return function (...args) {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), ms);
  };
}

export function throttle(fn, ms) {
  let last = 0;
  let timer;
  return function (...args) {
    const now = Date.now();
    clearTimeout(timer);
    if (now - last >= ms) {
      last = now;
      fn.apply(this, args);
    } else {
      timer = setTimeout(() => {
        last = Date.now();
        fn.apply(this, args);
      }, ms - (now - last));
    }
  };
}

export function isNumeric(val) {
  if (typeof val === 'number') return !isNaN(val);
  if (typeof val !== 'string' || val.trim() === '') return false;
  return !isNaN(Number(val));
}

export function toNumber(val) {
  if (typeof val === 'number') return val;
  if (typeof val === 'boolean') return val ? 1 : 0;
  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (trimmed === '') return 0;
    const pct = trimmed.match(/^(-?\d+\.?\d*)%$/);
    if (pct) return parseFloat(pct[1]) / 100;
    const currency = trimmed.replace(/^[$£€¥]/, '').replace(/,/g, '');
    const n = Number(currency);
    return isNaN(n) ? NaN : n;
  }
  return NaN;
}

export function cellKey(row, col) {
  return (row << 16) | col;
}

export function keyToRC(key) {
  return { row: key >> 16, col: key & 0xffff };
}

export function rectIntersects(r1, r2) {
  return r1.left < r2.right && r1.right > r2.left && r1.top < r2.bottom && r1.bottom > r2.top;
}

export function el(tag, attrs, ...children) {
  const elem = document.createElement(tag);
  if (attrs) {
    for (const [k, v] of Object.entries(attrs)) {
      if (k === 'style' && typeof v === 'object') {
        Object.assign(elem.style, v);
      } else if (k === 'className') {
        elem.className = v;
      } else if (k.startsWith('on')) {
        elem.addEventListener(k.slice(2).toLowerCase(), v);
      } else {
        elem.setAttribute(k, v);
      }
    }
  }
  for (const child of children) {
    if (typeof child === 'string') {
      elem.appendChild(document.createTextNode(child));
    } else if (child) {
      elem.appendChild(child);
    }
  }
  return elem;
}

export function cssColor(r, g, b, a = 1) {
  return a === 1 ? `rgb(${r},${g},${b})` : `rgba(${r},${g},${b},${a})`;
}

export function hexToRgb(hex) {
  const h = hex.replace('#', '');
  return {
    r: parseInt(h.substring(0, 2), 16),
    g: parseInt(h.substring(2, 4), 16),
    b: parseInt(h.substring(4, 6), 16),
  };
}

export function generateId() {
  return Math.random().toString(36).substring(2, 10);
}
