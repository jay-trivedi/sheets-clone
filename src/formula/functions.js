import { ERROR_TYPE } from '../utils/constants.js';
import { parseCellRef } from '../utils/helpers.js';

function isError(v) { return typeof v === 'string' && v.startsWith('#'); }
function numOrErr(v, ctx) {
  if (isError(v)) return v;
  const n = ctx.toNumber(v);
  return isNaN(n) ? ERROR_TYPE.VALUE : n;
}

export const FUNCTIONS = {
  // ── Math ──
  SUM(args, ctx) {
    const nums = ctx.flattenNumbers(args);
    return nums.reduce((a, b) => a + b, 0);
  },
  SUMIF(args, ctx) {
    if (args.length < 2) return ERROR_TYPE.VALUE;
    const range = ctx.resolveRange(args[0]);
    const criteria = args[1];
    const sumRange = args.length > 2 ? ctx.resolveRange(args[2]) : range;
    let sum = 0;
    for (let r = 0; r < range.length; r++) {
      for (let c = 0; c < range[r].length; c++) {
        if (matchCriteria(range[r][c], criteria)) {
          const sv = sumRange[r] && sumRange[r][c] !== undefined ? sumRange[r][c] : 0;
          const n = typeof sv === 'number' ? sv : parseFloat(sv);
          if (!isNaN(n)) sum += n;
        }
      }
    }
    return sum;
  },
  SUMIFS(args, ctx) {
    if (args.length < 3 || args.length % 2 === 0) return ERROR_TYPE.VALUE;
    const sumRange = ctx.resolveRange(args[0]);
    let sum = 0;
    const pairs = [];
    for (let i = 1; i < args.length; i += 2) {
      pairs.push({ range: ctx.resolveRange(args[i]), criteria: args[i + 1] });
    }
    for (let r = 0; r < sumRange.length; r++) {
      for (let c = 0; c < (sumRange[r] || []).length; c++) {
        let match = true;
        for (const { range, criteria } of pairs) {
          if (!range[r] || !matchCriteria(range[r][c], criteria)) { match = false; break; }
        }
        if (match) {
          const n = ctx.toNumber(sumRange[r][c]);
          if (!isNaN(n)) sum += n;
        }
      }
    }
    return sum;
  },
  SUMPRODUCT(args, ctx) {
    const arrays = args.map(a => ctx.resolveRange(a));
    const rows = arrays[0].length;
    const cols = arrays[0][0] ? arrays[0][0].length : 0;
    let sum = 0;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        let product = 1;
        for (const arr of arrays) {
          const v = arr[r] && arr[r][c] != null ? ctx.toNumber(arr[r][c]) : 0;
          product *= isNaN(v) ? 0 : v;
        }
        sum += product;
      }
    }
    return sum;
  },
  PRODUCT(args, ctx) {
    const nums = ctx.flattenNumbers(args);
    if (nums.length === 0) return 0;
    return nums.reduce((a, b) => a * b, 1);
  },
  ABS(args, ctx) {
    return Math.abs(numOrErr(args[0], ctx));
  },
  SQRT(args, ctx) {
    const n = numOrErr(args[0], ctx);
    if (isError(n)) return n;
    return n < 0 ? ERROR_TYPE.NUM : Math.sqrt(n);
  },
  POWER(args, ctx) {
    const base = numOrErr(args[0], ctx);
    const exp = numOrErr(args[1], ctx);
    if (isError(base)) return base;
    if (isError(exp)) return exp;
    return Math.pow(base, exp);
  },
  MOD(args, ctx) {
    const n = numOrErr(args[0], ctx);
    const d = numOrErr(args[1], ctx);
    if (isError(n)) return n;
    if (isError(d)) return d;
    if (d === 0) return ERROR_TYPE.DIV0;
    return n - d * Math.floor(n / d);
  },
  ROUND(args, ctx) {
    const n = numOrErr(args[0], ctx);
    const d = args.length > 1 ? numOrErr(args[1], ctx) : 0;
    if (isError(n)) return n;
    if (isError(d)) return d;
    const factor = Math.pow(10, d);
    return Math.round(n * factor) / factor;
  },
  ROUNDUP(args, ctx) {
    const n = numOrErr(args[0], ctx);
    const d = args.length > 1 ? numOrErr(args[1], ctx) : 0;
    if (isError(n)) return n;
    const factor = Math.pow(10, d);
    return Math.sign(n) * Math.ceil(Math.abs(n) * factor) / factor;
  },
  ROUNDDOWN(args, ctx) {
    const n = numOrErr(args[0], ctx);
    const d = args.length > 1 ? numOrErr(args[1], ctx) : 0;
    if (isError(n)) return n;
    const factor = Math.pow(10, d);
    return Math.sign(n) * Math.floor(Math.abs(n) * factor) / factor;
  },
  CEILING(args, ctx) {
    const n = numOrErr(args[0], ctx);
    const s = args.length > 1 ? numOrErr(args[1], ctx) : 1;
    if (isError(n)) return n;
    if (s === 0) return 0;
    return Math.ceil(n / s) * s;
  },
  FLOOR(args, ctx) {
    const n = numOrErr(args[0], ctx);
    const s = args.length > 1 ? numOrErr(args[1], ctx) : 1;
    if (isError(n)) return n;
    if (s === 0) return ERROR_TYPE.DIV0;
    return Math.floor(n / s) * s;
  },
  INT(args, ctx) {
    const n = numOrErr(args[0], ctx);
    if (isError(n)) return n;
    return Math.floor(n);
  },
  SIGN(args, ctx) {
    const n = numOrErr(args[0], ctx);
    if (isError(n)) return n;
    return Math.sign(n);
  },
  LN(args, ctx) {
    const n = numOrErr(args[0], ctx);
    if (isError(n)) return n;
    return n <= 0 ? ERROR_TYPE.NUM : Math.log(n);
  },
  LOG(args, ctx) {
    const n = numOrErr(args[0], ctx);
    const base = args.length > 1 ? numOrErr(args[1], ctx) : 10;
    if (isError(n)) return n;
    if (n <= 0 || base <= 0 || base === 1) return ERROR_TYPE.NUM;
    return Math.log(n) / Math.log(base);
  },
  LOG10(args, ctx) {
    const n = numOrErr(args[0], ctx);
    if (isError(n)) return n;
    return n <= 0 ? ERROR_TYPE.NUM : Math.log10(n);
  },
  EXP(args, ctx) {
    const n = numOrErr(args[0], ctx);
    if (isError(n)) return n;
    return Math.exp(n);
  },
  PI() { return Math.PI; },
  RAND() { return Math.random(); },
  RANDBETWEEN(args, ctx) {
    const lo = numOrErr(args[0], ctx);
    const hi = numOrErr(args[1], ctx);
    if (isError(lo)) return lo;
    if (isError(hi)) return hi;
    return Math.floor(Math.random() * (Math.floor(hi) - Math.ceil(lo) + 1)) + Math.ceil(lo);
  },
  SIN(args, ctx) { const n = numOrErr(args[0], ctx); return isError(n) ? n : Math.sin(n); },
  COS(args, ctx) { const n = numOrErr(args[0], ctx); return isError(n) ? n : Math.cos(n); },
  TAN(args, ctx) { const n = numOrErr(args[0], ctx); return isError(n) ? n : Math.tan(n); },
  ASIN(args, ctx) { const n = numOrErr(args[0], ctx); return isError(n) ? n : Math.asin(n); },
  ACOS(args, ctx) { const n = numOrErr(args[0], ctx); return isError(n) ? n : Math.acos(n); },
  ATAN(args, ctx) { const n = numOrErr(args[0], ctx); return isError(n) ? n : Math.atan(n); },
  ATAN2(args, ctx) {
    const y = numOrErr(args[0], ctx);
    const x = numOrErr(args[1], ctx);
    if (isError(y)) return y;
    if (isError(x)) return x;
    return Math.atan2(y, x);
  },
  DEGREES(args, ctx) { const n = numOrErr(args[0], ctx); return isError(n) ? n : n * 180 / Math.PI; },
  RADIANS(args, ctx) { const n = numOrErr(args[0], ctx); return isError(n) ? n : n * Math.PI / 180; },
  EVEN(args, ctx) {
    const n = numOrErr(args[0], ctx);
    if (isError(n)) return n;
    const ceil = Math.ceil(Math.abs(n));
    const result = ceil % 2 === 0 ? ceil : ceil + 1;
    return n < 0 ? -result : result;
  },
  ODD(args, ctx) {
    const n = numOrErr(args[0], ctx);
    if (isError(n)) return n;
    const ceil = Math.ceil(Math.abs(n));
    const result = ceil % 2 === 1 ? ceil : ceil + 1;
    return n < 0 ? -result : result;
  },
  FACT(args, ctx) {
    const n = numOrErr(args[0], ctx);
    if (isError(n)) return n;
    if (n < 0) return ERROR_TYPE.NUM;
    let f = 1;
    for (let i = 2; i <= Math.floor(n); i++) f *= i;
    return f;
  },
  GCD(args, ctx) {
    const nums = ctx.flattenNumbers(args);
    const gcd = (a, b) => b === 0 ? a : gcd(b, a % b);
    return nums.reduce((a, b) => gcd(Math.abs(Math.round(a)), Math.abs(Math.round(b))));
  },
  LCM(args, ctx) {
    const nums = ctx.flattenNumbers(args);
    const gcd = (a, b) => b === 0 ? a : gcd(b, a % b);
    return nums.reduce((a, b) => {
      a = Math.abs(Math.round(a)); b = Math.abs(Math.round(b));
      return (a * b) / gcd(a, b);
    });
  },

  // ── Statistical ──
  AVERAGE(args, ctx) {
    const nums = ctx.flattenNumbers(args);
    if (nums.length === 0) return ERROR_TYPE.DIV0;
    return nums.reduce((a, b) => a + b, 0) / nums.length;
  },
  AVERAGEIF(args, ctx) {
    if (args.length < 2) return ERROR_TYPE.VALUE;
    const range = ctx.resolveRange(args[0]);
    const criteria = args[1];
    const avgRange = args.length > 2 ? ctx.resolveRange(args[2]) : range;
    let sum = 0, count = 0;
    for (let r = 0; r < range.length; r++) {
      for (let c = 0; c < range[r].length; c++) {
        if (matchCriteria(range[r][c], criteria)) {
          const v = avgRange[r] && avgRange[r][c] != null ? ctx.toNumber(avgRange[r][c]) : NaN;
          if (!isNaN(v)) { sum += v; count++; }
        }
      }
    }
    return count === 0 ? ERROR_TYPE.DIV0 : sum / count;
  },
  MEDIAN(args, ctx) {
    const nums = ctx.flattenNumbers(args).sort((a, b) => a - b);
    if (nums.length === 0) return ERROR_TYPE.NUM;
    const mid = Math.floor(nums.length / 2);
    return nums.length % 2 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
  },
  MODE(args, ctx) {
    const nums = ctx.flattenNumbers(args);
    const freq = new Map();
    let maxCount = 0, mode = nums[0];
    for (const n of nums) {
      const c = (freq.get(n) || 0) + 1;
      freq.set(n, c);
      if (c > maxCount) { maxCount = c; mode = n; }
    }
    return maxCount <= 1 ? ERROR_TYPE.NA : mode;
  },
  COUNT(args, ctx) {
    const flat = ctx.flattenArgs(args);
    return flat.filter(v => v !== null && v !== undefined && v !== '' && typeof v !== 'boolean' && !isNaN(Number(v))).length;
  },
  COUNTA(args, ctx) {
    const flat = ctx.flattenArgs(args);
    return flat.filter(v => v !== null && v !== undefined && v !== '').length;
  },
  COUNTBLANK(args, ctx) {
    const flat = ctx.flattenArgs(args);
    return flat.filter(v => v === null || v === undefined || v === '').length;
  },
  COUNTIF(args, ctx) {
    const range = ctx.resolveRange(args[0]);
    const criteria = args[1];
    let count = 0;
    for (const row of range) for (const v of row) if (matchCriteria(v, criteria)) count++;
    return count;
  },
  COUNTIFS(args, ctx) {
    if (args.length < 2 || args.length % 2 !== 0) return ERROR_TYPE.VALUE;
    const pairs = [];
    for (let i = 0; i < args.length; i += 2) {
      pairs.push({ range: ctx.resolveRange(args[i]), criteria: args[i + 1] });
    }
    const rows = pairs[0].range.length;
    const cols = pairs[0].range[0] ? pairs[0].range[0].length : 0;
    let count = 0;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        let match = true;
        for (const { range, criteria } of pairs) {
          if (!range[r] || !matchCriteria(range[r][c], criteria)) { match = false; break; }
        }
        if (match) count++;
      }
    }
    return count;
  },
  MAX(args, ctx) {
    const nums = ctx.flattenNumbers(args);
    return nums.length === 0 ? 0 : Math.max(...nums);
  },
  MIN(args, ctx) {
    const nums = ctx.flattenNumbers(args);
    return nums.length === 0 ? 0 : Math.min(...nums);
  },
  LARGE(args, ctx) {
    const nums = ctx.flattenNumbers([args[0]]).sort((a, b) => b - a);
    const k = ctx.toNumber(args[1]);
    return k < 1 || k > nums.length ? ERROR_TYPE.NUM : nums[Math.floor(k) - 1];
  },
  SMALL(args, ctx) {
    const nums = ctx.flattenNumbers([args[0]]).sort((a, b) => a - b);
    const k = ctx.toNumber(args[1]);
    return k < 1 || k > nums.length ? ERROR_TYPE.NUM : nums[Math.floor(k) - 1];
  },
  STDEV(args, ctx) {
    const nums = ctx.flattenNumbers(args);
    if (nums.length < 2) return ERROR_TYPE.DIV0;
    const mean = nums.reduce((a, b) => a + b) / nums.length;
    const variance = nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (nums.length - 1);
    return Math.sqrt(variance);
  },
  STDEVP(args, ctx) {
    const nums = ctx.flattenNumbers(args);
    if (nums.length === 0) return ERROR_TYPE.DIV0;
    const mean = nums.reduce((a, b) => a + b) / nums.length;
    const variance = nums.reduce((a, b) => a + (b - mean) ** 2, 0) / nums.length;
    return Math.sqrt(variance);
  },
  VAR(args, ctx) {
    const nums = ctx.flattenNumbers(args);
    if (nums.length < 2) return ERROR_TYPE.DIV0;
    const mean = nums.reduce((a, b) => a + b) / nums.length;
    return nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (nums.length - 1);
  },
  VARP(args, ctx) {
    const nums = ctx.flattenNumbers(args);
    if (nums.length === 0) return ERROR_TYPE.DIV0;
    const mean = nums.reduce((a, b) => a + b) / nums.length;
    return nums.reduce((a, b) => a + (b - mean) ** 2, 0) / nums.length;
  },
  PERCENTILE(args, ctx) {
    const nums = ctx.flattenNumbers([args[0]]).sort((a, b) => a - b);
    const k = ctx.toNumber(args[1]);
    if (k < 0 || k > 1) return ERROR_TYPE.NUM;
    const idx = k * (nums.length - 1);
    const lo = Math.floor(idx), hi = Math.ceil(idx);
    return lo === hi ? nums[lo] : nums[lo] + (nums[hi] - nums[lo]) * (idx - lo);
  },
  RANK(args, ctx) {
    const val = ctx.toNumber(args[0]);
    const nums = ctx.flattenNumbers([args[1]]);
    const order = args.length > 2 ? ctx.toNumber(args[2]) : 0;
    const sorted = order ? [...nums].sort((a, b) => a - b) : [...nums].sort((a, b) => b - a);
    const idx = sorted.indexOf(val);
    return idx === -1 ? ERROR_TYPE.NA : idx + 1;
  },
  CORREL(args, ctx) {
    const x = ctx.flattenNumbers([args[0]]);
    const y = ctx.flattenNumbers([args[1]]);
    const n = Math.min(x.length, y.length);
    if (n < 2) return ERROR_TYPE.DIV0;
    const mx = x.reduce((a, b) => a + b) / n;
    const my = y.reduce((a, b) => a + b) / n;
    let num = 0, dx = 0, dy = 0;
    for (let i = 0; i < n; i++) {
      num += (x[i] - mx) * (y[i] - my);
      dx += (x[i] - mx) ** 2;
      dy += (y[i] - my) ** 2;
    }
    return dx === 0 || dy === 0 ? ERROR_TYPE.DIV0 : num / Math.sqrt(dx * dy);
  },

  // ── Logical ──
  IF(args, ctx) {
    const cond = ctx.toBoolean(args[0]);
    return cond ? args[1] : (args.length > 2 ? args[2] : false);
  },
  IFS(args, ctx) {
    for (let i = 0; i < args.length; i += 2) {
      if (ctx.toBoolean(args[i])) return args[i + 1];
    }
    return ERROR_TYPE.NA;
  },
  AND(args, ctx) {
    const flat = ctx.flattenArgs(args);
    for (const v of flat) {
      if (v === null || v === undefined || v === '') continue;
      if (!ctx.toBoolean(v)) return false;
    }
    return true;
  },
  OR(args, ctx) {
    const flat = ctx.flattenArgs(args);
    for (const v of flat) {
      if (v === null || v === undefined || v === '') continue;
      if (ctx.toBoolean(v)) return true;
    }
    return false;
  },
  NOT(args, ctx) {
    return !ctx.toBoolean(args[0]);
  },
  XOR(args, ctx) {
    const flat = ctx.flattenArgs(args);
    let count = 0;
    for (const v of flat) {
      if (ctx.toBoolean(v)) count++;
    }
    return count % 2 === 1;
  },
  IFERROR(args, ctx) {
    const val = args[0];
    return isError(val) ? (args.length > 1 ? args[1] : '') : val;
  },
  IFNA(args, ctx) {
    return args[0] === ERROR_TYPE.NA ? (args.length > 1 ? args[1] : '') : args[0];
  },
  SWITCH(args, ctx) {
    const expr = args[0];
    for (let i = 1; i < args.length - 1; i += 2) {
      if (expr === args[i] || (typeof expr === 'string' && typeof args[i] === 'string' &&
          expr.toLowerCase() === args[i].toLowerCase())) {
        return args[i + 1];
      }
    }
    return args.length % 2 === 0 ? args[args.length - 1] : ERROR_TYPE.NA;
  },
  CHOOSE(args, ctx) {
    const idx = ctx.toNumber(args[0]);
    if (idx < 1 || idx >= args.length) return ERROR_TYPE.VALUE;
    return args[Math.floor(idx)];
  },
  TRUE() { return true; },
  FALSE() { return false; },

  // ── Text ──
  CONCATENATE(args, ctx) {
    return args.map(a => {
      if (a && a._isRange) {
        return ctx.resolveRange(a).flat().map(v => ctx.toString(v)).join('');
      }
      return ctx.toString(a);
    }).join('');
  },
  CONCAT(args, ctx) { return FUNCTIONS.CONCATENATE(args, ctx); },
  TEXTJOIN(args, ctx) {
    const delim = ctx.toString(args[0]);
    const ignoreEmpty = ctx.toBoolean(args[1]);
    const values = ctx.flattenArgs(args.slice(2));
    const filtered = ignoreEmpty ? values.filter(v => v !== null && v !== undefined && v !== '') : values;
    return filtered.map(v => ctx.toString(v)).join(delim);
  },
  LEFT(args, ctx) {
    const str = ctx.toString(args[0]);
    const n = args.length > 1 ? ctx.toNumber(args[1]) : 1;
    return str.substring(0, n);
  },
  RIGHT(args, ctx) {
    const str = ctx.toString(args[0]);
    const n = args.length > 1 ? ctx.toNumber(args[1]) : 1;
    return str.substring(str.length - n);
  },
  MID(args, ctx) {
    const str = ctx.toString(args[0]);
    const start = ctx.toNumber(args[1]);
    const len = ctx.toNumber(args[2]);
    return str.substring(start - 1, start - 1 + len);
  },
  LEN(args, ctx) {
    return ctx.toString(args[0]).length;
  },
  TRIM(args, ctx) {
    return ctx.toString(args[0]).trim().replace(/\s+/g, ' ');
  },
  CLEAN(args, ctx) {
    return ctx.toString(args[0]).replace(/[\x00-\x1f]/g, '');
  },
  UPPER(args, ctx) { return ctx.toString(args[0]).toUpperCase(); },
  LOWER(args, ctx) { return ctx.toString(args[0]).toLowerCase(); },
  PROPER(args, ctx) {
    return ctx.toString(args[0]).replace(/\w\S*/g, t => t.charAt(0).toUpperCase() + t.substr(1).toLowerCase());
  },
  REPT(args, ctx) {
    return ctx.toString(args[0]).repeat(Math.max(0, ctx.toNumber(args[1])));
  },
  SUBSTITUTE(args, ctx) {
    const text = ctx.toString(args[0]);
    const old = ctx.toString(args[1]);
    const rep = ctx.toString(args[2]);
    if (args.length > 3) {
      const n = ctx.toNumber(args[3]);
      let count = 0, idx = -1;
      while (count < n) {
        idx = text.indexOf(old, idx + 1);
        if (idx === -1) return text;
        count++;
      }
      return text.substring(0, idx) + rep + text.substring(idx + old.length);
    }
    return text.split(old).join(rep);
  },
  REPLACE(args, ctx) {
    const text = ctx.toString(args[0]);
    const start = ctx.toNumber(args[1]) - 1;
    const len = ctx.toNumber(args[2]);
    const rep = ctx.toString(args[3]);
    return text.substring(0, start) + rep + text.substring(start + len);
  },
  FIND(args, ctx) {
    const needle = ctx.toString(args[0]);
    const haystack = ctx.toString(args[1]);
    const start = args.length > 2 ? ctx.toNumber(args[2]) - 1 : 0;
    const idx = haystack.indexOf(needle, start);
    return idx === -1 ? ERROR_TYPE.VALUE : idx + 1;
  },
  SEARCH(args, ctx) {
    const needle = ctx.toString(args[0]).toLowerCase();
    const haystack = ctx.toString(args[1]).toLowerCase();
    const start = args.length > 2 ? ctx.toNumber(args[2]) - 1 : 0;
    const idx = haystack.indexOf(needle, start);
    return idx === -1 ? ERROR_TYPE.VALUE : idx + 1;
  },
  EXACT(args, ctx) {
    return ctx.toString(args[0]) === ctx.toString(args[1]);
  },
  VALUE(args, ctx) {
    const n = ctx.toNumber(args[0]);
    return isNaN(n) ? ERROR_TYPE.VALUE : n;
  },
  TEXT(args, ctx) {
    const val = ctx.toNumber(args[0]);
    const fmt = ctx.toString(args[1]);
    if (isNaN(val)) return ERROR_TYPE.VALUE;
    return formatNumberText(val, fmt);
  },
  FIXED(args, ctx) {
    const n = ctx.toNumber(args[0]);
    const d = args.length > 1 ? ctx.toNumber(args[1]) : 2;
    const noComma = args.length > 2 ? ctx.toBoolean(args[2]) : false;
    if (isNaN(n)) return ERROR_TYPE.VALUE;
    const str = Math.abs(n).toFixed(d);
    if (noComma) return (n < 0 ? '-' : '') + str;
    const parts = str.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return (n < 0 ? '-' : '') + parts.join('.');
  },
  T(args) {
    return typeof args[0] === 'string' ? args[0] : '';
  },
  N(args, ctx) {
    if (typeof args[0] === 'number') return args[0];
    if (typeof args[0] === 'boolean') return args[0] ? 1 : 0;
    return 0;
  },
  CHAR(args, ctx) {
    return String.fromCharCode(ctx.toNumber(args[0]));
  },
  CODE(args, ctx) {
    const str = ctx.toString(args[0]);
    return str.length > 0 ? str.charCodeAt(0) : ERROR_TYPE.VALUE;
  },
  NUMBERVALUE(args, ctx) {
    let text = ctx.toString(args[0]);
    const decSep = args.length > 1 ? ctx.toString(args[1]) : '.';
    const grpSep = args.length > 2 ? ctx.toString(args[2]) : ',';
    text = text.split(grpSep).join('').split(decSep).join('.');
    const n = Number(text);
    return isNaN(n) ? ERROR_TYPE.VALUE : n;
  },

  // ── Lookup ──
  VLOOKUP(args, ctx) {
    const lookup = args[0];
    const table = ctx.resolveRange(args[1]);
    const colIdx = ctx.toNumber(args[2]) - 1;
    const exact = args.length > 3 ? !ctx.toBoolean(args[3]) : false;

    if (colIdx < 0 || !table[0] || colIdx >= table[0].length) return ERROR_TYPE.REF;

    if (exact) {
      for (let r = 0; r < table.length; r++) {
        if (looseEqual(table[r][0], lookup)) return table[r][colIdx];
      }
      return ERROR_TYPE.NA;
    }

    // Approximate match (assumes sorted)
    let best = -1;
    for (let r = 0; r < table.length; r++) {
      const v = table[r][0];
      if (v === null || v === undefined) continue;
      if (typeof v === typeof lookup && v <= lookup) best = r;
    }
    return best >= 0 ? table[best][colIdx] : ERROR_TYPE.NA;
  },
  HLOOKUP(args, ctx) {
    const lookup = args[0];
    const table = ctx.resolveRange(args[1]);
    const rowIdx = ctx.toNumber(args[2]) - 1;
    const exact = args.length > 3 ? !ctx.toBoolean(args[3]) : false;

    if (rowIdx < 0 || rowIdx >= table.length) return ERROR_TYPE.REF;

    const firstRow = table[0] || [];
    if (exact) {
      for (let c = 0; c < firstRow.length; c++) {
        if (looseEqual(firstRow[c], lookup)) return table[rowIdx][c];
      }
      return ERROR_TYPE.NA;
    }

    let best = -1;
    for (let c = 0; c < firstRow.length; c++) {
      if (firstRow[c] !== null && typeof firstRow[c] === typeof lookup && firstRow[c] <= lookup) best = c;
    }
    return best >= 0 ? table[rowIdx][best] : ERROR_TYPE.NA;
  },
  INDEX(args, ctx) {
    const data = ctx.resolveRange(args[0]);
    const rowNum = args.length > 1 ? ctx.toNumber(args[1]) : 0;
    const colNum = args.length > 2 ? ctx.toNumber(args[2]) : 0;

    if (rowNum === 0 && colNum === 0) return data;
    if (rowNum === 0) {
      return data.map(r => r[colNum - 1]);
    }
    if (colNum === 0) {
      return data[rowNum - 1];
    }
    if (rowNum < 1 || rowNum > data.length) return ERROR_TYPE.REF;
    if (colNum < 1 || colNum > (data[0] || []).length) return ERROR_TYPE.REF;
    return data[rowNum - 1][colNum - 1];
  },
  MATCH(args, ctx) {
    const lookup = args[0];
    const rangeData = ctx.resolveRange(args[1]);
    const type = args.length > 2 ? ctx.toNumber(args[2]) : 1;

    // Flatten to 1D
    const arr = rangeData.length === 1 ? rangeData[0] : rangeData.map(r => r[0]);

    if (type === 0) {
      // Exact match (supports wildcards)
      for (let i = 0; i < arr.length; i++) {
        if (looseEqual(arr[i], lookup)) return i + 1;
      }
      // Wildcard match
      if (typeof lookup === 'string' && (lookup.includes('*') || lookup.includes('?'))) {
        const re = new RegExp('^' + lookup.replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
        for (let i = 0; i < arr.length; i++) {
          if (typeof arr[i] === 'string' && re.test(arr[i])) return i + 1;
        }
      }
      return ERROR_TYPE.NA;
    }

    if (type === 1) {
      let best = -1;
      for (let i = 0; i < arr.length; i++) {
        if (arr[i] !== null && arr[i] <= lookup) best = i;
      }
      return best >= 0 ? best + 1 : ERROR_TYPE.NA;
    }

    if (type === -1) {
      let best = -1;
      for (let i = 0; i < arr.length; i++) {
        if (arr[i] !== null && arr[i] >= lookup) best = i;
      }
      return best >= 0 ? best + 1 : ERROR_TYPE.NA;
    }

    return ERROR_TYPE.NA;
  },
  XLOOKUP(args, ctx) {
    const lookup = args[0];
    const lookArr = ctx.resolveRange(args[1]);
    const retArr = ctx.resolveRange(args[2]);
    const notFound = args.length > 3 ? args[3] : ERROR_TYPE.NA;
    const matchMode = args.length > 4 ? ctx.toNumber(args[4]) : 0;

    const lookFlat = lookArr.length === 1 ? lookArr[0] : lookArr.map(r => r[0]);

    for (let i = 0; i < lookFlat.length; i++) {
      if (looseEqual(lookFlat[i], lookup)) {
        const row = retArr[retArr.length === 1 ? 0 : i];
        return retArr.length === 1 ? row[i] : row[0];
      }
    }
    return notFound;
  },
  OFFSET(args, ctx) {
    if (!args[0] || !args[0]._isRange) return ERROR_TYPE.VALUE;
    const { range, sheet } = args[0];
    const rowOff = ctx.toNumber(args[1]);
    const colOff = ctx.toNumber(args[2]);
    const height = args.length > 3 ? ctx.toNumber(args[3]) : range.rowCount;
    const width = args.length > 4 ? ctx.toNumber(args[4]) : range.colCount;
    const newRange = {
      _isRange: true,
      range: {
        startRow: range.startRow + rowOff,
        startCol: range.startCol + colOff,
        endRow: range.startRow + rowOff + height - 1,
        endCol: range.startCol + colOff + width - 1,
      },
      sheet,
    };
    if (height === 1 && width === 1) {
      const vals = ctx.resolveRange(newRange);
      return vals[0][0];
    }
    return newRange;
  },
  INDIRECT(args, ctx) {
    const ref = ctx.toString(args[0]);
    // Very simplified - just parse as cell ref
    const parsed = parseCellRef(ref);
    if (parsed) {
      const sheet = ctx.engine.spreadsheet.getSheetById(ctx.sheetId);
      if (!sheet) return ERROR_TYPE.REF;
      const cell = sheet.getCell(parsed.row, parsed.col);
      return cell ? (cell.computedValue !== null ? cell.computedValue : cell.rawValue) : null;
    }
    return ERROR_TYPE.REF;
  },
  ROW(args, ctx) {
    if (args.length === 0) return ctx.row + 1;
    if (args[0] && args[0]._isRange) return args[0].range.startRow + 1;
    return ctx.row + 1;
  },
  COLUMN(args, ctx) {
    if (args.length === 0) return ctx.col + 1;
    if (args[0] && args[0]._isRange) return args[0].range.startCol + 1;
    return ctx.col + 1;
  },
  ROWS(args, ctx) {
    if (args[0] && args[0]._isRange) return args[0].range.endRow - args[0].range.startRow + 1;
    const data = ctx.resolveRange(args[0]);
    return data.length;
  },
  COLUMNS(args, ctx) {
    if (args[0] && args[0]._isRange) return args[0].range.endCol - args[0].range.startCol + 1;
    const data = ctx.resolveRange(args[0]);
    return data[0] ? data[0].length : 0;
  },
  TRANSPOSE(args, ctx) {
    const data = ctx.resolveRange(args[0]);
    const rows = data.length, cols = data[0] ? data[0].length : 0;
    const result = [];
    for (let c = 0; c < cols; c++) {
      result.push([]);
      for (let r = 0; r < rows; r++) {
        result[c].push(data[r][c]);
      }
    }
    return result;
  },
  UNIQUE(args, ctx) {
    const data = ctx.resolveRange(args[0]);
    const seen = new Set();
    const result = [];
    for (const row of data) {
      const key = JSON.stringify(row);
      if (!seen.has(key)) { seen.add(key); result.push(row); }
    }
    return result;
  },
  SORT(args, ctx) {
    const data = ctx.resolveRange(args[0]).map(r => [...r]);
    const colIdx = args.length > 1 ? ctx.toNumber(args[1]) - 1 : 0;
    const asc = args.length > 2 ? ctx.toBoolean(args[2]) : true;
    data.sort((a, b) => {
      const va = a[colIdx], vb = b[colIdx];
      if (va === vb) return 0;
      if (va === null) return 1;
      if (vb === null) return -1;
      const cmp = va < vb ? -1 : 1;
      return asc ? cmp : -cmp;
    });
    return data;
  },
  FILTER(args, ctx) {
    const data = ctx.resolveRange(args[0]);
    const include = ctx.resolveRange(args[1]);
    const result = [];
    for (let r = 0; r < data.length; r++) {
      if (include[r] && ctx.toBoolean(include[r][0])) result.push(data[r]);
    }
    return result.length ? result : (args.length > 2 ? args[2] : ERROR_TYPE.NA);
  },

  // ── Date / Time ──
  TODAY() {
    const d = new Date();
    return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate());
  },
  NOW() {
    const d = new Date();
    return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()) +
      (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400;
  },
  DATE(args, ctx) {
    return dateToSerial(ctx.toNumber(args[0]), ctx.toNumber(args[1]), ctx.toNumber(args[2]));
  },
  YEAR(args, ctx) { return serialToDate(ctx.toNumber(args[0])).year; },
  MONTH(args, ctx) { return serialToDate(ctx.toNumber(args[0])).month; },
  DAY(args, ctx) { return serialToDate(ctx.toNumber(args[0])).day; },
  HOUR(args, ctx) {
    const frac = ctx.toNumber(args[0]) % 1;
    return Math.floor(frac * 24);
  },
  MINUTE(args, ctx) {
    const frac = ctx.toNumber(args[0]) % 1;
    return Math.floor((frac * 24 * 60) % 60);
  },
  SECOND(args, ctx) {
    const frac = ctx.toNumber(args[0]) % 1;
    return Math.floor((frac * 24 * 3600) % 60);
  },
  DATEVALUE(args, ctx) {
    const d = new Date(ctx.toString(args[0]));
    if (isNaN(d)) return ERROR_TYPE.VALUE;
    return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate());
  },
  WEEKDAY(args, ctx) {
    const serial = ctx.toNumber(args[0]);
    const type = args.length > 1 ? ctx.toNumber(args[1]) : 1;
    const { year, month, day } = serialToDate(serial);
    const d = new Date(year, month - 1, day).getDay();
    if (type === 1) return d + 1;
    if (type === 2) return d === 0 ? 7 : d;
    return d === 0 ? 6 : d - 1;
  },
  EOMONTH(args, ctx) {
    const serial = ctx.toNumber(args[0]);
    const months = ctx.toNumber(args[1]);
    const { year, month } = serialToDate(serial);
    const d = new Date(year, month - 1 + months + 1, 0);
    return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate());
  },
  EDATE(args, ctx) {
    const serial = ctx.toNumber(args[0]);
    const months = ctx.toNumber(args[1]);
    const { year, month, day } = serialToDate(serial);
    const d = new Date(year, month - 1 + months, day);
    return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate());
  },
  DAYS(args, ctx) {
    return ctx.toNumber(args[0]) - ctx.toNumber(args[1]);
  },
  NETWORKDAYS(args, ctx) {
    let start = Math.floor(ctx.toNumber(args[0]));
    const end = Math.floor(ctx.toNumber(args[1]));
    let count = 0;
    const dir = start <= end ? 1 : -1;
    while (start !== end + dir) {
      const { year, month, day } = serialToDate(start);
      const dow = new Date(year, month - 1, day).getDay();
      if (dow !== 0 && dow !== 6) count++;
      start += dir;
    }
    return count;
  },
  DATEDIF(args, ctx) {
    const s = serialToDate(ctx.toNumber(args[0]));
    const e = serialToDate(ctx.toNumber(args[1]));
    const unit = ctx.toString(args[2]).toUpperCase();
    const d1 = new Date(s.year, s.month - 1, s.day);
    const d2 = new Date(e.year, e.month - 1, e.day);
    if (unit === 'D') return Math.floor((d2 - d1) / 86400000);
    if (unit === 'M') return (e.year - s.year) * 12 + (e.month - s.month);
    if (unit === 'Y') return e.year - s.year;
    return ERROR_TYPE.VALUE;
  },

  // ── Info / Type ──
  ISBLANK(args) { return args[0] === null || args[0] === undefined || args[0] === ''; },
  ISERROR(args) { return isError(args[0]); },
  ISNA(args) { return args[0] === ERROR_TYPE.NA; },
  ISNUMBER(args) { return typeof args[0] === 'number'; },
  ISTEXT(args) { return typeof args[0] === 'string' && !isError(args[0]); },
  ISLOGICAL(args) { return typeof args[0] === 'boolean'; },
  ISEVEN(args, ctx) { return Math.floor(ctx.toNumber(args[0])) % 2 === 0; },
  ISODD(args, ctx) { return Math.floor(ctx.toNumber(args[0])) % 2 !== 0; },
  TYPE(args) {
    const v = args[0];
    if (typeof v === 'number') return 1;
    if (typeof v === 'string') return isError(v) ? 16 : 2;
    if (typeof v === 'boolean') return 4;
    return 1;
  },
  ERROR_TYPE(args) {
    const errors = [ERROR_TYPE.NULL, ERROR_TYPE.DIV0, ERROR_TYPE.VALUE, ERROR_TYPE.REF, ERROR_TYPE.NAME, ERROR_TYPE.NUM, ERROR_TYPE.NA];
    const idx = errors.indexOf(args[0]);
    return idx >= 0 ? idx + 1 : ERROR_TYPE.NA;
  },
  NA() { return ERROR_TYPE.NA; },

  // ── Array-ish ──
  ARRAYFORMULA(args) { return args[0]; },

  // ── Sparklines ──
  SPARKLINE(args, ctx) {
    const data = args[0];
    let values;
    if (data && data._isRange) {
      values = ctx.resolveRange(data).flat().filter(v => typeof v === 'number');
    } else if (Array.isArray(data)) {
      values = data.flat().filter(v => typeof v === 'number');
    } else {
      return '';
    }

    // Options from second arg
    let type = 'line', color = '#4285f4';
    if (args.length > 1) {
      const opts = args[1];
      if (opts && opts._isRange) {
        const optData = ctx.resolveRange(opts);
        for (const row of optData) {
          if (row[0] === 'charttype' && row[1]) type = String(row[1]);
          if (row[0] === 'color' && row[1]) color = String(row[1]);
        }
      } else if (typeof opts === 'object' && opts) {
        type = opts.charttype || opts.type || type;
        color = opts.color || color;
      }
    }

    // Return a sparkline marker object — the renderer will detect this
    return { _sparkline: true, values, type, color };
  },

  // ── Financial ──
  PMT(args, ctx) {
    const rate = ctx.toNumber(args[0]);
    const nper = ctx.toNumber(args[1]);
    const pv = ctx.toNumber(args[2]);
    const fv = args.length > 3 ? ctx.toNumber(args[3]) : 0;
    const type = args.length > 4 ? ctx.toNumber(args[4]) : 0;
    if (rate === 0) return -(pv + fv) / nper;
    const pvif = Math.pow(1 + rate, nper);
    return -(rate * (pv * pvif + fv)) / (pvif - 1) / (1 + rate * type);
  },
  FV(args, ctx) {
    const rate = ctx.toNumber(args[0]);
    const nper = ctx.toNumber(args[1]);
    const pmt = ctx.toNumber(args[2]);
    const pv = args.length > 3 ? ctx.toNumber(args[3]) : 0;
    const type = args.length > 4 ? ctx.toNumber(args[4]) : 0;
    if (rate === 0) return -(pv + pmt * nper);
    const pvif = Math.pow(1 + rate, nper);
    return -(pv * pvif + pmt * (1 + rate * type) * (pvif - 1) / rate);
  },
  PV(args, ctx) {
    const rate = ctx.toNumber(args[0]);
    const nper = ctx.toNumber(args[1]);
    const pmt = ctx.toNumber(args[2]);
    const fv = args.length > 3 ? ctx.toNumber(args[3]) : 0;
    const type = args.length > 4 ? ctx.toNumber(args[4]) : 0;
    if (rate === 0) return -(fv + pmt * nper);
    const pvif = Math.pow(1 + rate, nper);
    return -(fv + pmt * (1 + rate * type) * (pvif - 1) / rate) / pvif;
  },
  NPV(args, ctx) {
    const rate = ctx.toNumber(args[0]);
    const values = ctx.flattenNumbers(args.slice(1));
    let npv = 0;
    for (let i = 0; i < values.length; i++) {
      npv += values[i] / Math.pow(1 + rate, i + 1);
    }
    return npv;
  },
  IRR(args, ctx) {
    const values = ctx.flattenNumbers([args[0]]);
    let guess = args.length > 1 ? ctx.toNumber(args[1]) : 0.1;
    for (let iter = 0; iter < 100; iter++) {
      let npv = 0, dnpv = 0;
      for (let i = 0; i < values.length; i++) {
        const f = Math.pow(1 + guess, i);
        npv += values[i] / f;
        dnpv -= i * values[i] / (f * (1 + guess));
      }
      if (Math.abs(npv) < 1e-10) return guess;
      if (dnpv === 0) return ERROR_TYPE.NUM;
      guess -= npv / dnpv;
    }
    return ERROR_TYPE.NUM;
  },
  RATE(args, ctx) {
    const nper = ctx.toNumber(args[0]);
    const pmt = ctx.toNumber(args[1]);
    const pv = ctx.toNumber(args[2]);
    const fv = args.length > 3 ? ctx.toNumber(args[3]) : 0;
    let guess = args.length > 5 ? ctx.toNumber(args[5]) : 0.1;
    for (let i = 0; i < 100; i++) {
      const pvif = Math.pow(1 + guess, nper);
      const y = pv * pvif + pmt * (pvif - 1) / guess + fv;
      const dy = pv * nper * Math.pow(1 + guess, nper - 1) +
        pmt * (nper * Math.pow(1 + guess, nper - 1) * guess - (pvif - 1)) / (guess * guess);
      const newGuess = guess - y / dy;
      if (Math.abs(newGuess - guess) < 1e-10) return newGuess;
      guess = newGuess;
    }
    return ERROR_TYPE.NUM;
  },
  SLN(args, ctx) {
    const cost = ctx.toNumber(args[0]);
    const salvage = ctx.toNumber(args[1]);
    const life = ctx.toNumber(args[2]);
    return (cost - salvage) / life;
  },
};

// ── Helpers ──
function matchCriteria(value, criteria) {
  if (criteria === null || criteria === undefined) return false;

  if (typeof criteria === 'number' || typeof criteria === 'boolean') {
    return value === criteria;
  }

  const str = String(criteria);
  if (str.startsWith('>=')) return value != null && value >= parseFloat(str.slice(2));
  if (str.startsWith('<=')) return value != null && value <= parseFloat(str.slice(2));
  if (str.startsWith('<>')) {
    const cmp = str.slice(2);
    const n = parseFloat(cmp);
    return !isNaN(n) ? value !== n : String(value).toLowerCase() !== cmp.toLowerCase();
  }
  if (str.startsWith('>')) return value != null && value > parseFloat(str.slice(1));
  if (str.startsWith('<')) return value != null && value < parseFloat(str.slice(1));
  if (str.startsWith('=')) {
    const cmp = str.slice(1);
    const n = parseFloat(cmp);
    return !isNaN(n) ? value === n : String(value).toLowerCase() === cmp.toLowerCase();
  }

  // Wildcard match
  if (str.includes('*') || str.includes('?')) {
    const re = new RegExp('^' + str.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
    return re.test(String(value));
  }

  // Direct comparison
  if (!isNaN(Number(str)) && typeof value === 'number') return value === Number(str);
  return String(value).toLowerCase() === str.toLowerCase();
}

function looseEqual(a, b) {
  if (a === b) return true;
  if (a === null || b === null) return false;
  if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
  return a == b;
}

// Excel date serial number helpers (base: 1900-01-01 = 1)
function dateToSerial(year, month, day) {
  const d = new Date(year, month - 1, day);
  const epoch = new Date(1899, 11, 30);
  return Math.floor((d - epoch) / 86400000);
}

function serialToDate(serial) {
  const d = new Date(1899, 11, 30);
  d.setDate(d.getDate() + Math.floor(serial));
  return { year: d.getFullYear(), month: d.getMonth() + 1, day: d.getDate() };
}

function formatNumberText(val, fmt) {
  if (fmt === '0') return Math.round(val).toString();
  if (fmt === '0.00') return val.toFixed(2);
  if (fmt === '#,##0') return Math.round(val).toLocaleString();
  if (fmt === '#,##0.00') return val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  if (fmt === '0%') return Math.round(val * 100) + '%';
  if (fmt === '0.00%') return (val * 100).toFixed(2) + '%';
  if (fmt.includes('$')) return '$' + val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  // Date formats
  if (fmt.includes('yyyy') || fmt.includes('MM') || fmt.includes('dd')) {
    const { year, month, day } = serialToDate(val);
    return fmt
      .replace('yyyy', year)
      .replace('yy', String(year).slice(-2))
      .replace('MM', String(month).padStart(2, '0'))
      .replace('M', month)
      .replace('dd', String(day).padStart(2, '0'))
      .replace('d', day);
  }

  // Default
  return val.toString();
}
