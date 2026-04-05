import CellRange from '../core/CellRange.js';
import { ERROR_TYPE, CELL_TYPE } from '../utils/constants.js';
import { parseCellRef, cellRefToString, colToIndex, indexToCol, cellKey, keyToRC } from '../utils/helpers.js';
import { FUNCTIONS } from './functions.js';

// ── Token Types ──
const T = {
  NUMBER: 'NUMBER',
  STRING: 'STRING',
  BOOLEAN: 'BOOLEAN',
  ERROR: 'ERROR',
  CELL_REF: 'CELL_REF',
  NAME: 'NAME',
  OP: 'OP',
  LPAREN: 'LPAREN',
  RPAREN: 'RPAREN',
  COMMA: 'COMMA',
  COLON: 'COLON',
  BANG: 'BANG',
  SEMICOLON: 'SEMICOLON',
  EOF: 'EOF',
};

// ── Tokenizer ──
function tokenize(formula) {
  const tokens = [];
  let i = 0;
  const src = formula;
  const len = src.length;

  while (i < len) {
    const ch = src[i];

    // Whitespace
    if (ch === ' ' || ch === '\t') { i++; continue; }

    // String literal
    if (ch === '"') {
      let str = '';
      i++;
      while (i < len) {
        if (src[i] === '"') {
          if (i + 1 < len && src[i + 1] === '"') {
            str += '"';
            i += 2;
          } else {
            i++;
            break;
          }
        } else {
          str += src[i++];
        }
      }
      tokens.push({ type: T.STRING, value: str });
      continue;
    }

    // Number
    if ((ch >= '0' && ch <= '9') || (ch === '.' && i + 1 < len && src[i + 1] >= '0' && src[i + 1] <= '9')) {
      let num = '';
      while (i < len && ((src[i] >= '0' && src[i] <= '9') || src[i] === '.')) {
        num += src[i++];
      }
      // Scientific notation
      if (i < len && (src[i] === 'e' || src[i] === 'E')) {
        num += src[i++];
        if (i < len && (src[i] === '+' || src[i] === '-')) num += src[i++];
        while (i < len && src[i] >= '0' && src[i] <= '9') num += src[i++];
      }
      // Percent
      if (i < len && src[i] === '%') {
        i++;
        tokens.push({ type: T.NUMBER, value: parseFloat(num) / 100 });
      } else {
        tokens.push({ type: T.NUMBER, value: parseFloat(num) });
      }
      continue;
    }

    // Operators
    if (ch === '+' || ch === '-' || ch === '*' || ch === '/' || ch === '^' || ch === '&') {
      tokens.push({ type: T.OP, value: ch });
      i++;
      continue;
    }

    if (ch === '=' || ch === '<' || ch === '>') {
      if (src[i + 1] === '=' && ch !== '=') {
        tokens.push({ type: T.OP, value: ch + '=' });
        i += 2;
      } else if (ch === '<' && src[i + 1] === '>') {
        tokens.push({ type: T.OP, value: '<>' });
        i += 2;
      } else {
        tokens.push({ type: T.OP, value: ch });
        i++;
      }
      continue;
    }

    if (ch === '%') {
      tokens.push({ type: T.OP, value: '%' });
      i++;
      continue;
    }

    if (ch === '(') { tokens.push({ type: T.LPAREN }); i++; continue; }
    if (ch === ')') { tokens.push({ type: T.RPAREN }); i++; continue; }
    if (ch === ',') { tokens.push({ type: T.COMMA }); i++; continue; }
    if (ch === ';') { tokens.push({ type: T.SEMICOLON }); i++; continue; }
    if (ch === ':') { tokens.push({ type: T.COLON }); i++; continue; }
    if (ch === '!') { tokens.push({ type: T.BANG }); i++; continue; }

    // Quoted sheet name
    if (ch === "'") {
      let name = '';
      i++;
      while (i < len && src[i] !== "'") {
        name += src[i++];
      }
      i++; // skip closing quote
      tokens.push({ type: T.NAME, value: name, quoted: true });
      continue;
    }

    // Error values
    if (ch === '#') {
      let err = '#';
      i++;
      while (i < len && src[i] !== ' ' && src[i] !== ')' && src[i] !== ',' && src[i] !== ';') {
        err += src[i++];
      }
      tokens.push({ type: T.ERROR, value: err });
      continue;
    }

    // Identifiers (cell refs, function names, TRUE/FALSE)
    if ((ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || ch === '_' || ch === '$') {
      let ident = '';
      while (i < len && (/[A-Za-z0-9_$]/.test(src[i]))) {
        ident += src[i++];
      }
      const upper = ident.toUpperCase();
      if (upper === 'TRUE') {
        tokens.push({ type: T.BOOLEAN, value: true });
      } else if (upper === 'FALSE') {
        tokens.push({ type: T.BOOLEAN, value: false });
      } else {
        // Check if it's a cell reference pattern: $?[A-Z]+$?\d+
        const cellMatch = ident.match(/^(\$?[A-Za-z]+)(\$?\d+)$/);
        if (cellMatch) {
          const colPart = cellMatch[1].replace('$', '');
          const isAllAlpha = /^[A-Za-z]+$/.test(colPart);
          const ci = isAllAlpha ? colToIndex(colPart.toUpperCase()) : -1;
          if (isAllAlpha && ci >= 0 && ci < 702) {
            tokens.push({ type: T.CELL_REF, value: ident });
          } else {
            tokens.push({ type: T.NAME, value: ident });
          }
        } else {
          tokens.push({ type: T.NAME, value: ident });
        }
      }
      continue;
    }

    // Unknown character, skip
    i++;
  }

  tokens.push({ type: T.EOF });
  return tokens;
}

// ── AST Node Types ──
const N = {
  NUMBER: 'num',
  STRING: 'str',
  BOOLEAN: 'bool',
  ERROR: 'err',
  CELL_REF: 'cell',
  RANGE_REF: 'range',
  BINARY: 'bin',
  UNARY: 'un',
  FUNC: 'fn',
  ARRAY: 'arr',
};

// ── Parser ──
class Parser {
  constructor(tokens) {
    this.tokens = tokens;
    this.pos = 0;
  }

  peek() { return this.tokens[this.pos]; }
  advance() { return this.tokens[this.pos++]; }

  expect(type) {
    const tok = this.advance();
    if (tok.type !== type) throw new Error(`Expected ${type}, got ${tok.type}`);
    return tok;
  }

  match(type, value) {
    const tok = this.peek();
    if (tok.type === type && (value === undefined || tok.value === value)) {
      return this.advance();
    }
    return null;
  }

  parse() {
    const node = this.expression();
    this.expect(T.EOF);
    return node;
  }

  expression() {
    return this.comparison();
  }

  comparison() {
    let left = this.concatenation();
    while (true) {
      const tok = this.peek();
      if (tok.type === T.OP && (tok.value === '=' || tok.value === '<>' || tok.value === '<' ||
          tok.value === '>' || tok.value === '<=' || tok.value === '>=')) {
        this.advance();
        left = { type: N.BINARY, op: tok.value, left, right: this.concatenation() };
      } else break;
    }
    return left;
  }

  concatenation() {
    let left = this.addition();
    while (this.match(T.OP, '&')) {
      left = { type: N.BINARY, op: '&', left, right: this.addition() };
    }
    return left;
  }

  addition() {
    let left = this.multiplication();
    while (true) {
      const tok = this.peek();
      if (tok.type === T.OP && (tok.value === '+' || tok.value === '-')) {
        this.advance();
        left = { type: N.BINARY, op: tok.value, left, right: this.multiplication() };
      } else break;
    }
    return left;
  }

  multiplication() {
    let left = this.power();
    while (true) {
      const tok = this.peek();
      if (tok.type === T.OP && (tok.value === '*' || tok.value === '/')) {
        this.advance();
        left = { type: N.BINARY, op: tok.value, left, right: this.power() };
      } else break;
    }
    return left;
  }

  power() {
    let left = this.unary();
    if (this.match(T.OP, '^')) {
      left = { type: N.BINARY, op: '^', left, right: this.power() }; // right-assoc
    }
    return left;
  }

  unary() {
    const tok = this.peek();
    if (tok.type === T.OP && (tok.value === '+' || tok.value === '-')) {
      this.advance();
      return { type: N.UNARY, op: tok.value, operand: this.unary() };
    }
    return this.postfix();
  }

  postfix() {
    let node = this.primary();
    if (this.match(T.OP, '%')) {
      node = { type: N.BINARY, op: '*', left: node, right: { type: N.NUMBER, value: 0.01 } };
    }
    return node;
  }

  primary() {
    const tok = this.peek();

    if (tok.type === T.NUMBER) {
      this.advance();
      return { type: N.NUMBER, value: tok.value };
    }

    if (tok.type === T.STRING) {
      this.advance();
      return { type: N.STRING, value: tok.value };
    }

    if (tok.type === T.BOOLEAN) {
      this.advance();
      return { type: N.BOOLEAN, value: tok.value };
    }

    if (tok.type === T.ERROR) {
      this.advance();
      return { type: N.ERROR, value: tok.value };
    }

    if (tok.type === T.LPAREN) {
      this.advance();
      const expr = this.expression();
      this.expect(T.RPAREN);
      return expr;
    }

    // Cell ref, range, or function call - also handle sheet!ref
    if (tok.type === T.CELL_REF || tok.type === T.NAME) {
      this.advance();
      let sheetName = null;

      // Sheet reference: Name!Ref or 'Sheet Name'!Ref
      if (this.match(T.BANG)) {
        sheetName = tok.value;
        const refTok = this.advance();
        if (refTok.type === T.CELL_REF) {
          const ref = parseCellRef(refTok.value);
          // Check for range A1:B2
          if (this.match(T.COLON)) {
            const endTok = this.expect(T.CELL_REF);
            const endRef = parseCellRef(endTok.value);
            return { type: N.RANGE_REF, sheet: sheetName, start: ref, end: endRef };
          }
          return { type: N.CELL_REF, sheet: sheetName, ref };
        }
        return { type: N.ERROR, value: ERROR_TYPE.REF };
      }

      // Function call: NAME(args)
      if (tok.type === T.NAME && this.peek().type === T.LPAREN) {
        this.advance(); // skip (
        const args = [];
        if (this.peek().type !== T.RPAREN) {
          args.push(this.expression());
          while (this.match(T.COMMA) || this.match(T.SEMICOLON)) {
            args.push(this.expression());
          }
        }
        this.expect(T.RPAREN);
        return { type: N.FUNC, name: tok.value.toUpperCase(), args };
      }

      // Cell reference
      if (tok.type === T.CELL_REF) {
        const ref = parseCellRef(tok.value);
        // Check for range
        if (this.match(T.COLON)) {
          const endTok = this.expect(T.CELL_REF);
          const endRef = parseCellRef(endTok.value);
          return { type: N.RANGE_REF, sheet: null, start: ref, end: endRef };
        }
        return { type: N.CELL_REF, sheet: null, ref };
      }

      // Named range or unknown name
      return { type: N.ERROR, value: ERROR_TYPE.NAME };
    }

    throw new Error(`Unexpected token: ${tok.type}`);
  }
}

// ── Dependency Graph ──
export class DependencyGraph {
  constructor() {
    this.deps = new Map();      // cellKey -> Set<cellKey>   (this cell depends on these)
    this.rdeps = new Map();     // cellKey -> Set<cellKey>   (these cells depend on this)
  }

  setDependencies(key, depKeys) {
    // Remove old reverse deps
    const oldDeps = this.deps.get(key);
    if (oldDeps) {
      for (const dk of oldDeps) {
        const rd = this.rdeps.get(dk);
        if (rd) rd.delete(key);
      }
    }

    if (depKeys.size === 0) {
      this.deps.delete(key);
      return;
    }

    this.deps.set(key, depKeys);

    // Add new reverse deps
    for (const dk of depKeys) {
      if (!this.rdeps.has(dk)) this.rdeps.set(dk, new Set());
      this.rdeps.get(dk).add(key);
    }
  }

  getDependents(key) {
    return this.rdeps.get(key) || new Set();
  }

  getEvalOrder(changedKeys) {
    const visited = new Set();
    const order = [];
    const inStack = new Set();

    const visit = (key) => {
      if (visited.has(key)) return;
      if (inStack.has(key)) return; // circular ref
      inStack.add(key);
      const dependents = this.rdeps.get(key);
      if (dependents) {
        for (const dk of dependents) {
          visit(dk);
        }
      }
      inStack.delete(key);
      visited.add(key);
      order.push(key);
    };

    for (const key of changedKeys) {
      visit(key);
    }

    return order.reverse();
  }

  removeDependencies(key) {
    this.setDependencies(key, new Set());
    this.rdeps.delete(key);
  }

  clear() {
    this.deps.clear();
    this.rdeps.clear();
  }
}

// ── Formula Evaluator ──
export default class FormulaEngine {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.depGraph = new DependencyGraph();
    this._evaluating = new Set();
  }

  evaluate(formula, sheetId, row, col) {
    const key = sheetId + ':' + cellKey(row, col);
    if (this._evaluating.has(key)) {
      return ERROR_TYPE.CIRCULAR;
    }

    try {
      const tokens = tokenize(formula.startsWith('=') ? formula.substring(1) : formula);
      const parser = new Parser(tokens);
      const ast = parser.parse();

      // Collect dependencies
      const deps = new Set();
      this._collectDeps(ast, sheetId, deps);
      this.depGraph.setDependencies(cellKey(row, col), deps);

      // Evaluate
      this._evaluating.add(key);
      const result = this._eval(ast, sheetId, row, col);
      this._evaluating.delete(key);

      return result;
    } catch (e) {
      this._evaluating.delete(sheetId + ':' + cellKey(row, col));
      return ERROR_TYPE.VALUE;
    }
  }

  _collectDeps(node, sheetId, deps) {
    switch (node.type) {
      case N.CELL_REF: {
        const ref = node.ref;
        deps.add(cellKey(ref.row, ref.col));
        break;
      }
      case N.RANGE_REF: {
        const { start, end } = node;
        const r1 = Math.min(start.row, end.row), r2 = Math.max(start.row, end.row);
        const c1 = Math.min(start.col, end.col), c2 = Math.max(start.col, end.col);
        for (let r = r1; r <= r2; r++) {
          for (let c = c1; c <= c2; c++) {
            deps.add(cellKey(r, c));
          }
        }
        break;
      }
      case N.BINARY:
        this._collectDeps(node.left, sheetId, deps);
        this._collectDeps(node.right, sheetId, deps);
        break;
      case N.UNARY:
        this._collectDeps(node.operand, sheetId, deps);
        break;
      case N.FUNC:
        for (const arg of node.args) {
          this._collectDeps(arg, sheetId, deps);
        }
        break;
    }
  }

  _eval(node, sheetId, row, col) {
    switch (node.type) {
      case N.NUMBER: return node.value;
      case N.STRING: return node.value;
      case N.BOOLEAN: return node.value;
      case N.ERROR: return node.value;

      case N.CELL_REF: {
        const sheet = node.sheet
          ? this.spreadsheet.getSheetByName(node.sheet)
          : this.spreadsheet.getSheetById(sheetId);
        if (!sheet) return ERROR_TYPE.REF;
        const cell = sheet.getCell(node.ref.row, node.ref.col);
        if (!cell) return null;
        if (cell.formula) {
          return cell.computedValue;
        }
        return cell.computedValue !== null ? cell.computedValue : cell.rawValue;
      }

      case N.RANGE_REF: {
        const sheet = node.sheet
          ? this.spreadsheet.getSheetByName(node.sheet)
          : this.spreadsheet.getSheetById(sheetId);
        if (!sheet) return ERROR_TYPE.REF;
        const r1 = Math.min(node.start.row, node.end.row);
        const c1 = Math.min(node.start.col, node.end.col);
        const r2 = Math.max(node.start.row, node.end.row);
        const c2 = Math.max(node.start.col, node.end.col);
        const range = new CellRange(r1, c1, r2, c2);
        return { _isRange: true, range, sheet };
      }

      case N.UNARY: {
        const val = this._eval(node.operand, sheetId, row, col);
        if (typeof val === 'string' && val.startsWith('#')) return val;
        const num = this._toNumber(val);
        if (isNaN(num)) return ERROR_TYPE.VALUE;
        return node.op === '-' ? -num : num;
      }

      case N.BINARY: {
        const left = this._eval(node.left, sheetId, row, col);
        const right = this._eval(node.right, sheetId, row, col);

        if (typeof left === 'string' && left.startsWith('#')) return left;
        if (typeof right === 'string' && right.startsWith('#')) return right;

        if (node.op === '&') {
          return this._toString(left) + this._toString(right);
        }

        if ('= <> < > <= >='.includes(node.op)) {
          return this._compare(left, right, node.op);
        }

        const l = this._toNumber(left);
        const r = this._toNumber(right);
        if (isNaN(l) || isNaN(r)) return ERROR_TYPE.VALUE;

        switch (node.op) {
          case '+': return l + r;
          case '-': return l - r;
          case '*': return l * r;
          case '/': return r === 0 ? ERROR_TYPE.DIV0 : l / r;
          case '^': return Math.pow(l, r);
        }
        return ERROR_TYPE.VALUE;
      }

      case N.FUNC: {
        const fn = FUNCTIONS[node.name];
        if (!fn) return ERROR_TYPE.NAME;

        const context = {
          engine: this,
          sheetId,
          row,
          col,
          evaluate: (n) => this._eval(n, sheetId, row, col),
          toNumber: (v) => this._toNumber(v),
          toString: (v) => this._toString(v),
          toBoolean: (v) => this._toBoolean(v),
          resolveRange: (v) => this._resolveRange(v),
          flattenArgs: (args) => this._flattenArgs(args, sheetId, row, col),
          flattenNumbers: (args) => this._flattenNumbers(args, sheetId, row, col),
        };

        try {
          const evaluatedArgs = node.args.map(a => this._eval(a, sheetId, row, col));
          return fn(evaluatedArgs, context, node.args);
        } catch (e) {
          return ERROR_TYPE.VALUE;
        }
      }
    }
    return ERROR_TYPE.VALUE;
  }

  _toNumber(val) {
    if (val === null || val === undefined || val === '') return 0;
    if (typeof val === 'number') return val;
    if (typeof val === 'boolean') return val ? 1 : 0;
    if (typeof val === 'string') {
      const n = Number(val);
      return isNaN(n) ? NaN : n;
    }
    return NaN;
  }

  _toString(val) {
    if (val === null || val === undefined) return '';
    if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE';
    return String(val);
  }

  _toBoolean(val) {
    if (typeof val === 'boolean') return val;
    if (typeof val === 'number') return val !== 0;
    if (typeof val === 'string') {
      if (val.toUpperCase() === 'TRUE') return true;
      if (val.toUpperCase() === 'FALSE') return false;
    }
    return false;
  }

  _compare(left, right, op) {
    let l = left === null ? '' : left;
    let r = right === null ? '' : right;

    if (typeof l === 'string' && typeof r === 'string') {
      l = l.toLowerCase();
      r = r.toLowerCase();
    }

    switch (op) {
      case '=': return l === r;
      case '<>': return l !== r;
      case '<': return l < r;
      case '>': return l > r;
      case '<=': return l <= r;
      case '>=': return l >= r;
    }
    return false;
  }

  _resolveRange(val) {
    if (val && val._isRange) {
      const { range, sheet } = val;
      const values = [];
      for (let r = range.startRow; r <= range.endRow; r++) {
        const row = [];
        for (let c = range.startCol; c <= range.endCol; c++) {
          const cell = sheet.getCell(r, c);
          row.push(cell ? (cell.computedValue !== null ? cell.computedValue : cell.rawValue) : null);
        }
        values.push(row);
      }
      return values;
    }
    return [[val]];
  }

  _flattenArgs(args, sheetId, row, col) {
    const result = [];
    for (const arg of args) {
      if (arg && arg._isRange) {
        const values = this._resolveRange(arg);
        for (const row of values) {
          for (const v of row) {
            result.push(v);
          }
        }
      } else {
        result.push(arg);
      }
    }
    return result;
  }

  _flattenNumbers(args, sheetId, row, col) {
    const flat = this._flattenArgs(args, sheetId, row, col);
    const result = [];
    for (const v of flat) {
      if (v === null || v === undefined || v === '') continue;
      if (typeof v === 'boolean') continue;
      const n = this._toNumber(v);
      if (!isNaN(n)) result.push(n);
    }
    return result;
  }

  recalculate(sheet) {
    for (const [key, cell] of sheet.cells) {
      if (cell.formula) {
        const row = key >> 16;
        const col = key & 0xffff;
        cell.computedValue = this.evaluate(cell.formula, sheet.id, row, col);
      }
    }
  }

  recalcChanged(sheet, changedKeys) {
    const order = this.depGraph.getEvalOrder(changedKeys);
    for (const key of order) {
      const row = key >> 16;
      const col = key & 0xffff;
      const cell = sheet.getCell(row, col);
      if (cell && cell.formula) {
        cell.computedValue = this.evaluate(cell.formula, sheet.id, row, col);
      }
    }
  }

  /**
   * Adjust all cell references in a formula when rows/cols are inserted or deleted.
   *
   * @param {string} formula - The formula string (e.g., "=SUM(A1:A10)")
   * @param {'row'|'col'} dimension - Whether rows or columns were changed
   * @param {number} at - 0-based index where the insert/delete happened
   * @param {number} count - Positive = inserted, negative = deleted
   * @returns {string} Adjusted formula
   *
   * Rules (matching Google Sheets):
   * - References at or after the insert/delete point shift by `count`
   * - Absolute references ($A$1) ALSO shift (unlike copy-paste)
   * - If a deletion removes a referenced cell, it becomes #REF!
   * - Ranges that span the insert point expand/contract
   */
  static adjustFormula(formula, dimension, at, count) {
    if (!formula || !formula.startsWith('=')) return formula;

    // Match cell refs and ranges, including $-prefixed
    // We need to handle ranges (A1:B5) as a unit to properly expand/contract
    const rangeRegex = /(\$?)([A-Z]+)(\$?)(\d+)(?::(\$?)([A-Z]+)(\$?)(\d+))?/gi;

    return formula.replace(rangeRegex, (match, ca1, cs1, ra1, rs1, ca2, cs2, ra2, rs2) => {
      let c1 = colToIndex(cs1.toUpperCase());
      let r1 = parseInt(rs1, 10) - 1;
      const isRange = cs2 !== undefined;
      let c2 = isRange ? colToIndex(cs2.toUpperCase()) : c1;
      let r2 = isRange ? parseInt(rs2, 10) - 1 : r1;

      if (dimension === 'row') {
        // Adjust rows
        if (count > 0) {
          // INSERT: shift refs at or after `at`
          if (r1 >= at) r1 += count;
          if (r2 >= at) r2 += count;
        } else {
          // DELETE: count is negative, |count| rows removed starting at `at`
          const delEnd = at + Math.abs(count) - 1;
          // Check if ref falls in deleted range
          if (!isRange) {
            if (r1 >= at && r1 <= delEnd) return ERROR_TYPE.REF;
            if (r1 > delEnd) r1 += count;
          } else {
            // Range: contract if deletion is inside, shift if after
            if (r1 >= at && r1 <= delEnd) r1 = at; // clamp start to deletion point
            else if (r1 > delEnd) r1 += count;
            if (r2 >= at && r2 <= delEnd) r2 = Math.max(at - 1, r1); // clamp end
            else if (r2 > delEnd) r2 += count;
            if (r1 > r2) return ERROR_TYPE.REF; // entire range deleted
          }
        }
      } else {
        // Adjust columns
        if (count > 0) {
          if (c1 >= at) c1 += count;
          if (c2 >= at) c2 += count;
        } else {
          const delEnd = at + Math.abs(count) - 1;
          if (!isRange) {
            if (c1 >= at && c1 <= delEnd) return ERROR_TYPE.REF;
            if (c1 > delEnd) c1 += count;
          } else {
            if (c1 >= at && c1 <= delEnd) c1 = at;
            else if (c1 > delEnd) c1 += count;
            if (c2 >= at && c2 <= delEnd) c2 = Math.max(at - 1, c1);
            else if (c2 > delEnd) c2 += count;
            if (c1 > c2) return ERROR_TYPE.REF;
          }
        }
      }

      if (r1 < 0 || c1 < 0) return ERROR_TYPE.REF;

      const ref1 = (ca1 || '') + indexToCol(c1) + (ra1 || '') + (r1 + 1);
      if (!isRange) return ref1;
      const ref2 = (ca2 || '') + indexToCol(c2) + (ra2 || '') + (r2 + 1);
      return ref1 + ':' + ref2;
    });
  }

  // Copy formula adjusting relative references
  static copyFormula(formula, dRow, dCol) {
    if (!formula || !formula.startsWith('=')) return formula;

    return formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/gi, (match, colAbs, colStr, rowAbs, rowStr) => {
      let col = colToIndex(colStr.toUpperCase());
      let row = parseInt(rowStr, 10) - 1;

      if (rowAbs !== '$') row += dRow;
      if (colAbs !== '$') col += dCol;

      if (row < 0 || col < 0) return ERROR_TYPE.REF;
      return (colAbs || '') + indexToCol(col) + (rowAbs || '') + (row + 1);
    });
  }
}
