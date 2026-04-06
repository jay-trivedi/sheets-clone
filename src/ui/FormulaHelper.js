import { el, indexToCol, cellRefToString } from '../utils/helpers.js';

// ── Function Catalog ──
const FUNC_CATALOG = [
  { name: 'SUM', sig: 'SUM(value1, [value2, ...])', desc: 'Returns the sum of a series of numbers' },
  { name: 'AVERAGE', sig: 'AVERAGE(value1, [value2, ...])', desc: 'Returns the average of a series of numbers' },
  { name: 'COUNT', sig: 'COUNT(value1, [value2, ...])', desc: 'Counts the number of numeric values' },
  { name: 'COUNTA', sig: 'COUNTA(value1, [value2, ...])', desc: 'Counts the number of non-empty values' },
  { name: 'COUNTBLANK', sig: 'COUNTBLANK(range)', desc: 'Counts the number of empty cells' },
  { name: 'COUNTIF', sig: 'COUNTIF(range, criteria)', desc: 'Counts cells matching a condition' },
  { name: 'COUNTIFS', sig: 'COUNTIFS(range1, criteria1, [range2, criteria2, ...])', desc: 'Counts cells matching multiple conditions' },
  { name: 'MAX', sig: 'MAX(value1, [value2, ...])', desc: 'Returns the maximum value' },
  { name: 'MIN', sig: 'MIN(value1, [value2, ...])', desc: 'Returns the minimum value' },
  { name: 'IF', sig: 'IF(condition, value_if_true, value_if_false)', desc: 'Returns one value if true, another if false' },
  { name: 'IFS', sig: 'IFS(condition1, value1, [condition2, value2, ...])', desc: 'Checks multiple conditions in sequence' },
  { name: 'IFERROR', sig: 'IFERROR(value, value_if_error)', desc: 'Returns value_if_error if value is an error' },
  { name: 'IFNA', sig: 'IFNA(value, value_if_na)', desc: 'Returns value_if_na if value is #N/A' },
  { name: 'AND', sig: 'AND(logical1, [logical2, ...])', desc: 'Returns TRUE if all arguments are true' },
  { name: 'OR', sig: 'OR(logical1, [logical2, ...])', desc: 'Returns TRUE if any argument is true' },
  { name: 'NOT', sig: 'NOT(logical)', desc: 'Returns the opposite of a logical value' },
  { name: 'XOR', sig: 'XOR(logical1, [logical2, ...])', desc: 'Returns TRUE if an odd number of arguments are true' },
  { name: 'VLOOKUP', sig: 'VLOOKUP(search_key, range, index, [is_sorted])', desc: 'Vertical lookup - searches first column of a range' },
  { name: 'HLOOKUP', sig: 'HLOOKUP(search_key, range, index, [is_sorted])', desc: 'Horizontal lookup - searches first row of a range' },
  { name: 'XLOOKUP', sig: 'XLOOKUP(search_key, lookup_range, result_range, [not_found], [match_mode])', desc: 'Searches a range and returns a matching item' },
  { name: 'INDEX', sig: 'INDEX(reference, row, [column])', desc: 'Returns the value at a given position in a range' },
  { name: 'MATCH', sig: 'MATCH(search_key, range, [type])', desc: 'Returns the position of a value in a range' },
  { name: 'OFFSET', sig: 'OFFSET(reference, rows, cols, [height], [width])', desc: 'Returns a range offset from a starting cell' },
  { name: 'INDIRECT', sig: 'INDIRECT(cell_reference)', desc: 'Returns the value of a cell specified by a string' },
  { name: 'ROW', sig: 'ROW([reference])', desc: 'Returns the row number of a cell' },
  { name: 'COLUMN', sig: 'COLUMN([reference])', desc: 'Returns the column number of a cell' },
  { name: 'ROWS', sig: 'ROWS(range)', desc: 'Returns the number of rows in a range' },
  { name: 'COLUMNS', sig: 'COLUMNS(range)', desc: 'Returns the number of columns in a range' },
  { name: 'CONCATENATE', sig: 'CONCATENATE(string1, [string2, ...])', desc: 'Joins strings together' },
  { name: 'CONCAT', sig: 'CONCAT(string1, [string2, ...])', desc: 'Joins strings together' },
  { name: 'TEXTJOIN', sig: 'TEXTJOIN(delimiter, ignore_empty, text1, [text2, ...])', desc: 'Joins text with a delimiter' },
  { name: 'LEFT', sig: 'LEFT(text, [num_chars])', desc: 'Returns characters from the start of a string' },
  { name: 'RIGHT', sig: 'RIGHT(text, [num_chars])', desc: 'Returns characters from the end of a string' },
  { name: 'MID', sig: 'MID(text, start, length)', desc: 'Returns characters from the middle of a string' },
  { name: 'LEN', sig: 'LEN(text)', desc: 'Returns the length of a string' },
  { name: 'TRIM', sig: 'TRIM(text)', desc: 'Removes extra spaces from text' },
  { name: 'UPPER', sig: 'UPPER(text)', desc: 'Converts text to uppercase' },
  { name: 'LOWER', sig: 'LOWER(text)', desc: 'Converts text to lowercase' },
  { name: 'PROPER', sig: 'PROPER(text)', desc: 'Capitalizes the first letter of each word' },
  { name: 'SUBSTITUTE', sig: 'SUBSTITUTE(text, old_text, new_text, [instance])', desc: 'Replaces occurrences of text' },
  { name: 'REPLACE', sig: 'REPLACE(text, position, length, new_text)', desc: 'Replaces part of a text string' },
  { name: 'FIND', sig: 'FIND(search_for, text_to_search, [start])', desc: 'Finds the position of a string (case-sensitive)' },
  { name: 'SEARCH', sig: 'SEARCH(search_for, text_to_search, [start])', desc: 'Finds the position of a string (case-insensitive)' },
  { name: 'TEXT', sig: 'TEXT(number, format)', desc: 'Formats a number as text' },
  { name: 'VALUE', sig: 'VALUE(text)', desc: 'Converts text to a number' },
  { name: 'EXACT', sig: 'EXACT(text1, text2)', desc: 'Tests whether two strings are identical' },
  { name: 'REPT', sig: 'REPT(text, times)', desc: 'Repeats text a given number of times' },
  { name: 'CHAR', sig: 'CHAR(number)', desc: 'Returns the character for a character code' },
  { name: 'CODE', sig: 'CODE(text)', desc: 'Returns the character code for the first character' },
  { name: 'CLEAN', sig: 'CLEAN(text)', desc: 'Removes non-printable characters from text' },
  { name: 'ABS', sig: 'ABS(number)', desc: 'Returns the absolute value' },
  { name: 'ROUND', sig: 'ROUND(number, [digits])', desc: 'Rounds a number to a given number of digits' },
  { name: 'ROUNDUP', sig: 'ROUNDUP(number, [digits])', desc: 'Rounds a number up' },
  { name: 'ROUNDDOWN', sig: 'ROUNDDOWN(number, [digits])', desc: 'Rounds a number down' },
  { name: 'CEILING', sig: 'CEILING(number, [significance])', desc: 'Rounds up to the nearest multiple' },
  { name: 'FLOOR', sig: 'FLOOR(number, [significance])', desc: 'Rounds down to the nearest multiple' },
  { name: 'INT', sig: 'INT(number)', desc: 'Rounds down to the nearest integer' },
  { name: 'MOD', sig: 'MOD(dividend, divisor)', desc: 'Returns the remainder after division' },
  { name: 'POWER', sig: 'POWER(base, exponent)', desc: 'Returns a number raised to a power' },
  { name: 'SQRT', sig: 'SQRT(number)', desc: 'Returns the square root' },
  { name: 'LOG', sig: 'LOG(number, [base])', desc: 'Returns the logarithm of a number' },
  { name: 'LOG10', sig: 'LOG10(number)', desc: 'Returns the base-10 logarithm' },
  { name: 'LN', sig: 'LN(number)', desc: 'Returns the natural logarithm' },
  { name: 'EXP', sig: 'EXP(number)', desc: 'Returns e raised to a power' },
  { name: 'PI', sig: 'PI()', desc: 'Returns the value of Pi' },
  { name: 'RAND', sig: 'RAND()', desc: 'Returns a random number between 0 and 1' },
  { name: 'RANDBETWEEN', sig: 'RANDBETWEEN(low, high)', desc: 'Returns a random integer in a range' },
  { name: 'SIGN', sig: 'SIGN(number)', desc: 'Returns the sign of a number (-1, 0, or 1)' },
  { name: 'PRODUCT', sig: 'PRODUCT(value1, [value2, ...])', desc: 'Returns the product of a series of numbers' },
  { name: 'SUMIF', sig: 'SUMIF(range, criteria, [sum_range])', desc: 'Sums cells matching a condition' },
  { name: 'SUMIFS', sig: 'SUMIFS(sum_range, range1, criteria1, [range2, criteria2, ...])', desc: 'Sums cells matching multiple conditions' },
  { name: 'SUMPRODUCT', sig: 'SUMPRODUCT(array1, [array2, ...])', desc: 'Returns the sum of products of corresponding arrays' },
  { name: 'AVERAGEIF', sig: 'AVERAGEIF(range, criteria, [average_range])', desc: 'Averages cells matching a condition' },
  { name: 'MEDIAN', sig: 'MEDIAN(value1, [value2, ...])', desc: 'Returns the median value' },
  { name: 'MODE', sig: 'MODE(value1, [value2, ...])', desc: 'Returns the most common value' },
  { name: 'STDEV', sig: 'STDEV(value1, [value2, ...])', desc: 'Returns the standard deviation (sample)' },
  { name: 'STDEVP', sig: 'STDEVP(value1, [value2, ...])', desc: 'Returns the standard deviation (population)' },
  { name: 'VAR', sig: 'VAR(value1, [value2, ...])', desc: 'Returns the variance (sample)' },
  { name: 'LARGE', sig: 'LARGE(range, k)', desc: 'Returns the k-th largest value' },
  { name: 'SMALL', sig: 'SMALL(range, k)', desc: 'Returns the k-th smallest value' },
  { name: 'RANK', sig: 'RANK(value, range, [order])', desc: 'Returns the rank of a value in a dataset' },
  { name: 'PERCENTILE', sig: 'PERCENTILE(range, percentile)', desc: 'Returns the value at a given percentile' },
  { name: 'CORREL', sig: 'CORREL(data_y, data_x)', desc: 'Returns the correlation coefficient' },
  { name: 'TODAY', sig: 'TODAY()', desc: 'Returns the current date' },
  { name: 'NOW', sig: 'NOW()', desc: 'Returns the current date and time' },
  { name: 'DATE', sig: 'DATE(year, month, day)', desc: 'Creates a date value' },
  { name: 'YEAR', sig: 'YEAR(date)', desc: 'Returns the year from a date' },
  { name: 'MONTH', sig: 'MONTH(date)', desc: 'Returns the month from a date' },
  { name: 'DAY', sig: 'DAY(date)', desc: 'Returns the day from a date' },
  { name: 'HOUR', sig: 'HOUR(time)', desc: 'Returns the hour from a time' },
  { name: 'MINUTE', sig: 'MINUTE(time)', desc: 'Returns the minute from a time' },
  { name: 'SECOND', sig: 'SECOND(time)', desc: 'Returns the second from a time' },
  { name: 'WEEKDAY', sig: 'WEEKDAY(date, [type])', desc: 'Returns the day of the week' },
  { name: 'EOMONTH', sig: 'EOMONTH(start_date, months)', desc: 'Returns the last day of a month' },
  { name: 'EDATE', sig: 'EDATE(start_date, months)', desc: 'Returns a date offset by months' },
  { name: 'DAYS', sig: 'DAYS(end_date, start_date)', desc: 'Returns the number of days between two dates' },
  { name: 'DATEDIF', sig: 'DATEDIF(start_date, end_date, unit)', desc: 'Calculates difference between dates in given unit' },
  { name: 'DATEVALUE', sig: 'DATEVALUE(date_string)', desc: 'Converts a date string to a date value' },
  { name: 'NETWORKDAYS', sig: 'NETWORKDAYS(start, end, [holidays])', desc: 'Returns the number of working days between dates' },
  { name: 'SWITCH', sig: 'SWITCH(expression, case1, value1, [case2, value2, ...], [default])', desc: 'Matches an expression against a list of cases' },
  { name: 'CHOOSE', sig: 'CHOOSE(index, choice1, [choice2, ...])', desc: 'Returns an element from a list based on index' },
  { name: 'ISBLANK', sig: 'ISBLANK(value)', desc: 'Returns TRUE if the cell is empty' },
  { name: 'ISERROR', sig: 'ISERROR(value)', desc: 'Returns TRUE if the value is an error' },
  { name: 'ISNA', sig: 'ISNA(value)', desc: 'Returns TRUE if the value is #N/A' },
  { name: 'ISNUMBER', sig: 'ISNUMBER(value)', desc: 'Returns TRUE if the value is a number' },
  { name: 'ISTEXT', sig: 'ISTEXT(value)', desc: 'Returns TRUE if the value is text' },
  { name: 'ISLOGICAL', sig: 'ISLOGICAL(value)', desc: 'Returns TRUE if the value is a boolean' },
  { name: 'TYPE', sig: 'TYPE(value)', desc: 'Returns the type of a value as a number' },
  { name: 'TRUE', sig: 'TRUE()', desc: 'Returns the logical value TRUE' },
  { name: 'FALSE', sig: 'FALSE()', desc: 'Returns the logical value FALSE' },
  { name: 'NA', sig: 'NA()', desc: 'Returns the #N/A error value' },
  { name: 'TRANSPOSE', sig: 'TRANSPOSE(range)', desc: 'Transposes rows and columns of a range' },
  { name: 'UNIQUE', sig: 'UNIQUE(range)', desc: 'Returns unique values from a range' },
  { name: 'SORT', sig: 'SORT(range, [sort_column], [ascending])', desc: 'Sorts the rows of a range' },
  { name: 'SPARKLINE', sig: 'SPARKLINE(data, [options])', desc: 'Creates a miniature chart in a cell' },
  { name: 'FILTER', sig: 'FILTER(range, condition, [if_empty])', desc: 'Filters a range based on a condition' },
  { name: 'FIXED', sig: 'FIXED(number, [decimals], [no_commas])', desc: 'Formats a number with a fixed number of decimal places' },
  { name: 'PMT', sig: 'PMT(rate, nper, pv, [fv], [type])', desc: 'Calculates the payment for a loan' },
  { name: 'FV', sig: 'FV(rate, nper, pmt, [pv], [type])', desc: 'Calculates the future value of an investment' },
  { name: 'PV', sig: 'PV(rate, nper, pmt, [fv], [type])', desc: 'Calculates the present value of an investment' },
  { name: 'NPV', sig: 'NPV(rate, value1, [value2, ...])', desc: 'Calculates the net present value' },
  { name: 'IRR', sig: 'IRR(cashflows, [guess])', desc: 'Calculates the internal rate of return' },
  { name: 'RATE', sig: 'RATE(nper, pmt, pv, [fv], [type], [guess])', desc: 'Calculates the interest rate per period' },
  { name: 'SLN', sig: 'SLN(cost, salvage, life)', desc: 'Returns straight-line depreciation' },
  { name: 'SIN', sig: 'SIN(angle)', desc: 'Returns the sine of an angle in radians' },
  { name: 'COS', sig: 'COS(angle)', desc: 'Returns the cosine of an angle in radians' },
  { name: 'TAN', sig: 'TAN(angle)', desc: 'Returns the tangent of an angle in radians' },
  { name: 'ASIN', sig: 'ASIN(value)', desc: 'Returns the arcsine in radians' },
  { name: 'ACOS', sig: 'ACOS(value)', desc: 'Returns the arccosine in radians' },
  { name: 'ATAN', sig: 'ATAN(value)', desc: 'Returns the arctangent in radians' },
  { name: 'ATAN2', sig: 'ATAN2(x, y)', desc: 'Returns the angle from x-axis to a point' },
  { name: 'DEGREES', sig: 'DEGREES(angle)', desc: 'Converts radians to degrees' },
  { name: 'RADIANS', sig: 'RADIANS(angle)', desc: 'Converts degrees to radians' },
  { name: 'EVEN', sig: 'EVEN(number)', desc: 'Rounds up to the nearest even integer' },
  { name: 'ODD', sig: 'ODD(number)', desc: 'Rounds up to the nearest odd integer' },
  { name: 'FACT', sig: 'FACT(number)', desc: 'Returns the factorial' },
  { name: 'GCD', sig: 'GCD(value1, value2, [...])', desc: 'Returns the greatest common divisor' },
  { name: 'LCM', sig: 'LCM(value1, value2, [...])', desc: 'Returns the least common multiple' },
  { name: 'NUMBERVALUE', sig: 'NUMBERVALUE(text, [decimal_sep], [group_sep])', desc: 'Converts localized text to a number' },
  { name: 'T', sig: 'T(value)', desc: 'Returns text if value is text, otherwise empty' },
  { name: 'N', sig: 'N(value)', desc: 'Returns the value as a number' },
];

// Index for fast lookup
const FUNC_MAP = new Map(FUNC_CATALOG.map(f => [f.name, f]));

// Characters after which we enter "point mode" (cell reference insertion)
const POINT_MODE_TRIGGERS = new Set(['(', ',', ';', '+', '-', '*', '/', '^', '&', '=', '<', '>', '!', ' ']);

// Reference colors for highlighting (Google Sheets style)
const REF_COLORS = [
  '#4285f4', '#ea4335', '#9334e6', '#e8710a',
  '#34a853', '#fa7b17', '#f538a0', '#00897b',
  '#1a73e8', '#d93025', '#7627bb', '#f9ab00',
];

export default class FormulaHelper {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.dropdown = null;
    this.hint = null;
    this.active = false;
    this.pointMode = false;
    this.pointAnchor = null;    // { row, col } where point-mode selection started
    this.pointCurrent = null;   // { row, col } current point-mode position
    this.formulaRefs = [];      // parsed cell references for highlighting
    this._selectedIndex = 0;
    this._matches = [];
    this._currentFunc = null;
    this._currentArgIndex = 0;
  }

  init(container) {
    this.container = container;

    this.dropdown = el('div', { className: 'sheets-formula-dropdown' });
    this.dropdown.style.display = 'none';
    container.appendChild(this.dropdown);

    this.hint = el('div', { className: 'sheets-formula-hint' });
    this.hint.style.display = 'none';
    container.appendChild(this.hint);
  }

  // Called by Editor on every input change while editing a formula
  onInput(text, cursorPos, textareaRect) {
    if (!text.startsWith('=')) {
      this.hide();
      return;
    }

    this.active = true;
    this._textareaRect = textareaRect;

    const beforeCursor = text.substring(0, cursorPos);

    // Parse formula references for coloring
    this.formulaRefs = this._parseRefs(text.substring(1));

    // Check if we should show autocomplete or hint
    const tokenInfo = this._getTokenAtCursor(beforeCursor);

    // Always update point mode first
    this._updatePointMode(beforeCursor, tokenInfo);

    if (tokenInfo.type === 'funcName') {
      this._showAutocomplete(tokenInfo.value, textareaRect);
      this._hideHint();
    } else if (tokenInfo.type === 'insideFunc') {
      this._hideAutocomplete();
      this._showHint(tokenInfo.funcName, tokenInfo.argIndex, textareaRect);
    } else {
      this._hideAutocomplete();
      this._hideHint();
    }
  }

  _getTokenAtCursor(beforeCursor) {
    const formula = beforeCursor.substring(1); // strip leading =
    if (!formula) return { type: 'empty' };

    // Walk backward to find what we're in the middle of
    let i = formula.length - 1;

    // Check if we're typing a function name (letters right after = or operator)
    const funcNameMatch = formula.match(/([A-Za-z_]\w*)$/);
    if (funcNameMatch) {
      // Make sure it's not inside a completed function call
      const before = formula.substring(0, formula.length - funcNameMatch[1].length);
      const openParens = (before.match(/\(/g) || []).length;
      const closeParens = (before.match(/\)/g) || []).length;
      // If unclosed parens, this might be an argument not a new function
      // But if the char before is an operator or start, it's a function name
      const charBefore = before.slice(-1);
      if (!charBefore || POINT_MODE_TRIGGERS.has(charBefore) || charBefore === ')') {
        // Check if next char will be '(' — that means we already completed it
        // For now, treat as function name being typed
        return { type: 'funcName', value: funcNameMatch[1].toUpperCase() };
      }
    }

    // Check if we're inside a function call
    let depth = 0;
    let funcName = null;
    let argIndex = 0;
    for (let j = formula.length - 1; j >= 0; j--) {
      const ch = formula[j];
      if (ch === ')') depth++;
      else if (ch === '(') {
        if (depth === 0) {
          // Find function name before this paren
          const beforeParen = formula.substring(0, j);
          const nameMatch = beforeParen.match(/([A-Za-z_]\w*)$/);
          if (nameMatch) {
            funcName = nameMatch[1].toUpperCase();
            // Count commas to get arg index
            const inside = formula.substring(j + 1);
            let d = 0;
            for (const c of inside) {
              if (c === '(') d++;
              else if (c === ')') d--;
              else if ((c === ',' || c === ';') && d === 0) argIndex++;
            }
            return { type: 'insideFunc', funcName, argIndex };
          }
          break;
        }
        depth--;
      }
    }

    return { type: 'other' };
  }

  _showAutocomplete(partial, rect) {
    if (!partial || partial.length < 1) {
      this._hideAutocomplete();
      return;
    }

    const upper = partial.toUpperCase();
    this._matches = FUNC_CATALOG.filter(f => f.name.startsWith(upper)).slice(0, 8);

    if (this._matches.length === 0 || (this._matches.length === 1 && this._matches[0].name === upper)) {
      this._hideAutocomplete();
      return;
    }

    this._selectedIndex = 0;
    this._renderDropdown(rect);
  }

  _renderDropdown(rect) {
    this.dropdown.innerHTML = '';
    this.dropdown.style.display = 'block';

    // Position below the textarea
    if (rect) {
      this.dropdown.style.left = rect.left + 'px';
      this.dropdown.style.top = (rect.bottom + 2) + 'px';
    }

    for (let i = 0; i < this._matches.length; i++) {
      const func = this._matches[i];
      const item = el('div', {
        className: 'sheets-formula-dropdown-item' + (i === this._selectedIndex ? ' selected' : ''),
      });

      const nameSpan = el('span', { className: 'sheets-fd-name' }, func.name);
      const descSpan = el('span', { className: 'sheets-fd-desc' }, func.desc);

      item.appendChild(nameSpan);
      item.appendChild(descSpan);

      const idx = i;
      item.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        this._acceptAutocomplete(idx);
      });
      item.addEventListener('mouseenter', () => {
        this._selectedIndex = idx;
        this._renderDropdown(rect);
      });

      this.dropdown.appendChild(item);
    }
  }

  _acceptAutocomplete(index) {
    const func = this._matches[index];
    if (!func) return;

    const editor = this.spreadsheet.editor;
    if (!editor || !editor.isActive) return;

    const textarea = editor.textarea;
    const text = textarea.value;
    const cursorPos = textarea.selectionStart;

    // Find the start of the partial function name
    const beforeCursor = text.substring(0, cursorPos);
    const nameMatch = beforeCursor.match(/([A-Za-z_]\w*)$/);
    if (!nameMatch) return;

    const start = cursorPos - nameMatch[1].length;
    const completion = func.name + '(';
    const newText = text.substring(0, start) + completion + text.substring(cursorPos);

    textarea.value = newText;
    const newPos = start + completion.length;
    textarea.selectionStart = newPos;
    textarea.selectionEnd = newPos;
    textarea.focus();

    this._hideAutocomplete();

    // Trigger input event so editor updates
    textarea.dispatchEvent(new Event('input'));
  }

  _hideAutocomplete() {
    if (this.dropdown) this.dropdown.style.display = 'none';
    this._matches = [];
  }

  _showHint(funcName, argIndex, rect) {
    const func = FUNC_MAP.get(funcName);
    if (!func) {
      this._hideHint();
      return;
    }

    this._currentFunc = func;
    this._currentArgIndex = argIndex;

    // Parse signature to highlight current argument
    const sig = func.sig;
    const openParen = sig.indexOf('(');
    const closeParen = sig.lastIndexOf(')');
    if (openParen === -1) { this._hideHint(); return; }

    const fnPart = sig.substring(0, openParen + 1);
    const argsPart = sig.substring(openParen + 1, closeParen);
    const args = this._splitArgs(argsPart);

    this.hint.innerHTML = '';
    this.hint.style.display = 'block';

    if (rect) {
      this.hint.style.left = rect.left + 'px';
      this.hint.style.top = (rect.bottom + 2) + 'px';
    }

    const fnSpan = el('span', { className: 'sheets-fh-func' }, fnPart);
    this.hint.appendChild(fnSpan);

    for (let i = 0; i < args.length; i++) {
      if (i > 0) this.hint.appendChild(document.createTextNode(', '));
      const argSpan = el('span', {
        className: 'sheets-fh-arg' + (i === argIndex ? ' current' : ''),
      }, args[i].trim());
      this.hint.appendChild(argSpan);
    }

    this.hint.appendChild(document.createTextNode(')'));

    // Add description below
    const descDiv = el('div', { className: 'sheets-fh-desc' }, func.desc);
    this.hint.appendChild(descDiv);
  }

  _splitArgs(argsStr) {
    const args = [];
    let depth = 0;
    let current = '';
    for (const ch of argsStr) {
      if (ch === '[' || ch === '(') depth++;
      if (ch === ']' || ch === ')') depth--;
      if (ch === ',' && depth === 0) {
        args.push(current);
        current = '';
      } else {
        current += ch;
      }
    }
    if (current) args.push(current);
    return args;
  }

  _hideHint() {
    if (this.hint) this.hint.style.display = 'none';
    this._currentFunc = null;
  }

  // ── Point Mode (cell selection during formula editing) ──

  _updatePointMode(beforeCursor, tokenInfo) {
    if (!beforeCursor || beforeCursor.length < 1) {
      this.pointMode = false;
      return;
    }

    // Point mode = arrow keys / clicks insert cell references instead of moving cursor.
    // Active when cursor is right after: = ( , ; + - * / ^ & < > <= >= <> !
    // Also active when cursor is right after a cell reference we just inserted (so
    // repeated arrows extend/move the ref). NOT active when typing function names,
    // strings, or numbers (arrow keys should move text cursor).

    const trimmed = beforeCursor.trimEnd();
    const lastChar = trimmed.slice(-1);

    if (POINT_MODE_TRIGGERS.has(lastChar)) {
      // Just typed an operator — ready for new reference
      this.pointMode = true;
      this.pointAnchor = null;
      this.pointCurrent = null;
    } else if (this.pointAnchor) {
      // We have an active anchor from a previous arrow/click — stay in point mode
      // so shift+arrow can extend. The anchor persists until user types an operator
      // or non-ref character.
      this.pointMode = true;
    } else {
      // Typing a function name, number, or string — not in point mode
      this.pointMode = false;
    }
  }

  // Handle autocomplete keyboard navigation (Tab/Enter/Arrow/Escape)
  handleAutocompleteKey(e, textarea) {
    if (!this.dropdown || this.dropdown.style.display === 'none' || this._matches.length === 0) {
      return false;
    }
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      this._selectedIndex = Math.min(this._selectedIndex + 1, this._matches.length - 1);
      this._renderDropdown(this._textareaRect);
      return true;
    }
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      this._selectedIndex = Math.max(this._selectedIndex - 1, 0);
      this._renderDropdown(this._textareaRect);
      return true;
    }
    if (e.key === 'Tab' || (e.key === 'Enter' && this._matches.length > 0)) {
      e.preventDefault();
      this._acceptAutocomplete(this._selectedIndex);
      return true;
    }
    if (e.key === 'Escape') {
      e.preventDefault();
      this._hideAutocomplete();
      return true;
    }
    return false;
  }

  // Handle arrow keys in POINT mode (insert/extend cell references)
  handlePointArrow(e, textarea, ss) {
    const sheet = ss.activeSheet;
    if (!sheet) return;

    const dr = e.key === 'ArrowDown' ? 1 : e.key === 'ArrowUp' ? -1 : 0;
    const dc = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0;

    if (!this.pointAnchor) {
      // First arrow: single cell reference at the target cell
      const r = Math.max(0, Math.min(ss.activeRow + dr, sheet.rowCount - 1));
      const c = Math.max(0, Math.min(ss.activeCol + dc, sheet.colCount - 1));
      this.pointAnchor = { row: r, col: c };
      this.pointCurrent = { row: r, col: c };
    } else if (e.shiftKey) {
      this.pointCurrent = {
        row: Math.max(0, Math.min(this.pointCurrent.row + dr, sheet.rowCount - 1)),
        col: Math.max(0, Math.min(this.pointCurrent.col + dc, sheet.colCount - 1)),
      };
    } else {
      this.pointAnchor = {
        row: Math.max(0, Math.min((this.pointCurrent || this.pointAnchor).row + dr, sheet.rowCount - 1)),
        col: Math.max(0, Math.min((this.pointCurrent || this.pointAnchor).col + dc, sheet.colCount - 1)),
      };
      this.pointCurrent = { ...this.pointAnchor };
    }

    this._insertRefAtCursor(textarea);
    ss.renderer.ensureCellVisible(this.pointCurrent.row, this.pointCurrent.col);
    ss.render();
  }

  // Called when user clicks a cell while editing a formula
  handlePointModeClick(row, col, shiftKey) {
    const editor = this.spreadsheet.editor;
    if (!editor || !editor.isActive) return false;

    const text = editor.textarea.value;
    if (!text.startsWith('=')) return false;

    // Always allow click-to-insert in formula mode
    const canInsert = this.pointMode || this.pointAnchor;

    if (!canInsert) return false;

    if (shiftKey && this.pointAnchor) {
      this.pointCurrent = { row, col };
    } else {
      this.pointAnchor = { row, col };
      this.pointCurrent = { row, col };
    }

    this._insertRefAtCursor(editor.textarea);
    this.spreadsheet.render();
    return true;
  }

  _insertRefAtCursor(textarea) {
    if (!this.pointAnchor || !this.pointCurrent) return;

    const text = textarea.value;
    const cursorPos = textarea.selectionStart;

    // Find the range of any existing reference we're replacing
    const { start, end } = this._findRefBounds(text, cursorPos);

    // Build reference string
    let ref;
    if (this.pointAnchor.row === this.pointCurrent.row &&
        this.pointAnchor.col === this.pointCurrent.col) {
      ref = cellRefToString(this.pointAnchor.col, this.pointAnchor.row);
    } else {
      const r1 = Math.min(this.pointAnchor.row, this.pointCurrent.row);
      const c1 = Math.min(this.pointAnchor.col, this.pointCurrent.col);
      const r2 = Math.max(this.pointAnchor.row, this.pointCurrent.row);
      const c2 = Math.max(this.pointAnchor.col, this.pointCurrent.col);
      ref = cellRefToString(c1, r1) + ':' + cellRefToString(c2, r2);
    }

    // Cross-sheet prefix: if clicking on a different sheet than the one being edited
    const ss = this.spreadsheet;
    const editSheetId = ss._formulaEditSheetId || (ss.editor ? ss.sheets.find(s => {
      const er = ss.editor.editRow, ec = ss.editor.editCol;
      return er >= 0 && s.getCell(er, ec);
    })?.id : null) || ss.activeSheet?.id;
    const viewSheetId = ss.activeSheet?.id;

    if (editSheetId && viewSheetId && editSheetId !== viewSheetId) {
      const viewSheet = ss.activeSheet;
      const name = viewSheet.name;
      // Quote sheet name if it contains spaces or special chars
      const needsQuote = /[^A-Za-z0-9_]/.test(name);
      const prefix = needsQuote ? "'" + name + "'!" : name + '!';
      ref = prefix + ref;
    }

    const newText = text.substring(0, start) + ref + text.substring(end);
    textarea.value = newText;
    const newPos = start + ref.length;
    textarea.selectionStart = newPos;
    textarea.selectionEnd = newPos;
    textarea.focus();
    textarea.dispatchEvent(new Event('input'));
  }

  _findRefBounds(text, cursorPos) {
    // Look backward from cursor for start of a cell ref (or operator)
    let start = cursorPos;
    let end = cursorPos;

    // If there's already a reference at cursor, select it for replacement
    // Walk back to find start of ref
    while (start > 0 && /[A-Za-z0-9$:]/.test(text[start - 1])) start--;

    // Check if what's between start and cursorPos looks like a cell ref
    const segment = text.substring(start, cursorPos);
    if (/^\$?[A-Za-z]+\$?\d*(:\$?[A-Za-z]+\$?\d*)?$/.test(segment)) {
      // Also consume any remaining ref chars after cursor
      while (end < text.length && /[A-Za-z0-9$:]/.test(text[end])) end++;
      return { start, end };
    }

    // No existing ref, insert at cursor
    return { start: cursorPos, end: cursorPos };
  }

  // ── Reference highlighting ──

  _parseRefs(formulaBody) {
    const refs = [];
    const regex = /(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?/gi;
    let match;
    let colorIdx = 0;

    while ((match = regex.exec(formulaBody)) !== null) {
      const color = REF_COLORS[colorIdx % REF_COLORS.length];
      const startRef = match[1].toUpperCase();
      const endRef = match[2] ? match[2].toUpperCase() : null;
      refs.push({ startRef, endRef, color });
      colorIdx++;
    }

    return refs;
  }

  getFormulaRefs() {
    return this.formulaRefs;
  }

  // Returns colored HTML for formula text (references get colored spans)
  getColoredFormulaHTML(text) {
    if (!text || !text.startsWith('=')) return this._escapeHTML(text);

    const body = text.substring(1); // after =
    const regex = /(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?/gi;
    const segments = [];
    let lastIdx = 0;
    let colorIdx = 0;

    let match;
    while ((match = regex.exec(body)) !== null) {
      const color = REF_COLORS[colorIdx % REF_COLORS.length];
      // Text before this match
      if (match.index > lastIdx) {
        segments.push(this._escapeHTML(body.substring(lastIdx, match.index)));
      }
      // The colored reference
      segments.push(`<span style="color:${color};font-weight:600">${this._escapeHTML(match[0])}</span>`);
      lastIdx = match.index + match[0].length;
      colorIdx++;
    }
    // Remainder
    if (lastIdx < body.length) {
      segments.push(this._escapeHTML(body.substring(lastIdx)));
    }

    return this._escapeHTML('=') + segments.join('');
  }

  _escapeHTML(str) {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  // ── Lifecycle ──

  hide() {
    this._hideAutocomplete();
    this._hideHint();
    this.active = false;
    this.pointMode = false;
    this.pointAnchor = null;
    this.pointCurrent = null;
    this.formulaRefs = [];
  }

  destroy() {
    if (this.dropdown && this.dropdown.parentElement) this.dropdown.remove();
    if (this.hint && this.hint.parentElement) this.hint.remove();
  }
}
