export default class NumberFormat {
  constructor() {
    this._cache = new Map();
  }

  format(value, fmt) {
    if (fmt === 'General' || !fmt) {
      return this._formatGeneral(value);
    }
    if (fmt === '@') return String(value);

    const key = fmt + ':' + value;
    if (this._cache.has(key)) return this._cache.get(key);

    let result;
    try {
      result = this._applyFormat(value, fmt);
    } catch (e) {
      result = String(value);
    }

    if (this._cache.size > 5000) this._cache.clear();
    this._cache.set(key, result);
    return result;
  }

  _formatGeneral(value) {
    if (value === null || value === undefined) return '';
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    if (typeof value === 'string') return value;
    if (typeof value !== 'number') return String(value);

    // Smart number formatting
    if (Number.isInteger(value) && Math.abs(value) < 1e15) {
      return value.toString();
    }
    if (Math.abs(value) >= 1e11 || (Math.abs(value) < 0.0001 && value !== 0)) {
      return value.toExponential(5);
    }
    // Avoid floating point artifacts
    const str = value.toPrecision(10);
    return parseFloat(str).toString();
  }

  _applyFormat(value, fmt) {
    // Handle conditional sections (positive;negative;zero)
    const sections = fmt.split(';');
    if (sections.length > 1) {
      if (value > 0) return this._applyFormat(value, sections[0]);
      if (value < 0) return this._applyFormat(Math.abs(value), sections[1] || sections[0]);
      return this._applyFormat(value, sections[2] || sections[0]);
    }

    // Percentage
    if (fmt.includes('%')) {
      const pctValue = value * 100;
      const decimals = (fmt.match(/0/g) || []).length - 1;
      return pctValue.toFixed(Math.max(0, decimals)) + '%';
    }

    // Currency / Accounting
    const currencyMatch = fmt.match(/^([_\(]?)([$£€¥₹])/);
    let prefix = '', suffix = '';
    let numFmt = fmt;

    if (currencyMatch) {
      prefix = currencyMatch[2];
      numFmt = fmt.replace(/[_\(]?[$£€¥₹]\s?\*?\s?/, '').replace(/[_\)]/, '');
    }

    if (fmt.startsWith('$') || fmt.includes('$')) {
      prefix = '$';
      numFmt = fmt.replace(/\$/g, '');
    }

    // Scientific notation
    if (fmt.includes('E+') || fmt.includes('E-')) {
      const decimals = (fmt.split('.')[1] || '').replace(/E.*/, '').length;
      return value.toExponential(decimals).toUpperCase();
    }

    // Date formats
    if (fmt.includes('M') || fmt.includes('d') || fmt.includes('y') ||
        fmt.includes('m') || fmt.includes('h') || fmt.includes('s')) {
      return this._formatDate(value, fmt);
    }

    // Number format
    return prefix + this._formatNumber(value, numFmt) + suffix;
  }

  _formatNumber(value, fmt) {
    const hasComma = fmt.includes(',');
    const parts = fmt.split('.');
    const intFmt = parts[0].replace(/[#,]/g, '');
    const decFmt = parts[1] || '';

    const minInts = (intFmt.match(/0/g) || []).length;
    const decimals = decFmt.length;

    const isNeg = value < 0;
    let num = Math.abs(value);

    let str;
    if (decimals > 0) {
      str = num.toFixed(decimals);
    } else {
      str = Math.round(num).toString();
    }

    let [intPart, decPart] = str.split('.');

    // Pad integer part
    while (intPart.length < minInts) intPart = '0' + intPart;

    // Add commas
    if (hasComma) {
      intPart = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    }

    let result = decPart ? intPart + '.' + decPart : intPart;
    if (isNeg) result = '-' + result;

    return result;
  }

  _formatDate(serial, fmt) {
    const d = new Date(1899, 11, 30);
    const wholeDays = Math.floor(serial);
    const frac = serial - wholeDays;

    d.setDate(d.getDate() + wholeDays);

    const year = d.getFullYear();
    const month = d.getMonth() + 1;
    const day = d.getDate();
    const hours24 = Math.floor(frac * 24);
    const minutes = Math.floor((frac * 1440) % 60);
    const seconds = Math.floor((frac * 86400) % 60);

    const isAMPM = fmt.includes('AM/PM') || fmt.includes('am/pm');
    let hours = hours24;
    let ampm = '';
    if (isAMPM) {
      ampm = hours24 >= 12 ? 'PM' : 'AM';
      hours = hours24 % 12 || 12;
    }

    const months = ['January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'];
    const monthsShort = months.map(m => m.substring(0, 3));
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

    return fmt
      .replace('yyyy', year)
      .replace('yy', String(year).slice(-2))
      .replace('MMMM', months[month - 1])
      .replace('MMM', monthsShort[month - 1])
      .replace('MM', String(month).padStart(2, '0'))
      .replace(/M(?!M)/, month)
      .replace('dd', String(day).padStart(2, '0'))
      .replace(/d(?!d)/, day)
      .replace('hh', String(hours).padStart(2, '0'))
      .replace(/h(?!h)/, hours)
      .replace('mm', String(minutes).padStart(2, '0'))
      .replace('ss', String(seconds).padStart(2, '0'))
      .replace('AM/PM', ampm)
      .replace('am/pm', ampm.toLowerCase());
  }
}
