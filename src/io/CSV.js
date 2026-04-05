export function parseCSV(text, delimiter = ',') {
  const rows = [];
  let row = [];
  let cell = '';
  let inQuote = false;
  let i = 0;

  while (i < text.length) {
    const ch = text[i];

    if (inQuote) {
      if (ch === '"') {
        if (i + 1 < text.length && text[i + 1] === '"') {
          cell += '"';
          i += 2;
        } else {
          inQuote = false;
          i++;
        }
      } else {
        cell += ch;
        i++;
      }
    } else {
      if (ch === '"') {
        inQuote = true;
        i++;
      } else if (ch === delimiter) {
        row.push(cell);
        cell = '';
        i++;
      } else if (ch === '\n' || ch === '\r') {
        row.push(cell);
        cell = '';
        rows.push(row);
        row = [];
        if (ch === '\r' && i + 1 < text.length && text[i + 1] === '\n') i++;
        i++;
      } else {
        cell += ch;
        i++;
      }
    }
  }

  if (cell || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }

  return rows;
}

export function generateCSV(sheet, delimiter = ',') {
  const used = sheet.getUsedRange();
  if (!used) return '';

  const rows = [];
  for (let r = used.startRow; r <= used.endRow; r++) {
    const cols = [];
    for (let c = used.startCol; c <= used.endCol; c++) {
      const cell = sheet.getCell(r, c);
      let val = cell ? cell.displayValue : '';
      // Escape quotes and wrap if needed
      if (val.includes(delimiter) || val.includes('"') || val.includes('\n')) {
        val = '"' + val.replace(/"/g, '""') + '"';
      }
      cols.push(val);
    }
    rows.push(cols.join(delimiter));
  }

  return rows.join('\n');
}
