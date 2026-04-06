import Spreadsheet from './core/Spreadsheet.js';
import Sheet from './core/Sheet.js';
import Cell, { CellStyle } from './core/Cell.js';
import CellRange from './core/CellRange.js';
import Range from './api/Range.js';
import FormulaEngine from './formula/FormulaEngine.js';
import YjsAdapter from './collab/YjsAdapter.js';
import GoogleSheetsSync from './collab/GoogleSheetsSync.js';
import { parseCSV, generateCSV } from './io/CSV.js';

export {
  Spreadsheet, Sheet, Cell, CellStyle, CellRange, Range,
  FormulaEngine, YjsAdapter, GoogleSheetsSync,
  parseCSV, generateCSV,
};
export default Spreadsheet;

// UMD-style global export
if (typeof window !== 'undefined') {
  window.Sheets = {
    Spreadsheet, Sheet, Cell, CellStyle, CellRange, Range,
    FormulaEngine, YjsAdapter, GoogleSheetsSync,
    parseCSV, generateCSV,
  };
}
