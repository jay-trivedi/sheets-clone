import CellRange from '../core/CellRange.js';

/**
 * Sheet & Range Protection
 *
 * Supports:
 * - Whole sheet protection (all cells locked by default, specific ranges unlocked)
 * - Range protection (specific ranges locked)
 * - API for agents to set/check permissions
 */
export default class ProtectionManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
  }

  // ── Sheet-level protection ──

  protectSheet(sheet, options = {}) {
    sheet._protected = true;
    sheet._protectionOptions = {
      formatCells: options.formatCells ?? false,
      formatColumns: options.formatColumns ?? false,
      formatRows: options.formatRows ?? false,
      insertColumns: options.insertColumns ?? false,
      insertRows: options.insertRows ?? false,
      deleteColumns: options.deleteColumns ?? false,
      deleteRows: options.deleteRows ?? false,
      sort: options.sort ?? false,
      autoFilter: options.autoFilter ?? false,
      editObjects: options.editObjects ?? false,
    };
    sheet._unprotectedRanges = options.unprotectedRanges || [];
  }

  unprotectSheet(sheet) {
    sheet._protected = false;
    sheet._protectionOptions = null;
    sheet._unprotectedRanges = [];
  }

  isSheetProtected(sheet) {
    return !!sheet._protected;
  }

  // ── Range-level protection ──

  protectRange(sheet, rangeStr, description = '') {
    if (!sheet._protectedRanges) sheet._protectedRanges = [];
    sheet._protectedRanges.push({ range: rangeStr, description });
  }

  unprotectRange(sheet, rangeStr) {
    if (!sheet._protectedRanges) return;
    sheet._protectedRanges = sheet._protectedRanges.filter(r => r.range !== rangeStr);
  }

  getProtectedRanges(sheet) {
    return sheet._protectedRanges || [];
  }

  // ── Check if editing is allowed ──

  canEdit(sheet, row, col) {
    if (!sheet._protected && !sheet._protectedRanges?.length) return true;

    // Sheet protection
    if (sheet._protected) {
      // Check if cell is in an unprotected range
      for (const rangeStr of sheet._unprotectedRanges) {
        const range = CellRange.fromString(rangeStr);
        if (range && range.contains(row, col)) return true;
      }
      return false; // Protected sheet, cell not in unprotected range
    }

    // Range protection
    if (sheet._protectedRanges) {
      for (const { range: rangeStr } of sheet._protectedRanges) {
        const range = CellRange.fromString(rangeStr);
        if (range && range.contains(row, col)) return false;
      }
    }

    return true;
  }

  // ── Check structural permissions ──

  canInsertRows(sheet) {
    return !sheet._protected || sheet._protectionOptions?.insertRows;
  }

  canInsertCols(sheet) {
    return !sheet._protected || sheet._protectionOptions?.insertColumns;
  }

  canDeleteRows(sheet) {
    return !sheet._protected || sheet._protectionOptions?.deleteRows;
  }

  canDeleteCols(sheet) {
    return !sheet._protected || sheet._protectionOptions?.deleteColumns;
  }

  canSort(sheet) {
    return !sheet._protected || sheet._protectionOptions?.sort;
  }

  canFormat(sheet) {
    return !sheet._protected || sheet._protectionOptions?.formatCells;
  }

  destroy() {}
}
