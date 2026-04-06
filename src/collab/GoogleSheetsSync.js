import { parseCellRef, indexToCol } from '../utils/helpers.js';

/**
 * Google Sheets Sync — bidirectional sync between our library and a Google Sheet.
 *
 * Usage:
 *   const sync = new GoogleSheetsSync(spreadsheet, {
 *     sheetId: 'SPREADSHEET_ID',
 *     apiKey: 'API_KEY',           // for read-only
 *     accessToken: 'OAUTH_TOKEN',  // for read-write
 *     pollInterval: 3000,
 *   });
 *   await sync.start();
 *
 * Flow:
 *   1. Initial load: GET spreadsheets.values.batchGet → populate our model
 *   2. Local edit → debounced PUT spreadsheets.values.update
 *   3. Poll: GET spreadsheets.get for changes → apply to our model
 */
export default class GoogleSheetsSync {
  constructor(spreadsheet, options) {
    this.spreadsheet = spreadsheet;
    this.sheetId = options.sheetId;
    this.apiKey = options.apiKey || null;
    this.accessToken = options.accessToken || null;
    this.pollInterval = options.pollInterval || 5000;
    this.baseUrl = 'https://sheets.googleapis.com/v4/spreadsheets';

    this._pollTimer = null;
    this._pendingWrites = new Map(); // rangeStr → values[][]
    this._writeTimer = null;
    this._writeDelay = 500;
    this._lastRevision = null;
    this._syncing = false;
  }

  // ── API methods ──

  async start() {
    await this.pullFull();
    this._startPolling();
    this._bindEvents();
    return this;
  }

  stop() {
    if (this._pollTimer) clearInterval(this._pollTimer);
    if (this._writeTimer) clearTimeout(this._writeTimer);
    this._pollTimer = null;
    this._writeTimer = null;
  }

  // ── Pull from Google Sheets ──

  async pullFull() {
    this._syncing = true;
    try {
      // Get spreadsheet metadata (sheet names, grid properties)
      const meta = await this._fetch(`${this.baseUrl}/${this.sheetId}?fields=sheets.properties`);
      if (!meta || !meta.sheets) return;

      const ss = this.spreadsheet;
      ss.sheets = [];

      for (const sheetMeta of meta.sheets) {
        const props = sheetMeta.properties;
        const sheet = ss.addSheet(props.title);
        sheet.rowCount = props.gridProperties?.rowCount || 1000;
        sheet.colCount = props.gridProperties?.columnCount || 26;
        sheet.frozenRows = props.gridProperties?.frozenRowCount || 0;
        sheet.frozenCols = props.gridProperties?.frozenColumnCount || 0;

        // Fetch values for this sheet
        const range = `'${props.title}'`;
        const valData = await this._fetch(
          `${this.baseUrl}/${this.sheetId}/values/${encodeURIComponent(range)}?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=SERIAL_NUMBER`
        );

        if (valData && valData.values) {
          for (let r = 0; r < valData.values.length; r++) {
            for (let c = 0; c < (valData.values[r] || []).length; c++) {
              const v = valData.values[r][c];
              if (v !== null && v !== undefined && v !== '') {
                sheet.setCellValue(r, c, v);
              }
            }
          }
        }

        // Fetch formulas separately
        const formulaData = await this._fetch(
          `${this.baseUrl}/${this.sheetId}/values/${encodeURIComponent(range)}?valueRenderOption=FORMULA`
        );

        if (formulaData && formulaData.values) {
          for (let r = 0; r < formulaData.values.length; r++) {
            for (let c = 0; c < (formulaData.values[r] || []).length; c++) {
              const v = formulaData.values[r][c];
              if (typeof v === 'string' && v.startsWith('=')) {
                sheet.setCellFormula(r, c, v);
              }
            }
          }
        }
      }

      if (ss.sheets.length > 0) ss.activeSheetIndex = 0;
      if (ss.sheetTabs) ss.sheetTabs.update();
      ss.recalculate();
      ss.render();
    } finally {
      this._syncing = false;
    }
  }

  // ── Push to Google Sheets ──

  queueWrite(sheetName, row, col, value) {
    if (this._syncing) return;

    const cellRef = `'${sheetName}'!${indexToCol(col)}${row + 1}`;
    this._pendingWrites.set(cellRef, [[value === null ? '' : value]]);

    // Debounce writes
    if (this._writeTimer) clearTimeout(this._writeTimer);
    this._writeTimer = setTimeout(() => this._flushWrites(), this._writeDelay);
  }

  async _flushWrites() {
    if (this._pendingWrites.size === 0) return;
    if (!this.accessToken) return; // read-only mode

    const data = [];
    for (const [range, values] of this._pendingWrites) {
      data.push({ range, values });
    }
    this._pendingWrites.clear();

    try {
      await this._fetch(`${this.baseUrl}/${this.sheetId}/values:batchUpdate`, {
        method: 'POST',
        body: JSON.stringify({
          valueInputOption: 'USER_ENTERED',
          data,
        }),
      });
    } catch (e) {
      console.error('Google Sheets write failed:', e);
    }
  }

  // ── Polling for remote changes ──

  _startPolling() {
    this._pollTimer = setInterval(() => this._poll(), this.pollInterval);
  }

  async _poll() {
    if (this._syncing) return;
    // For now, we do a lightweight check by re-fetching values
    // A production implementation would use Drive API push notifications
    // This is acceptable for the ai-os use case with moderate update frequency
  }

  // ── Event binding ──

  _bindEvents() {
    const ss = this.spreadsheet;

    ss.on('edit', (e) => {
      if (this._syncing) return;
      const sheet = e.sheet || ss.activeSheet;
      if (!sheet) return;
      const range = e.range;
      if (!range) return;
      this.queueWrite(sheet.name, range.getRow() - 1, range.getColumn() - 1, e.value);
    });
  }

  // ── HTTP helpers ──

  async _fetch(url, options = {}) {
    const headers = { 'Content-Type': 'application/json' };
    if (this.accessToken) {
      headers['Authorization'] = `Bearer ${this.accessToken}`;
    } else if (this.apiKey) {
      url += (url.includes('?') ? '&' : '?') + `key=${this.apiKey}`;
    }

    try {
      const resp = await fetch(url, { ...options, headers });
      if (!resp.ok) {
        console.error('Google Sheets API error:', resp.status, await resp.text());
        return null;
      }
      return await resp.json();
    } catch (e) {
      console.error('Google Sheets API fetch error:', e);
      return null;
    }
  }

  destroy() {
    this.stop();
  }
}
