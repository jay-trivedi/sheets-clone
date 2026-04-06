import * as Y from 'yjs';
import { cellKey, keyToRC, indexToCol } from '../utils/helpers.js';

/**
 * Yjs Adapter — bridges the Spreadsheet data model with a Yjs CRDT document.
 *
 * The Y.Doc structure:
 *   doc.getMap('meta')           → { activeSheetIndex, sheetOrder: [...ids] }
 *   doc.getMap('sheets')         → Map<sheetId, Y.Map>
 *     sheet.get('name')          → string
 *     sheet.get('colCount')      → number
 *     sheet.get('rowCount')      → number
 *     sheet.get('frozenRows')    → number
 *     sheet.get('frozenCols')    → number
 *     sheet.get('colWidths')     → Y.Map<col, width>
 *     sheet.get('rowHeights')    → Y.Map<row, height>
 *     sheet.get('cells')         → Y.Map<cellKey, Y.Map>
 *       cell.get('v')            → raw value
 *       cell.get('f')            → formula
 *       cell.get('s')            → style JSON
 *   doc.getMap('awareness')      → managed by awareness protocol
 *
 * Changes flow:
 *   Local edit → update Spreadsheet → YjsAdapter.pushLocal() → Y.Doc → network → peers
 *   Remote edit → Y.Doc observer → YjsAdapter.applyRemote() → update Spreadsheet → render
 */
export default class YjsAdapter {
  constructor(spreadsheet, options = {}) {
    this.spreadsheet = spreadsheet;
    this.doc = options.doc || new Y.Doc();
    this.provider = null;
    this.awareness = null;
    this._suppressRemote = false;
    this._suppressLocal = false;
    this._initialized = false;

    // User identity
    this.userId = options.userId || Math.random().toString(36).substring(2, 8);
    this.userName = options.userName || 'User';
    this.userColor = options.userColor || this._randomColor();
  }

  /**
   * Connect to a WebSocket collaboration server.
   */
  connect(url, roomId) {
    // Dynamic import to avoid bundling y-websocket if not used
    return import('y-websocket').then(({ WebsocketProvider }) => {
      this.provider = new WebsocketProvider(url, roomId, this.doc);
      this.awareness = this.provider.awareness;

      // Set local user state
      this.awareness.setLocalState({
        user: { id: this.userId, name: this.userName, color: this.userColor },
        cursor: null,
        selection: null,
      });

      // Listen for awareness changes (other users' cursors)
      this.awareness.on('change', () => this._onAwarenessChange());

      this._setupObservers();
      this._initialized = true;

      // If doc already has data (reconnecting), apply it
      if (this.doc.getMap('sheets').size > 0) {
        this._applyFullState();
      } else {
        // First connection — push local state to doc
        this.pushFullState();
      }

      return this;
    });
  }

  /**
   * Connect without a server — for local-only Yjs (testing, offline).
   */
  connectLocal() {
    this._setupObservers();
    this._initialized = true;
    this.pushFullState();
    return this;
  }

  /**
   * Disconnect from the collaboration server.
   */
  disconnect() {
    if (this.provider) {
      this.provider.disconnect();
      this.provider.destroy();
      this.provider = null;
    }
    this._initialized = false;
  }

  // ── Push local state to Y.Doc ──

  pushFullState() {
    this.doc.transact(() => {
      const ss = this.spreadsheet;
      const sheetsMap = this.doc.getMap('sheets');
      const meta = this.doc.getMap('meta');

      meta.set('sheetOrder', ss.sheets.map(s => s.id));

      for (const sheet of ss.sheets) {
        this._pushSheet(sheetsMap, sheet);
      }
    });
  }

  _pushSheet(sheetsMap, sheet) {
    let ySheet = sheetsMap.get(sheet.id);
    if (!ySheet) {
      ySheet = new Y.Map();
      sheetsMap.set(sheet.id, ySheet);
    }

    ySheet.set('name', sheet.name);
    ySheet.set('colCount', sheet.colCount);
    ySheet.set('rowCount', sheet.rowCount);
    ySheet.set('frozenRows', sheet.frozenRows);
    ySheet.set('frozenCols', sheet.frozenCols);

    // Cells
    let yCells = ySheet.get('cells');
    if (!yCells) {
      yCells = new Y.Map();
      ySheet.set('cells', yCells);
    }

    for (const [key, cell] of sheet.cells) {
      const cellData = {};
      if (cell.rawValue !== null && cell.rawValue !== undefined) cellData.v = cell.rawValue;
      if (cell.formula) cellData.f = cell.formula;
      if (cell.style) cellData.s = cell.style.toJSON();
      if (cell.comment) cellData.c = cell.comment;

      yCells.set(String(key), cellData);
    }

    // Column widths
    let yColWidths = ySheet.get('colWidths');
    if (!yColWidths) { yColWidths = new Y.Map(); ySheet.set('colWidths', yColWidths); }
    for (const [col, w] of sheet.colWidths) {
      yColWidths.set(String(col), w);
    }

    // Row heights
    let yRowHeights = ySheet.get('rowHeights');
    if (!yRowHeights) { yRowHeights = new Y.Map(); ySheet.set('rowHeights', yRowHeights); }
    for (const [row, h] of sheet.rowHeights) {
      yRowHeights.set(String(row), h);
    }
  }

  /**
   * Push a single cell change to Y.Doc.
   * Called after any local edit.
   */
  pushCellChange(sheetId, row, col, cell) {
    if (this._suppressLocal || !this._initialized) return;

    this.doc.transact(() => {
      const ySheet = this.doc.getMap('sheets').get(sheetId);
      if (!ySheet) return;
      const yCells = ySheet.get('cells');
      if (!yCells) return;

      const key = String(cellKey(row, col));
      if (!cell || cell.isEmpty) {
        yCells.delete(key);
      } else {
        const cellData = {};
        if (cell.rawValue !== null && cell.rawValue !== undefined) cellData.v = cell.rawValue;
        if (cell.formula) cellData.f = cell.formula;
        if (cell.style) cellData.s = cell.style.toJSON();
        if (cell.comment) cellData.c = cell.comment;
        yCells.set(key, cellData);
      }
    });
  }

  /**
   * Push cursor/selection to awareness.
   */
  pushCursor(row, col, selectionRange) {
    if (!this.awareness) return;
    this.awareness.setLocalStateField('cursor', { row, col });
    this.awareness.setLocalStateField('selection', selectionRange ? selectionRange.toString() : null);
  }

  // ── Apply remote state from Y.Doc ──

  _applyFullState() {
    this._suppressLocal = true;
    const ss = this.spreadsheet;
    const sheetsMap = this.doc.getMap('sheets');
    const meta = this.doc.getMap('meta');

    const sheetOrder = meta.get('sheetOrder') || [];

    // Rebuild sheets from Y.Doc
    ss.sheets = [];
    for (const sheetId of sheetOrder) {
      const ySheet = sheetsMap.get(sheetId);
      if (!ySheet) continue;

      const sheet = ss.addSheet(ySheet.get('name') || 'Sheet');
      sheet.id = sheetId;
      sheet.colCount = ySheet.get('colCount') || 26;
      sheet.rowCount = ySheet.get('rowCount') || 1000;
      sheet.frozenRows = ySheet.get('frozenRows') || 0;
      sheet.frozenCols = ySheet.get('frozenCols') || 0;

      // Cells
      const yCells = ySheet.get('cells');
      if (yCells) {
        yCells.forEach((cellData, key) => {
          const k = parseInt(key);
          const row = k >> 16;
          const col = k & 0xffff;
          this._applyCellData(sheet, row, col, cellData);
        });
      }

      // Column widths
      const yColWidths = ySheet.get('colWidths');
      if (yColWidths) {
        yColWidths.forEach((w, col) => sheet.colWidths.set(parseInt(col), w));
      }

      // Row heights
      const yRowHeights = ySheet.get('rowHeights');
      if (yRowHeights) {
        yRowHeights.forEach((h, row) => sheet.rowHeights.set(parseInt(row), h));
      }
    }

    if (ss.sheets.length > 0) ss.activeSheetIndex = 0;
    if (ss.sheetTabs) ss.sheetTabs.update();
    ss.recalculate();
    ss.render();

    this._suppressLocal = false;
  }

  _applyCellData(sheet, row, col, cellData) {
    if (!cellData) { sheet.clearCell(row, col); return; }
    if (cellData.f) {
      sheet.setCellFormula(row, col, cellData.f);
    } else if (cellData.v !== undefined) {
      sheet.setCellValue(row, col, cellData.v);
    }
    if (cellData.s) {
      sheet.setCellStyle(row, col, cellData.s);
    }
    if (cellData.c) {
      const cell = sheet.getOrCreateCell(row, col);
      cell.comment = cellData.c;
    }
  }

  // ── Y.Doc observers (listen for remote changes) ──

  _setupObservers() {
    const sheetsMap = this.doc.getMap('sheets');

    // Deep observe all sheet changes
    sheetsMap.observeDeep((events) => {
      if (this._suppressRemote) return;

      this._suppressLocal = true;
      const ss = this.spreadsheet;

      for (const event of events) {
        // Cell-level changes
        if (event.path.length >= 2 && event.path[event.path.length - 1] === 'cells') {
          const sheetId = event.path[0];
          const sheet = ss.getSheetById(sheetId);
          if (!sheet) continue;

          // Apply each cell change
          event.changes.keys.forEach((change, key) => {
            const k = parseInt(key);
            const row = k >> 16;
            const col = k & 0xffff;

            if (change.action === 'delete') {
              sheet.clearCell(row, col);
            } else {
              const yCells = sheetsMap.get(sheetId)?.get('cells');
              if (yCells) {
                const cellData = yCells.get(key);
                if (cellData) this._applyCellData(sheet, row, col, cellData);
              }
            }
          });
        }

        // Sheet-level property changes (name, frozenRows, etc.)
        if (event.path.length === 1) {
          const sheetId = event.path[0];
          const sheet = ss.getSheetById(sheetId);
          const ySheet = sheetsMap.get(sheetId);
          if (sheet && ySheet) {
            if (ySheet.get('name') !== sheet.name) sheet.name = ySheet.get('name');
            if (ySheet.get('frozenRows') !== undefined) sheet.frozenRows = ySheet.get('frozenRows');
            if (ySheet.get('frozenCols') !== undefined) sheet.frozenCols = ySheet.get('frozenCols');
          }
        }
      }

      ss.recalculate();
      ss.render();
      if (ss.sheetTabs) ss.sheetTabs.update();

      this._suppressLocal = false;
    });
  }

  // ── Awareness (other users' cursors) ──

  _onAwarenessChange() {
    const states = this.awareness.getStates();
    const peers = [];

    states.forEach((state, clientId) => {
      if (clientId === this.doc.clientID) return; // skip self
      if (state.user) {
        peers.push({
          id: state.user.id,
          name: state.user.name,
          color: state.user.color,
          cursor: state.cursor || null,
          selection: state.selection || null,
          clientId,
        });
      }
    });

    this.spreadsheet.emit('awarenessChange', peers);
    this._remotePeers = peers;
  }

  /**
   * Get all connected peers (for rendering their cursors).
   */
  getPeers() {
    return this._remotePeers || [];
  }

  // ── Hook into Spreadsheet events ──

  /**
   * Call this after initializing to auto-sync all edits.
   */
  bindEvents() {
    const ss = this.spreadsheet;

    // Sync cell edits
    ss.on('edit', (e) => {
      if (this._suppressLocal) return;
      const sheet = e.sheet || ss.activeSheet;
      if (!sheet) return;
      const range = e.range;
      if (range) {
        const row = range.getRow() - 1;
        const col = range.getColumn() - 1;
        const cell = sheet.getCell(row, col);
        this.pushCellChange(sheet.id, row, col, cell);
      }
    });

    // Sync selection/cursor
    ss.on('selectionChanged', (sel) => {
      this.pushCursor(
        ss.selectionManager.activeRow,
        ss.selectionManager.activeCol,
        sel,
      );
    });

    // Sync structural changes
    ss.on('change', () => {
      if (this._suppressLocal) return;
      this.pushFullState(); // full sync on structural changes
    });

    // Sync sheet add/delete
    ss.on('sheetAdded', () => {
      if (this._suppressLocal) return;
      this.pushFullState();
    });
  }

  // ── Helpers ──

  _randomColor() {
    const colors = ['#4285f4', '#ea4335', '#fbbc04', '#34a853', '#ff6d01', '#46bdc6', '#9334e6', '#f538a0'];
    return colors[Math.floor(Math.random() * colors.length)];
  }

  destroy() {
    this.disconnect();
  }
}
