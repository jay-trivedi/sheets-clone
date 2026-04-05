import CellRange from '../core/CellRange.js';
import { ROW_HEADER_WIDTH, COL_HEADER_HEIGHT, SCROLLBAR_SIZE, DOUBLE_CLICK_MS,
  AUTOSCROLL_SPEED, AUTOSCROLL_INTERVAL } from '../utils/constants.js';

export default class MouseHandler {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this._mode = null; // 'select', 'resize-col', 'resize-row', 'fill', 'scrollbar-h', 'scrollbar-v'
    this._startX = 0;
    this._startY = 0;
    this._resizeIndex = 0;
    this._resizeStart = 0;
    this._lastClickTime = 0;
    this._lastClickRow = -1;
    this._lastClickCol = -1;
    this._autoscrollTimer = null;
    this._fillStart = null;

    this._onMouseDown = this._onMouseDown.bind(this);
    this._onMouseMove = this._onMouseMove.bind(this);
    this._onMouseUp = this._onMouseUp.bind(this);
    this._onWheel = this._onWheel.bind(this);
    this._onDblClick = this._onDblClick.bind(this);
    this._onContextMenu = this._onContextMenu.bind(this);
  }

  init(canvas) {
    this.canvas = canvas;
    canvas.addEventListener('mousedown', this._onMouseDown);
    canvas.addEventListener('mousemove', this._onMouseMoveIdle = this._onMouseMoveIdle.bind(this));
    canvas.addEventListener('wheel', this._onWheel, { passive: false });
    canvas.addEventListener('dblclick', this._onDblClick);
    canvas.addEventListener('contextmenu', this._onContextMenu);
  }

  destroy() {
    this.canvas.removeEventListener('mousedown', this._onMouseDown);
    this.canvas.removeEventListener('mousemove', this._onMouseMoveIdle);
    this.canvas.removeEventListener('wheel', this._onWheel);
    this.canvas.removeEventListener('dblclick', this._onDblClick);
    this.canvas.removeEventListener('contextmenu', this._onContextMenu);
    this._stopAutoscroll();
  }

  _getXY(e) {
    const rect = this.canvas.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  }

  _onMouseMoveIdle(e) {
    if (this._mode) return;
    const { x, y } = this._getXY(e);
    const renderer = this.spreadsheet.renderer;

    // Cursor changes
    const header = renderer.getHeaderAtPoint(x, y);
    if (header && header.isResize) {
      this.canvas.style.cursor = header.type === 'col' ? 'col-resize' : 'row-resize';
    } else if (renderer.isFillHandle(x, y)) {
      this.canvas.style.cursor = 'crosshair';
    } else if (x < ROW_HEADER_WIDTH && y < COL_HEADER_HEIGHT) {
      this.canvas.style.cursor = 'default';
    } else if (x < ROW_HEADER_WIDTH || y < COL_HEADER_HEIGHT) {
      this.canvas.style.cursor = 'default';
    } else {
      this.canvas.style.cursor = 'cell';
    }
  }

  _onMouseDown(e) {
    if (e.button !== 0 && e.button !== 2) return;
    const { x, y } = this._getXY(e);
    const renderer = this.spreadsheet.renderer;
    const ss = this.spreadsheet;

    // Data validation dropdown
    const dvCell = renderer.getCellAtPoint(x, y);
    if (dvCell && x > 0) {
      const sheet = ss.activeSheet;
      if (sheet) {
        const cellData = sheet.getCell(dvCell.row, dvCell.col);
        if (cellData && cellData.validation && cellData.validation.type === 'list' && cellData.validation.showDropdown) {
          const cellRect = renderer._getCellRect(sheet, dvCell.row, dvCell.col);
          if (x >= cellRect.x + cellRect.w - 18) {
            ss.dataValidation.showDropdown(dvCell.row, dvCell.col);
            return;
          }
        }
      }
    }

    // Format painter
    if (ss._formatPainterStyle) {
      const cell = renderer.getCellAtPoint(x, y);
      if (cell) {
        ss.applyFormatPainter(ss.selection || { forEach: (fn) => fn(cell.row, cell.col) });
        return;
      }
    }

    // Filter arrow click
    const filterCol = renderer.isFilterArrow(x, y);
    if (filterCol !== null) {
      ss.showFilterDropdown(filterCol);
      return;
    }

    // Formula point-mode: clicking a cell inserts a reference instead of closing editor
    if (ss.editor && ss.editor.isActive && ss.editor.isFormulaMode && ss.formulaHelper) {
      const cell = renderer.getCellAtPoint(x, y);
      if (cell) {
        const handled = ss.formulaHelper.handlePointModeClick(cell.row, cell.col, e.shiftKey);
        if (handled) {
          e.preventDefault(); // Prevent browser from stealing focus from textarea
          return;
        }
      }
    }

    // Close editor if clicking outside
    if (ss.editor && ss.editor.isActive) {
      ss.editor.commit();
    }

    // Scrollbar
    const scrollbar = renderer.isScrollbar(x, y);
    if (scrollbar === 'horizontal') {
      this._mode = 'scrollbar-h';
      this._startX = x;
      this._scrollStart = renderer.scrollX;
      this._addDragListeners();
      return;
    }
    if (scrollbar === 'vertical') {
      this._mode = 'scrollbar-v';
      this._startY = y;
      this._scrollStart = renderer.scrollY;
      this._addDragListeners();
      return;
    }

    // Header resize
    const header = renderer.getHeaderAtPoint(x, y);
    if (header && header.isResize) {
      if (header.type === 'col') {
        this._mode = 'resize-col';
        this._resizeIndex = header.index;
        this._startX = x;
        this._resizeStart = ss.activeSheet.getColWidth(header.index);
      } else {
        this._mode = 'resize-row';
        this._resizeIndex = header.index;
        this._startY = y;
        this._resizeStart = ss.activeSheet.getRowHeight(header.index);
      }
      this._addDragListeners();
      return;
    }

    // Header select
    if (header && !header.isResize) {
      if (header.type === 'col') {
        ss.selectionManager.selectCol(header.index);
      } else if (header.type === 'row') {
        ss.selectionManager.selectRow(header.index);
      } else if (header.type === 'corner') {
        ss.selectionManager.selectAll();
      }
      ss.render();
      return;
    }

    // Fill handle
    if (renderer.isFillHandle(x, y)) {
      this._mode = 'fill';
      this._fillStart = ss.selection ? { ...ss.selection } : null;
      this._addDragListeners();
      return;
    }

    // Cell selection
    const cell = renderer.getCellAtPoint(x, y);
    if (cell) {
      if (e.button === 2) {
        // Right-click: don't change selection if clicking within it
        const sel = ss.selection;
        if (sel && sel.contains(cell.row, cell.col)) {
          return; // Context menu will fire
        }
      }

      this._mode = 'select';
      if (e.shiftKey) {
        ss.selectionManager.extendTo(cell.row, cell.col);
      } else {
        ss.selectionManager.select(cell.row, cell.col);
      }
      this._addDragListeners();
      this._startAutoscroll();
      ss.render();
    }
  }

  _onMouseMove(e) {
    const { x, y } = this._getXY(e);
    const ss = this.spreadsheet;
    const renderer = ss.renderer;

    switch (this._mode) {
      case 'select': {
        // Auto-scroll when dragging near/beyond edges
        const margin = 20;
        const viewRight = renderer.width - SCROLLBAR_SIZE;
        const viewBottom = renderer.height - SCROLLBAR_SIZE;
        let scrollDx = 0, scrollDy = 0;

        if (x > viewRight - margin) scrollDx = Math.min(30, (x - viewRight + margin) * 0.5);
        else if (x < ROW_HEADER_WIDTH + margin) scrollDx = -Math.min(30, (ROW_HEADER_WIDTH + margin - x) * 0.5);
        if (y > viewBottom - margin) scrollDy = Math.min(30, (y - viewBottom + margin) * 0.5);
        else if (y < COL_HEADER_HEIGHT + margin) scrollDy = -Math.min(30, (COL_HEADER_HEIGHT + margin - y) * 0.5);

        if (scrollDx || scrollDy) renderer.scrollBy(scrollDx, scrollDy);

        // Get cell at mouse position (allow edge clamping for beyond-viewport drags)
        const clampedX = Math.max(ROW_HEADER_WIDTH + 1, x);
        const clampedY = Math.max(COL_HEADER_HEIGHT + 1, y);
        const cell = renderer.getCellAtPoint(clampedX, clampedY);
        if (cell) {
          ss.selectionManager.extendTo(cell.row, cell.col);
          ss.render();
        }
        break;
      }

      case 'resize-col': {
        const delta = x - this._startX;
        const newWidth = Math.max(20, this._resizeStart + delta);
        ss.activeSheet.setColWidth(this._resizeIndex, newWidth);
        ss.render();
        break;
      }

      case 'resize-row': {
        const delta = y - this._startY;
        const newHeight = Math.max(10, this._resizeStart + delta);
        ss.activeSheet.setRowHeight(this._resizeIndex, newHeight);
        ss.render();
        break;
      }

      case 'fill': {
        const cell = renderer.getCellAtPoint(x, y);
        if (cell && this._fillStart) {
          const sel = ss.selection;
          // Extend in one direction only
          const dr = cell.row - this._fillStart.endRow;
          const dc = cell.col - this._fillStart.endCol;
          if (Math.abs(dr) >= Math.abs(dc)) {
            ss.selectionManager.selection = new CellRange(
              this._fillStart.startRow, this._fillStart.startCol,
              cell.row, this._fillStart.endCol,
            );
          } else {
            ss.selectionManager.selection = new CellRange(
              this._fillStart.startRow, this._fillStart.startCol,
              this._fillStart.endRow, cell.col,
            );
          }
          ss.render();
        }
        break;
      }

      case 'scrollbar-h': {
        const trackW = renderer.width - ROW_HEADER_WIDTH - SCROLLBAR_SIZE;
        const dx = x - this._startX;
        const scrollDelta = (dx / trackW) * (renderer.maxScrollX + trackW);
        renderer.scrollTo(this._scrollStart + scrollDelta, renderer.scrollY);
        break;
      }

      case 'scrollbar-v': {
        const trackH = renderer.height - COL_HEADER_HEIGHT - SCROLLBAR_SIZE;
        const dy = y - this._startY;
        const scrollDelta = (dy / trackH) * (renderer.maxScrollY + trackH);
        renderer.scrollTo(renderer.scrollX, this._scrollStart + scrollDelta);
        break;
      }
    }
  }

  _onMouseUp(e) {
    const ss = this.spreadsheet;

    if (this._mode === 'fill' && this._fillStart) {
      ss.autoFill(this._fillStart, ss.selection);
    }

    if (this._mode === 'resize-col' || this._mode === 'resize-row') {
      ss.commandManager.execute({
        type: 'resize',
        mode: this._mode,
        index: this._resizeIndex,
        oldSize: this._resizeStart,
        newSize: this._mode === 'resize-col'
          ? ss.activeSheet.getColWidth(this._resizeIndex)
          : ss.activeSheet.getRowHeight(this._resizeIndex),
      });
    }

    this._mode = null;
    this._removeDragListeners();
    this._stopAutoscroll();
    this.canvas.style.cursor = 'cell';
  }

  _onWheel(e) {
    e.preventDefault();
    const renderer = this.spreadsheet.renderer;
    const dx = e.deltaX || 0;
    const dy = e.deltaY || 0;
    renderer.scrollBy(dx, dy);
  }

  _onDblClick(e) {
    const { x, y } = this._getXY(e);
    const renderer = this.spreadsheet.renderer;
    const ss = this.spreadsheet;

    // Double-click header border to auto-fit
    const header = renderer.getHeaderAtPoint(x, y);
    if (header && header.isResize) {
      if (header.type === 'col') {
        ss.activeSheet.autoFitColWidth(header.index, renderer.ctx);
        ss.render();
      }
      return;
    }

    // Double-click cell to edit
    const cell = renderer.getCellAtPoint(x, y);
    if (cell && ss.editor) {
      const cellData = ss.activeSheet.getCell(cell.row, cell.col);
      const val = cellData ? (cellData.formula || cellData.displayValue) : '';
      ss.editor.begin(val, true);
    }
  }

  _onContextMenu(e) {
    e.preventDefault();
    const { x, y } = this._getXY(e);
    const ss = this.spreadsheet;
    if (ss.contextMenu) {
      const cell = ss.renderer.getCellAtPoint(x, y);
      const header = ss.renderer.getHeaderAtPoint(x, y);
      ss.contextMenu.show(e.clientX, e.clientY, cell, header);
    }
  }

  _addDragListeners() {
    document.addEventListener('mousemove', this._onMouseMove);
    document.addEventListener('mouseup', this._onMouseUp);
  }

  _removeDragListeners() {
    document.removeEventListener('mousemove', this._onMouseMove);
    document.removeEventListener('mouseup', this._onMouseUp);
  }

  _startAutoscroll() {
    // Auto-scroll when dragging near edges
    this._autoscrollTimer = setInterval(() => {
      if (this._mode !== 'select') return;
      // Check if mouse is near edges and scroll
    }, AUTOSCROLL_INTERVAL);
  }

  _stopAutoscroll() {
    if (this._autoscrollTimer) {
      clearInterval(this._autoscrollTimer);
      this._autoscrollTimer = null;
    }
  }
}
