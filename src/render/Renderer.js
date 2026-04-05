import {
  ROW_HEADER_WIDTH, COL_HEADER_HEIGHT, CELL_PADDING, CELL_BORDER_COLOR,
  HEADER_BG_COLOR, HEADER_BORDER_COLOR, HEADER_TEXT_COLOR,
  SELECTION_BORDER_COLOR, SELECTION_BG_COLOR, SELECTION_HEADER_BG,
  FILL_HANDLE_SIZE, FILL_HANDLE_COLOR, GRID_LINE_WIDTH,
  DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE, DEFAULT_TEXT_COLOR,
  FREEZE_LINE_COLOR, FREEZE_LINE_WIDTH, SCROLLBAR_SIZE,
  H_ALIGN, V_ALIGN, CELL_TYPE, DEFAULT_BG_COLOR,
} from '../utils/constants.js';
import { indexToCol, throttle, parseCellRef } from '../utils/helpers.js';

export default class Renderer {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.canvas = null;
    this.ctx = null;
    this.width = 0;
    this.height = 0;
    this.dpr = window.devicePixelRatio || 1;

    this.scrollX = 0;
    this.scrollY = 0;
    this.maxScrollX = 0;
    this.maxScrollY = 0;

    this._rafId = null;
    this._dirty = true;
  }

  init(canvas) {
    this.canvas = canvas;
    this.ctx = canvas.getContext('2d');
    this.resize();
  }

  resize() {
    const rect = this.canvas.parentElement.getBoundingClientRect();
    this.width = rect.width;
    this.height = rect.height;
    this.canvas.width = this.width * this.dpr;
    this.canvas.height = this.height * this.dpr;
    this.canvas.style.width = this.width + 'px';
    this.canvas.style.height = this.height + 'px';
    this.ctx.scale(this.dpr, this.dpr);
    this._updateMaxScroll();
    this.requestRender();
  }

  _updateMaxScroll() {
    const sheet = this.spreadsheet.activeSheet;
    if (!sheet) return;
    this.maxScrollX = Math.max(0, sheet.getTotalWidth() - (this.width - ROW_HEADER_WIDTH - SCROLLBAR_SIZE));
    this.maxScrollY = Math.max(0, sheet.getTotalHeight() - (this.height - COL_HEADER_HEIGHT - SCROLLBAR_SIZE));
  }

  scrollTo(x, y) {
    this._updateMaxScroll();
    this.scrollX = Math.max(0, Math.min(this.maxScrollX, x));
    this.scrollY = Math.max(0, Math.min(this.maxScrollY, y));
    this.requestRender();
  }

  scrollBy(dx, dy) {
    this.scrollTo(this.scrollX + dx, this.scrollY + dy);
  }

  ensureCellVisible(row, col) {
    const sheet = this.spreadsheet.activeSheet;
    if (!sheet) return;

    const cellX = sheet.getColX(col);
    const cellY = sheet.getRowY(row);
    const cellW = sheet.getColWidth(col);
    const cellH = sheet.getRowHeight(row);

    const viewLeft = this.scrollX;
    const viewTop = this.scrollY;
    const viewRight = this.scrollX + (this.width - ROW_HEADER_WIDTH - SCROLLBAR_SIZE);
    const viewBottom = this.scrollY + (this.height - COL_HEADER_HEIGHT - SCROLLBAR_SIZE);

    const frozenW = this._getFrozenWidth(sheet);
    const frozenH = this._getFrozenHeight(sheet);

    let sx = this.scrollX;
    let sy = this.scrollY;

    if (col >= sheet.frozenCols) {
      if (cellX < viewLeft + frozenW) sx = cellX - frozenW;
      else if (cellX + cellW > viewRight) sx = cellX + cellW - (this.width - ROW_HEADER_WIDTH - SCROLLBAR_SIZE);
    }
    if (row >= sheet.frozenRows) {
      if (cellY < viewTop + frozenH) sy = cellY - frozenH;
      else if (cellY + cellH > viewBottom) sy = cellY + cellH - (this.height - COL_HEADER_HEIGHT - SCROLLBAR_SIZE);
    }

    if (sx !== this.scrollX || sy !== this.scrollY) {
      this.scrollTo(sx, sy);
    }
  }

  requestRender() {
    this._dirty = true;
    if (!this._rafId) {
      this._rafId = requestAnimationFrame(() => {
        this._rafId = null;
        if (this._dirty) {
          this._dirty = false;
          this.render();
        }
      });
    }
  }

  render() {
    const ctx = this.ctx;
    const sheet = this.spreadsheet.activeSheet;
    if (!ctx || !sheet) return;

    ctx.save();
    ctx.clearRect(0, 0, this.width, this.height);

    const vp = this._getViewport(sheet);

    // Draw cells (main area)
    ctx.save();
    ctx.beginPath();
    ctx.rect(ROW_HEADER_WIDTH, COL_HEADER_HEIGHT,
      this.width - ROW_HEADER_WIDTH - SCROLLBAR_SIZE,
      this.height - COL_HEADER_HEIGHT - SCROLLBAR_SIZE);
    ctx.clip();
    this._drawCells(ctx, sheet, vp);
    this._drawGridLines(ctx, sheet, vp);
    this._drawFormulaRefs(ctx, sheet, vp);
    this._drawSelection(ctx, sheet, vp);
    this._drawFreezeLines(ctx, sheet);
    ctx.restore();

    // Draw headers
    this._drawColHeaders(ctx, sheet, vp);
    this._drawRowHeaders(ctx, sheet, vp);
    this._drawCorner(ctx);

    // Draw scrollbars
    this._drawScrollbars(ctx, sheet);

    ctx.restore();
  }

  _getViewport(sheet) {
    const viewW = this.width - ROW_HEADER_WIDTH - SCROLLBAR_SIZE;
    const viewH = this.height - COL_HEADER_HEIGHT - SCROLLBAR_SIZE;

    let startCol = 0, startRow = 0;
    let x = 0, y = 0;

    // Find first visible column
    for (let c = 0; c < sheet.colCount; c++) {
      if (!sheet.isColVisible(c)) continue;
      const w = sheet.getColWidth(c);
      if (x + w > this.scrollX) { startCol = c; break; }
      x += w;
    }

    // Find first visible row
    for (let r = 0; r < sheet.rowCount; r++) {
      if (!sheet.isRowVisible(r)) continue;
      const h = sheet.getRowHeight(r);
      if (y + h > this.scrollY) { startRow = r; break; }
      y += h;
    }

    // Find end col/row
    let endCol = startCol, endRow = startRow;
    let cx = sheet.getColX(startCol) - this.scrollX;
    for (let c = startCol; c < sheet.colCount; c++) {
      if (!sheet.isColVisible(c)) continue;
      if (cx > viewW) break;
      endCol = c;
      cx += sheet.getColWidth(c);
    }

    let cy = sheet.getRowY(startRow) - this.scrollY;
    for (let r = startRow; r < sheet.rowCount; r++) {
      if (!sheet.isRowVisible(r)) continue;
      if (cy > viewH) break;
      endRow = r;
      cy += sheet.getRowHeight(r);
    }

    return { startCol, startRow, endCol, endRow };
  }

  _getCellRect(sheet, row, col) {
    return {
      x: ROW_HEADER_WIDTH + sheet.getColX(col) - this.scrollX,
      y: COL_HEADER_HEIGHT + sheet.getRowY(row) - this.scrollY,
      w: sheet.getColWidth(col),
      h: sheet.getRowHeight(row),
    };
  }

  _getFrozenWidth(sheet) {
    let w = 0;
    for (let c = 0; c < sheet.frozenCols; c++) {
      if (sheet.isColVisible(c)) w += sheet.getColWidth(c);
    }
    return w;
  }

  _getFrozenHeight(sheet) {
    let h = 0;
    for (let r = 0; r < sheet.frozenRows; r++) {
      if (sheet.isRowVisible(r)) h += sheet.getRowHeight(r);
    }
    return h;
  }

  _drawCells(ctx, sheet, vp) {
    const numberFormat = this.spreadsheet.numberFormat;

    for (let r = vp.startRow; r <= vp.endRow; r++) {
      if (!sheet.isRowVisible(r)) continue;
      for (let c = vp.startCol; c <= vp.endCol; c++) {
        if (!sheet.isColVisible(c)) continue;

        const rect = this._getCellRect(sheet, r, c);
        const cell = sheet.getCell(r, c);

        // Skip cells that are part of a merge (not the top-left)
        if (cell && cell.mergeParent) continue;

        // Merge span
        let cellW = rect.w, cellH = rect.h;
        if (cell && cell.mergeSpan) {
          for (let mc = c + 1; mc < c + cell.mergeSpan.cols; mc++) {
            cellW += sheet.getColWidth(mc);
          }
          for (let mr = r + 1; mr < r + cell.mergeSpan.rows; mr++) {
            cellH += sheet.getRowHeight(mr);
          }
        }

        const style = cell ? cell.getStyle() : null;

        // Background
        const bg = style && style.bgColor ? style.bgColor : null;
        if (bg) {
          ctx.fillStyle = bg;
          ctx.fillRect(rect.x, rect.y, cellW, cellH);
        }

        // Conditional formatting
        const cfStyle = this.spreadsheet.getConditionalStyle(sheet, r, c);
        if (cfStyle) {
          if (cfStyle.bgColor) {
            ctx.fillStyle = cfStyle.bgColor;
            ctx.fillRect(rect.x, rect.y, cellW, cellH);
          }
        }

        // Cell text
        if (cell && !cell.isEmpty) {
          const effectiveStyle = style || cell.getStyle();
          const textColor = (cfStyle && cfStyle.textColor) || effectiveStyle.textColor || DEFAULT_TEXT_COLOR;
          const cellType = cell.effectiveType;

          let displayText = cell.displayValue;
          if (effectiveStyle.numberFormat && effectiveStyle.numberFormat !== 'General' && typeof cell.computedValue === 'number') {
            displayText = numberFormat.format(cell.computedValue, effectiveStyle.numberFormat);
          }

          ctx.fillStyle = textColor;
          ctx.font = effectiveStyle.getFont();
          ctx.textBaseline = 'middle';

          const hAlign = effectiveStyle.getEffectiveHAlign(cellType);
          const padding = CELL_PADDING + (effectiveStyle.indent * 8);

          let tx, maxW = cellW - padding * 2;
          if (hAlign === H_ALIGN.RIGHT) {
            ctx.textAlign = 'right';
            tx = rect.x + cellW - padding;
          } else if (hAlign === H_ALIGN.CENTER) {
            ctx.textAlign = 'center';
            tx = rect.x + cellW / 2;
          } else {
            ctx.textAlign = 'left';
            tx = rect.x + padding;
          }

          let ty;
          if (effectiveStyle.vAlign === V_ALIGN.TOP) {
            ty = rect.y + effectiveStyle.fontSize * 0.6 + 4;
          } else if (effectiveStyle.vAlign === V_ALIGN.MIDDLE) {
            ty = rect.y + cellH / 2;
          } else {
            ty = rect.y + cellH - effectiveStyle.fontSize * 0.4 - 4;
          }

          if (effectiveStyle.wrap) {
            this._drawWrappedText(ctx, displayText, tx, rect.y + CELL_PADDING, maxW, cellH - CELL_PADDING * 2, effectiveStyle);
          } else {
            // Clip text to cell
            ctx.save();
            ctx.beginPath();
            ctx.rect(rect.x + 1, rect.y + 1, cellW - 2, cellH - 2);
            ctx.clip();
            ctx.fillText(displayText, tx, ty);

            // Strikethrough
            if (effectiveStyle.strikethrough) {
              const metrics = ctx.measureText(displayText);
              const textW = metrics.width;
              const stX = hAlign === H_ALIGN.RIGHT ? tx - textW : hAlign === H_ALIGN.CENTER ? tx - textW / 2 : tx;
              ctx.beginPath();
              ctx.strokeStyle = textColor;
              ctx.lineWidth = 1;
              ctx.moveTo(stX, ty);
              ctx.lineTo(stX + textW, ty);
              ctx.stroke();
            }

            // Underline
            if (effectiveStyle.underline) {
              const metrics = ctx.measureText(displayText);
              const textW = metrics.width;
              const uX = hAlign === H_ALIGN.RIGHT ? tx - textW : hAlign === H_ALIGN.CENTER ? tx - textW / 2 : tx;
              ctx.beginPath();
              ctx.strokeStyle = textColor;
              ctx.lineWidth = 1;
              ctx.moveTo(uX, ty + effectiveStyle.fontSize * 0.5);
              ctx.lineTo(uX + textW, ty + effectiveStyle.fontSize * 0.5);
              ctx.stroke();
            }

            ctx.restore();
          }
        }

        // Borders
        if (style) {
          this._drawCellBorders(ctx, rect.x, rect.y, cellW, cellH, style);
        }
      }
    }
  }

  _drawWrappedText(ctx, text, x, y, maxW, maxH, style) {
    const lineHeight = style.fontSize * 1.4;
    const words = text.split(' ');
    let line = '';
    let ty = y + lineHeight;

    for (const word of words) {
      const testLine = line + (line ? ' ' : '') + word;
      const metrics = ctx.measureText(testLine);
      if (metrics.width > maxW && line) {
        ctx.fillText(line, x, ty);
        line = word;
        ty += lineHeight;
        if (ty > y + maxH) return;
      } else {
        line = testLine;
      }
    }
    if (line) ctx.fillText(line, x, ty);
  }

  _drawCellBorders(ctx, x, y, w, h, style) {
    ctx.lineWidth = 1;
    if (style.borderTop) {
      ctx.strokeStyle = style.borderTop.color || '#000';
      ctx.lineWidth = style.borderTop.width || 1;
      ctx.beginPath();
      ctx.moveTo(x, y + 0.5);
      ctx.lineTo(x + w, y + 0.5);
      ctx.stroke();
    }
    if (style.borderBottom) {
      ctx.strokeStyle = style.borderBottom.color || '#000';
      ctx.lineWidth = style.borderBottom.width || 1;
      ctx.beginPath();
      ctx.moveTo(x, y + h - 0.5);
      ctx.lineTo(x + w, y + h - 0.5);
      ctx.stroke();
    }
    if (style.borderLeft) {
      ctx.strokeStyle = style.borderLeft.color || '#000';
      ctx.lineWidth = style.borderLeft.width || 1;
      ctx.beginPath();
      ctx.moveTo(x + 0.5, y);
      ctx.lineTo(x + 0.5, y + h);
      ctx.stroke();
    }
    if (style.borderRight) {
      ctx.strokeStyle = style.borderRight.color || '#000';
      ctx.lineWidth = style.borderRight.width || 1;
      ctx.beginPath();
      ctx.moveTo(x + w - 0.5, y);
      ctx.lineTo(x + w - 0.5, y + h);
      ctx.stroke();
    }
  }

  _drawGridLines(ctx, sheet, vp) {
    ctx.strokeStyle = CELL_BORDER_COLOR;
    ctx.lineWidth = GRID_LINE_WIDTH;

    // Vertical lines
    for (let c = vp.startCol; c <= vp.endCol + 1; c++) {
      if (!sheet.isColVisible(c)) continue;
      const x = ROW_HEADER_WIDTH + sheet.getColX(c) - this.scrollX;
      ctx.beginPath();
      ctx.moveTo(Math.round(x) + 0.5, COL_HEADER_HEIGHT);
      ctx.lineTo(Math.round(x) + 0.5, this.height - SCROLLBAR_SIZE);
      ctx.stroke();
    }

    // Horizontal lines
    for (let r = vp.startRow; r <= vp.endRow + 1; r++) {
      if (!sheet.isRowVisible(r)) continue;
      const y = COL_HEADER_HEIGHT + sheet.getRowY(r) - this.scrollY;
      ctx.beginPath();
      ctx.moveTo(ROW_HEADER_WIDTH, Math.round(y) + 0.5);
      ctx.lineTo(this.width - SCROLLBAR_SIZE, Math.round(y) + 0.5);
      ctx.stroke();
    }
  }

  _drawSelection(ctx, sheet, vp) {
    const sel = this.spreadsheet.selection;
    if (!sel) return;

    // Multi-cell selection highlight
    if (!sel.isSingleCell) {
      for (let r = Math.max(sel.startRow, vp.startRow); r <= Math.min(sel.endRow, vp.endRow); r++) {
        for (let c = Math.max(sel.startCol, vp.startCol); c <= Math.min(sel.endCol, vp.endCol); c++) {
          if (r === this.spreadsheet.activeRow && c === this.spreadsheet.activeCol) continue;
          const rect = this._getCellRect(sheet, r, c);
          ctx.fillStyle = SELECTION_BG_COLOR;
          ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
        }
      }
    }

    // Selection border
    const selStartRect = this._getCellRect(sheet, sel.startRow, sel.startCol);
    const selEndRect = this._getCellRect(sheet, sel.endRow, sel.endCol);
    const sx = selStartRect.x;
    const sy = selStartRect.y;
    const sw = selEndRect.x + selEndRect.w - sx;
    const sh = selEndRect.y + selEndRect.h - sy;

    ctx.strokeStyle = SELECTION_BORDER_COLOR;
    ctx.lineWidth = 2;
    ctx.strokeRect(sx, sy, sw, sh);

    // Active cell border (white outline + blue border)
    const activeRect = this._getCellRect(sheet, this.spreadsheet.activeRow, this.spreadsheet.activeCol);
    ctx.strokeStyle = '#fff';
    ctx.lineWidth = 1;
    ctx.strokeRect(activeRect.x + 0.5, activeRect.y + 0.5, activeRect.w - 1, activeRect.h - 1);
    ctx.strokeStyle = SELECTION_BORDER_COLOR;
    ctx.lineWidth = 2;
    ctx.strokeRect(activeRect.x, activeRect.y, activeRect.w, activeRect.h);

    // Fill handle
    ctx.fillStyle = FILL_HANDLE_COLOR;
    const fhx = sx + sw - FILL_HANDLE_SIZE / 2;
    const fhy = sy + sh - FILL_HANDLE_SIZE / 2;
    ctx.fillRect(fhx, fhy, FILL_HANDLE_SIZE, FILL_HANDLE_SIZE);
    ctx.strokeStyle = '#fff';
    ctx.lineWidth = 1;
    ctx.strokeRect(fhx, fhy, FILL_HANDLE_SIZE, FILL_HANDLE_SIZE);

    // Copy indicator (marching ants)
    if (this.spreadsheet.copyRange) {
      const cr = this.spreadsheet.copyRange;
      const crStart = this._getCellRect(sheet, cr.startRow, cr.startCol);
      const crEnd = this._getCellRect(sheet, cr.endRow, cr.endCol);
      ctx.setLineDash([5, 3]);
      ctx.strokeStyle = SELECTION_BORDER_COLOR;
      ctx.lineWidth = 2;
      ctx.strokeRect(crStart.x, crStart.y,
        crEnd.x + crEnd.w - crStart.x,
        crEnd.y + crEnd.h - crStart.y);
      ctx.setLineDash([]);
    }
  }

  _drawFormulaRefs(ctx, sheet, vp) {
    const fh = this.spreadsheet.formulaHelper;
    if (!fh || !fh.active) return;

    const refs = fh.getFormulaRefs();
    for (const ref of refs) {
      const startParsed = parseCellRef(ref.startRef);
      if (!startParsed) continue;

      let r1 = startParsed.row, c1 = startParsed.col;
      let r2 = r1, c2 = c1;

      if (ref.endRef) {
        const endParsed = parseCellRef(ref.endRef);
        if (endParsed) {
          r2 = endParsed.row;
          c2 = endParsed.col;
        }
      }

      // Normalize
      if (r1 > r2) { const t = r1; r1 = r2; r2 = t; }
      if (c1 > c2) { const t = c1; c1 = c2; c2 = t; }

      const startRect = this._getCellRect(sheet, r1, c1);
      const endRect = this._getCellRect(sheet, r2, c2);
      const x = startRect.x;
      const y = startRect.y;
      const w = endRect.x + endRect.w - x;
      const h = endRect.y + endRect.h - y;

      // Fill with semi-transparent color
      ctx.fillStyle = ref.color + '18';
      ctx.fillRect(x, y, w, h);

      // Border
      ctx.strokeStyle = ref.color;
      ctx.lineWidth = 2;
      ctx.strokeRect(x, y, w, h);
    }
  }

  _drawFreezeLines(ctx, sheet) {
    if (sheet.frozenRows > 0) {
      const y = COL_HEADER_HEIGHT + this._getFrozenHeight(sheet);
      ctx.strokeStyle = FREEZE_LINE_COLOR;
      ctx.lineWidth = FREEZE_LINE_WIDTH;
      ctx.beginPath();
      ctx.moveTo(ROW_HEADER_WIDTH, y);
      ctx.lineTo(this.width - SCROLLBAR_SIZE, y);
      ctx.stroke();
    }
    if (sheet.frozenCols > 0) {
      const x = ROW_HEADER_WIDTH + this._getFrozenWidth(sheet);
      ctx.strokeStyle = FREEZE_LINE_COLOR;
      ctx.lineWidth = FREEZE_LINE_WIDTH;
      ctx.beginPath();
      ctx.moveTo(x, COL_HEADER_HEIGHT);
      ctx.lineTo(x, this.height - SCROLLBAR_SIZE);
      ctx.stroke();
    }
  }

  _drawColHeaders(ctx, sheet, vp) {
    const sel = this.spreadsheet.selection;

    ctx.fillStyle = HEADER_BG_COLOR;
    ctx.fillRect(ROW_HEADER_WIDTH, 0, this.width - ROW_HEADER_WIDTH, COL_HEADER_HEIGHT);

    ctx.font = `bold 11px ${DEFAULT_FONT_FAMILY}`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';

    for (let c = vp.startCol; c <= vp.endCol; c++) {
      if (!sheet.isColVisible(c)) continue;
      const x = ROW_HEADER_WIDTH + sheet.getColX(c) - this.scrollX;
      const w = sheet.getColWidth(c);

      // Highlight if in selection
      if (sel && c >= sel.startCol && c <= sel.endCol) {
        ctx.fillStyle = SELECTION_HEADER_BG;
        ctx.fillRect(x, 0, w, COL_HEADER_HEIGHT);
        ctx.fillStyle = SELECTION_BORDER_COLOR;
      } else {
        ctx.fillStyle = HEADER_TEXT_COLOR;
      }

      ctx.fillText(indexToCol(c), x + w / 2, COL_HEADER_HEIGHT / 2);

      // Border
      ctx.strokeStyle = HEADER_BORDER_COLOR;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(Math.round(x + w) + 0.5, 0);
      ctx.lineTo(Math.round(x + w) + 0.5, COL_HEADER_HEIGHT);
      ctx.stroke();
    }

    // Bottom border
    ctx.strokeStyle = HEADER_BORDER_COLOR;
    ctx.beginPath();
    ctx.moveTo(ROW_HEADER_WIDTH, COL_HEADER_HEIGHT - 0.5);
    ctx.lineTo(this.width, COL_HEADER_HEIGHT - 0.5);
    ctx.stroke();
  }

  _drawRowHeaders(ctx, sheet, vp) {
    const sel = this.spreadsheet.selection;

    ctx.fillStyle = HEADER_BG_COLOR;
    ctx.fillRect(0, COL_HEADER_HEIGHT, ROW_HEADER_WIDTH, this.height - COL_HEADER_HEIGHT);

    ctx.font = `11px ${DEFAULT_FONT_FAMILY}`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';

    for (let r = vp.startRow; r <= vp.endRow; r++) {
      if (!sheet.isRowVisible(r)) continue;
      const y = COL_HEADER_HEIGHT + sheet.getRowY(r) - this.scrollY;
      const h = sheet.getRowHeight(r);

      if (sel && r >= sel.startRow && r <= sel.endRow) {
        ctx.fillStyle = SELECTION_HEADER_BG;
        ctx.fillRect(0, y, ROW_HEADER_WIDTH, h);
        ctx.fillStyle = SELECTION_BORDER_COLOR;
      } else {
        ctx.fillStyle = HEADER_TEXT_COLOR;
      }

      ctx.fillText(String(r + 1), ROW_HEADER_WIDTH / 2, y + h / 2);

      // Border
      ctx.strokeStyle = HEADER_BORDER_COLOR;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(0, Math.round(y + h) + 0.5);
      ctx.lineTo(ROW_HEADER_WIDTH, Math.round(y + h) + 0.5);
      ctx.stroke();
    }

    // Right border
    ctx.strokeStyle = HEADER_BORDER_COLOR;
    ctx.beginPath();
    ctx.moveTo(ROW_HEADER_WIDTH - 0.5, COL_HEADER_HEIGHT);
    ctx.lineTo(ROW_HEADER_WIDTH - 0.5, this.height);
    ctx.stroke();
  }

  _drawCorner(ctx) {
    ctx.fillStyle = HEADER_BG_COLOR;
    ctx.fillRect(0, 0, ROW_HEADER_WIDTH, COL_HEADER_HEIGHT);
    ctx.strokeStyle = HEADER_BORDER_COLOR;
    ctx.lineWidth = 1;
    ctx.strokeRect(0, 0, ROW_HEADER_WIDTH, COL_HEADER_HEIGHT);
  }

  _drawScrollbars(ctx, sheet) {
    const trackColor = '#f1f1f1';
    const thumbColor = '#c1c1c1';
    const thumbHover = '#a8a8a8';

    // Horizontal scrollbar
    const hTrackX = ROW_HEADER_WIDTH;
    const hTrackY = this.height - SCROLLBAR_SIZE;
    const hTrackW = this.width - ROW_HEADER_WIDTH - SCROLLBAR_SIZE;

    ctx.fillStyle = trackColor;
    ctx.fillRect(hTrackX, hTrackY, hTrackW, SCROLLBAR_SIZE);

    if (this.maxScrollX > 0) {
      const ratio = hTrackW / (this.maxScrollX + hTrackW);
      const thumbW = Math.max(30, hTrackW * ratio);
      const thumbX = hTrackX + (this.scrollX / this.maxScrollX) * (hTrackW - thumbW);
      ctx.fillStyle = thumbColor;
      ctx.beginPath();
      const r = 3;
      this._roundRect(ctx, thumbX + 2, hTrackY + 2, thumbW - 4, SCROLLBAR_SIZE - 4, r);
      ctx.fill();
    }

    // Vertical scrollbar
    const vTrackX = this.width - SCROLLBAR_SIZE;
    const vTrackY = COL_HEADER_HEIGHT;
    const vTrackH = this.height - COL_HEADER_HEIGHT - SCROLLBAR_SIZE;

    ctx.fillStyle = trackColor;
    ctx.fillRect(vTrackX, vTrackY, SCROLLBAR_SIZE, vTrackH);

    if (this.maxScrollY > 0) {
      const ratio = vTrackH / (this.maxScrollY + vTrackH);
      const thumbH = Math.max(30, vTrackH * ratio);
      const thumbY = vTrackY + (this.scrollY / this.maxScrollY) * (vTrackH - thumbH);
      ctx.fillStyle = thumbColor;
      ctx.beginPath();
      const r = 3;
      this._roundRect(ctx, vTrackX + 2, thumbY + 2, SCROLLBAR_SIZE - 4, thumbH - 4, r);
      ctx.fill();
    }

    // Corner between scrollbars
    ctx.fillStyle = trackColor;
    ctx.fillRect(this.width - SCROLLBAR_SIZE, this.height - SCROLLBAR_SIZE, SCROLLBAR_SIZE, SCROLLBAR_SIZE);
  }

  _roundRect(ctx, x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.lineTo(x + w - r, y);
    ctx.quadraticCurveTo(x + w, y, x + w, y + r);
    ctx.lineTo(x + w, y + h - r);
    ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
    ctx.lineTo(x + r, y + h);
    ctx.quadraticCurveTo(x, y + h, x, y + h - r);
    ctx.lineTo(x, y + r);
    ctx.quadraticCurveTo(x, y, x + r, y);
    ctx.closePath();
  }

  // Hit testing
  getCellAtPoint(x, y) {
    const sheet = this.spreadsheet.activeSheet;
    if (!sheet) return null;

    if (x < ROW_HEADER_WIDTH || y < COL_HEADER_HEIGHT) return null;

    const gridX = x - ROW_HEADER_WIDTH + this.scrollX;
    const gridY = y - COL_HEADER_HEIGHT + this.scrollY;

    const col = sheet.getColAtX(gridX);
    const row = sheet.getRowAtY(gridY);

    return { row, col };
  }

  getHeaderAtPoint(x, y) {
    const sheet = this.spreadsheet.activeSheet;
    if (!sheet) return null;

    if (y < COL_HEADER_HEIGHT && x > ROW_HEADER_WIDTH) {
      const gridX = x - ROW_HEADER_WIDTH + this.scrollX;
      const col = sheet.getColAtX(gridX);
      // Check if near edge for resize
      const colRight = sheet.getColX(col) + sheet.getColWidth(col);
      const isResize = Math.abs(gridX - colRight) < 4;
      return { type: 'col', index: col, isResize };
    }

    if (x < ROW_HEADER_WIDTH && y > COL_HEADER_HEIGHT) {
      const gridY = y - COL_HEADER_HEIGHT + this.scrollY;
      const row = sheet.getRowAtY(gridY);
      const rowBottom = sheet.getRowY(row) + sheet.getRowHeight(row);
      const isResize = Math.abs(gridY - rowBottom) < 4;
      return { type: 'row', index: row, isResize };
    }

    if (x < ROW_HEADER_WIDTH && y < COL_HEADER_HEIGHT) {
      return { type: 'corner' };
    }

    return null;
  }

  isFillHandle(x, y) {
    const sel = this.spreadsheet.selection;
    if (!sel) return false;
    const sheet = this.spreadsheet.activeSheet;
    const endRect = this._getCellRect(sheet, sel.endRow, sel.endCol);
    const fhx = endRect.x + endRect.w;
    const fhy = endRect.y + endRect.h;
    return Math.abs(x - fhx) < FILL_HANDLE_SIZE && Math.abs(y - fhy) < FILL_HANDLE_SIZE;
  }

  isScrollbar(x, y) {
    if (x > this.width - SCROLLBAR_SIZE && y > COL_HEADER_HEIGHT) return 'vertical';
    if (y > this.height - SCROLLBAR_SIZE && x > ROW_HEADER_WIDTH) return 'horizontal';
    return null;
  }

  destroy() {
    if (this._rafId) cancelAnimationFrame(this._rafId);
  }
}
