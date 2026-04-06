import { el, indexToCol } from '../utils/helpers.js';
import CellRange from '../core/CellRange.js';

/**
 * Chart Manager — creates and manages Chart.js charts as overlays on the spreadsheet.
 * Charts are anchored to a cell position and rendered in a floating container.
 */
export default class ChartManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.charts = []; // { id, chart, container, anchorRow, anchorCol, width, height, sheetId }
    this._nextId = 1;
  }

  /**
   * Insert a chart from a data range.
   * @param {string} rangeStr - Data range in A1 notation (e.g., "A1:D10")
   * @param {string} type - Chart type: 'bar','line','pie','doughnut','scatter','radar','polarArea','bubble'
   * @param {object} options - { title, width, height, anchorRow, anchorCol, useFirstRowAsLabels, useFirstColAsLabels }
   */
  insert(rangeStr, type = 'bar', options = {}) {
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    if (!sheet) return null;

    const range = CellRange.fromString(rangeStr);
    if (!range) return null;

    const {
      title = '',
      width = 500,
      height = 300,
      anchorRow = range.endRow + 2,
      anchorCol = range.startCol,
      useFirstRowAsLabels = true,
      useFirstColAsLabels = true,
    } = options;

    // Extract data from range
    const rawData = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const row = [];
      for (let c = range.startCol; c <= range.endCol; c++) {
        row.push(sheet.getCellValue(r, c));
      }
      rawData.push(row);
    }

    // Parse into Chart.js format
    const { labels, datasets } = this._parseData(rawData, type, useFirstRowAsLabels, useFirstColAsLabels);

    // Create chart container
    const container = el('div', {
      className: 'sheets-chart-container',
      style: {
        position: 'absolute',
        zIndex: '50',
        background: '#fff',
        border: '1px solid #dadada',
        borderRadius: '8px',
        boxShadow: '0 2px 8px rgba(0,0,0,0.1)',
        padding: '8px',
        cursor: 'move',
        width: width + 'px',
        height: height + 'px',
      },
    });

    // Chart canvas
    const canvas = el('canvas', { width: width - 16, height: height - 40 });
    container.appendChild(canvas);

    // Close button
    const closeBtn = el('button', {
      className: 'sheets-chart-close',
      onClick: () => this.remove(chartObj.id),
    });
    closeBtn.innerHTML = '✕';
    container.appendChild(closeBtn);

    // Append to grid area
    ss.gridArea.appendChild(container);

    // Chart.js colors
    const COLORS = [
      '#4285f4', '#ea4335', '#fbbc04', '#34a853', '#ff6d01',
      '#46bdc6', '#7baaf7', '#f07b72', '#fdd663', '#57bb8a',
      '#e8710a', '#9334e6', '#f538a0', '#00897b', '#3949ab',
    ];

    const BG_COLORS = COLORS.map(c => c + '99'); // semi-transparent

    // Build Chart.js config
    const isPie = type === 'pie' || type === 'doughnut' || type === 'polarArea';

    const chartConfig = {
      type,
      data: {
        labels,
        datasets: datasets.map((ds, i) => ({
          label: ds.label,
          data: ds.data,
          backgroundColor: isPie ? BG_COLORS.slice(0, ds.data.length) : BG_COLORS[i % BG_COLORS.length],
          borderColor: isPie ? COLORS.slice(0, ds.data.length) : COLORS[i % COLORS.length],
          borderWidth: isPie ? 2 : 2,
          tension: type === 'line' ? 0.3 : 0,
          fill: type === 'line' ? false : undefined,
          pointRadius: type === 'scatter' ? 5 : type === 'line' ? 3 : undefined,
        })),
      },
      options: {
        responsive: false,
        maintainAspectRatio: false,
        plugins: {
          title: { display: !!title, text: title, font: { family: "'Google Sans', Roboto, Arial", size: 14 } },
          legend: { position: isPie ? 'right' : 'top', labels: { font: { family: "'Google Sans', Roboto, Arial", size: 11 } } },
        },
        scales: isPie ? {} : {
          x: { ticks: { font: { family: 'Arial', size: 11 } }, grid: { color: '#e3e3e3' } },
          y: { ticks: { font: { family: 'Arial', size: 11 } }, grid: { color: '#e3e3e3' } },
        },
      },
    };

    // Dynamic import Chart.js (it's bundled)
    let chart;
    try {
      const Chart = window.Chart || (typeof require !== 'undefined' ? require('chart.js/auto') : null);
      if (!Chart) {
        container.innerHTML = '<div style="padding:20px;color:#666;text-align:center">Chart.js not loaded.<br>Add &lt;script src="https://cdn.jsdelivr.net/npm/chart.js"&gt; to your page.</div>';
        const chartObj = { id: this._nextId++, chart: null, container, anchorRow, anchorCol, width, height, sheetId: sheet.id };
        this.charts.push(chartObj);
        this._positionChart(chartObj);
        return chartObj.id;
      }
      chart = new Chart(canvas.getContext('2d'), chartConfig);
    } catch (e) {
      container.innerHTML = `<div style="padding:20px;color:#d93025">${e.message}</div>`;
    }

    const chartObj = {
      id: this._nextId++,
      chart,
      container,
      anchorRow,
      anchorCol,
      width,
      height,
      sheetId: sheet.id,
      rangeStr,
      type,
      options,
    };

    this.charts.push(chartObj);
    this._positionChart(chartObj);
    this._makeDraggable(chartObj);

    return chartObj.id;
  }

  _parseData(rawData, type, useFirstRowAsLabels, useFirstColAsLabels) {
    if (rawData.length === 0) return { labels: [], datasets: [] };

    let labels = [];
    let datasets = [];

    if (useFirstRowAsLabels && useFirstColAsLabels) {
      // First row = series names, first col = category labels
      const seriesNames = rawData[0].slice(1);
      labels = rawData.slice(1).map(row => String(row[0] || ''));
      datasets = seriesNames.map((name, i) => ({
        label: String(name || 'Series ' + (i + 1)),
        data: rawData.slice(1).map(row => {
          const v = row[i + 1];
          return typeof v === 'number' ? v : parseFloat(v) || 0;
        }),
      }));
    } else if (useFirstRowAsLabels) {
      const headers = rawData[0];
      labels = headers.map(h => String(h || ''));
      datasets = [{
        label: 'Data',
        data: rawData.length > 1 ? rawData[1].map(v => typeof v === 'number' ? v : parseFloat(v) || 0) : [],
      }];
    } else {
      // Auto-detect: columns are series
      labels = rawData.map((_, i) => String(i + 1));
      for (let c = 0; c < (rawData[0] || []).length; c++) {
        datasets.push({
          label: 'Series ' + (c + 1),
          data: rawData.map(row => {
            const v = row[c];
            return typeof v === 'number' ? v : parseFloat(v) || 0;
          }),
        });
      }
    }

    return { labels, datasets };
  }

  _positionChart(chartObj) {
    const ss = this.spreadsheet;
    const sheet = ss.activeSheet;
    if (!sheet || sheet.id !== chartObj.sheetId) {
      chartObj.container.style.display = 'none';
      return;
    }
    chartObj.container.style.display = '';
    const rect = ss.renderer._getCellRect(sheet, chartObj.anchorRow, chartObj.anchorCol);
    chartObj.container.style.left = rect.x + 'px';
    chartObj.container.style.top = rect.y + 'px';
  }

  _makeDraggable(chartObj) {
    let startX, startY, startLeft, startTop;
    const container = chartObj.container;

    const onMouseDown = (e) => {
      if (e.target.tagName === 'BUTTON' || e.target.tagName === 'CANVAS') return;
      startX = e.clientX;
      startY = e.clientY;
      startLeft = parseInt(container.style.left) || 0;
      startTop = parseInt(container.style.top) || 0;
      document.addEventListener('mousemove', onMouseMove);
      document.addEventListener('mouseup', onMouseUp);
      e.preventDefault();
    };

    const onMouseMove = (e) => {
      container.style.left = (startLeft + e.clientX - startX) + 'px';
      container.style.top = (startTop + e.clientY - startY) + 'px';
    };

    const onMouseUp = () => {
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
    };

    container.addEventListener('mousedown', onMouseDown);
  }

  remove(id) {
    const idx = this.charts.findIndex(c => c.id === id);
    if (idx === -1) return;
    const chartObj = this.charts[idx];
    if (chartObj.chart) chartObj.chart.destroy();
    chartObj.container.remove();
    this.charts.splice(idx, 1);
  }

  updatePositions() {
    for (const chartObj of this.charts) {
      this._positionChart(chartObj);
    }
  }

  // Show chart insertion dialog
  showDialog() {
    const ss = this.spreadsheet;
    const sel = ss.selection;
    if (!sel) return;

    const overlay = el('div', { className: 'sheets-modal-overlay' });
    const dialog = el('div', { className: 'sheets-chart-dialog' });

    dialog.innerHTML = `
      <h3>Insert chart</h3>
      <div class="sheets-chart-field">
        <label>Data range</label>
        <input type="text" class="sheets-chart-range" value="${sel.toString()}" />
      </div>
      <div class="sheets-chart-field">
        <label>Chart type</label>
        <div class="sheets-chart-types">
          <button data-type="bar" class="active" title="Bar">📊</button>
          <button data-type="line" title="Line">📈</button>
          <button data-type="pie" title="Pie">🥧</button>
          <button data-type="doughnut" title="Doughnut">🍩</button>
          <button data-type="scatter" title="Scatter">⚬</button>
          <button data-type="radar" title="Radar">🕸</button>
          <button data-type="polarArea" title="Polar">🎯</button>
        </div>
      </div>
      <div class="sheets-chart-field">
        <label>Title</label>
        <input type="text" class="sheets-chart-title" placeholder="Chart title" />
      </div>
      <div class="sheets-chart-field">
        <label><input type="checkbox" class="sheets-chart-header-row" checked /> Use first row as labels</label>
      </div>
      <div class="sheets-chart-field">
        <label><input type="checkbox" class="sheets-chart-header-col" checked /> Use first column as labels</label>
      </div>
      <div class="sheets-chart-actions">
        <button class="sheets-chart-cancel">Cancel</button>
        <button class="sheets-chart-insert">Insert</button>
      </div>
    `;

    // Chart type selection
    let selectedType = 'bar';
    dialog.querySelectorAll('.sheets-chart-types button').forEach(btn => {
      btn.addEventListener('click', () => {
        dialog.querySelectorAll('.sheets-chart-types button').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        selectedType = btn.dataset.type;
      });
    });

    dialog.querySelector('.sheets-chart-cancel').addEventListener('click', () => overlay.remove());
    dialog.querySelector('.sheets-chart-insert').addEventListener('click', () => {
      const rangeStr = dialog.querySelector('.sheets-chart-range').value;
      const title = dialog.querySelector('.sheets-chart-title').value;
      const useFirstRowAsLabels = dialog.querySelector('.sheets-chart-header-row').checked;
      const useFirstColAsLabels = dialog.querySelector('.sheets-chart-header-col').checked;
      this.insert(rangeStr, selectedType, { title, useFirstRowAsLabels, useFirstColAsLabels });
      overlay.remove();
    });

    overlay.appendChild(dialog);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });
    document.body.appendChild(overlay);
  }

  destroy() {
    for (const c of this.charts) {
      if (c.chart) c.chart.destroy();
      c.container.remove();
    }
    this.charts = [];
  }
}
