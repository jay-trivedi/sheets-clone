export default class CommandManager {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet;
    this.undoStack = [];
    this.redoStack = [];
    this.maxHistory = 100;
    this._batch = null;
  }

  execute(command) {
    if (this._batch) {
      this._batch.push(command);
      return;
    }
    this.undoStack.push([command]);
    if (this.undoStack.length > this.maxHistory) {
      this.undoStack.shift();
    }
    this.redoStack = [];
  }

  beginBatch() {
    this._batch = [];
  }

  pushToCurrentBatch(command) {
    if (this._batch) {
      this._batch.push(command);
    }
  }

  endBatch() {
    if (this._batch && this._batch.length > 0) {
      this.undoStack.push(this._batch);
      if (this.undoStack.length > this.maxHistory) {
        this.undoStack.shift();
      }
      this.redoStack = [];
    }
    this._batch = null;
  }

  undo() {
    const batch = this.undoStack.pop();
    if (!batch) return;

    const redoBatch = [];
    const sheet = this.spreadsheet.activeSheet;
    if (!sheet) return;

    for (let i = batch.length - 1; i >= 0; i--) {
      const cmd = batch[i];
      const redoCmd = this._undoCommand(cmd, sheet);
      if (redoCmd) redoBatch.unshift(redoCmd);
    }

    this.redoStack.push(redoBatch);
    this.spreadsheet.recalculate();
    this.spreadsheet.render();
  }

  redo() {
    const batch = this.redoStack.pop();
    if (!batch) return;

    const undoBatch = [];
    const sheet = this.spreadsheet.activeSheet;
    if (!sheet) return;

    for (const cmd of batch) {
      const undoCmd = this._redoCommand(cmd, sheet);
      if (undoCmd) undoBatch.push(undoCmd);
    }

    this.undoStack.push(undoBatch);
    this.spreadsheet.recalculate();
    this.spreadsheet.render();
  }

  _undoCommand(cmd, sheet) {
    switch (cmd.type) {
      case 'setCellValue': {
        const cell = sheet.getCell(cmd.row, cmd.col);
        const currentValue = cell ? cell.rawValue : null;
        const currentFormula = cell ? cell.formula : null;
        const currentStyle = cell && cell.style ? cell.style.clone() : null;

        if (cmd.oldFormula) {
          sheet.setCellFormula(cmd.row, cmd.col, cmd.oldFormula);
        } else if (cmd.oldValue !== null && cmd.oldValue !== undefined) {
          sheet.setCellValue(cmd.row, cmd.col, cmd.oldValue);
        } else {
          sheet.clearCell(cmd.row, cmd.col);
        }

        if (cmd.oldStyle) {
          sheet.setCellStyle(cmd.row, cmd.col, cmd.oldStyle);
        }

        return { ...cmd, oldValue: currentValue, oldFormula: currentFormula, oldStyle: currentStyle, newValue: cmd.oldValue };
      }

      case 'setStyle': {
        const cell = sheet.getCell(cmd.row, cmd.col);
        const current = cell && cell.style ? { ...cell.style } : {};

        for (const [key, val] of Object.entries(cmd.oldProps || {})) {
          sheet.setCellStyle(cmd.row, cmd.col, { [key]: val });
        }

        return { ...cmd, oldProps: cmd.newProps, newProps: cmd.oldProps };
      }

      case 'resize': {
        if (cmd.mode === 'resize-col') {
          sheet.setColWidth(cmd.index, cmd.oldSize);
        } else {
          sheet.setRowHeight(cmd.index, cmd.oldSize);
        }
        return { ...cmd, oldSize: cmd.newSize, newSize: cmd.oldSize };
      }

      default:
        return null;
    }
  }

  _redoCommand(cmd, sheet) {
    return this._undoCommand(cmd, sheet);
  }

  get canUndo() { return this.undoStack.length > 0; }
  get canRedo() { return this.redoStack.length > 0; }

  clear() {
    this.undoStack = [];
    this.redoStack = [];
  }
}
