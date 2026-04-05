# Google Sheets Interaction Reference

## Cell Modes State Machine

### READY (default)
- Arrow keys navigate between cells
- Typing a printable character → ENTER mode (clears cell, starts fresh)
- F2 → EDIT mode (cursor at end of existing content)
- Double-click → EDIT mode (cursor at click position)
- Click formula bar → EDIT mode

### ENTER (new input / overwrite)
- Cell content cleared, typing replaces it
- If content starts with `=`: formula mode — arrow keys → POINT mode
- If non-formula: arrow keys commit + navigate (→ READY)
- Enter: commit + move down
- Tab: commit + move right
- Escape: discard → READY
- F2: switch to EDIT mode

### EDIT (in-cell cursor editing)
- Shows existing content with blinking cursor
- Arrow Left/Right: move text cursor (NOT navigate grid)
- Arrow Up: cursor to start
- Arrow Down: cursor to end
- F2: toggle to POINT mode (if formula)
- Enter: commit + move down
- Escape: discard → READY

### POINT (formula reference selection)
- Arrow keys navigate grid and INSERT cell references
- Shift+Arrow: extend to range
- Click cell: insert reference
- Typing operator (+,-,*,/,(,,): finalize ref → ENTER
- F2: toggle back to EDIT
- Enter: commit
- Escape: discard

## Point Mode Triggers
After these characters, arrow keys/clicks insert refs:
`= + - * / ^ ( , ; < > <= >= <> ! &`

## Enter / Shift+Enter behavior
| Context | Enter | Shift+Enter |
|---------|-------|-------------|
| READY mode | Move down | Move up |
| ENTER/EDIT mode | Commit + move down | Commit + move up |
| POINT mode | Commit + move down | Commit + move up |
| Formula bar | Same as cell editing | Same as cell editing |
| After Tab-navigation | Commit + move to tab-start column, next row | Commit + move up (keep column) |

Note: Alt+Enter inserts a newline (NOT Shift+Enter). Ctrl+Enter commits but stays in same cell.

## Key Behaviors
- Backspace in READY: clear cell + enter EDIT mode
- Ctrl+Enter: commit, stay in same cell
- Alt+Enter: insert newline in cell
- Tab tracking: Enter after Tab-navigation returns to starting column
- Delete: clear contents only (preserve formatting)
- F4: cycle A1 → $A$1 → A$1 → $A1 → A1
