# Google Sheets Feature Audit

Comprehensive extraction of every feature, UI element, user interaction, and capability
mentioned in "The Ultimate Guide to Google Sheets" by Zapier (2016).

Source: `/home/ubuntu/sheets-clone/0913-the-ultimate-guide-to-google-sheets.pdf` (186 pages)

---

## 1. SPREADSHEET STRUCTURE AND TERMINOLOGY

### Core Concepts
- **Cell**: A single data point or element; identified by column letter + row number (e.g., A1, B7)
- **Column**: A vertical set of cells, labeled with letters (A, B, C, ... Z, AA, AB, ...)
- **Row**: A horizontal set of cells, labeled with numbers (1, 2, 3, ...)
- **Range**: A selection of cells extending across a row, column, or both (e.g., A1:B10, A1:A, B2:H)
- **Function**: A built-in operation (SUM, AVERAGE, COUNT, etc.)
- **Formula**: A combination of functions, cells, rows, columns, and ranges to compute a result
- **Worksheet (Sheet)**: A named set of rows and columns; one spreadsheet can have multiple sheets
- **Spreadsheet**: The entire document containing all worksheets
- **Named Range**: A user-defined alias for a range (e.g., "name" instead of A:A), case-sensitive

### Grid Dimensions
- Columns labeled A through Z, then AA, AB, etc.
- Rows labeled with numbers starting at 1
- Grid occupies most of the screen area (white-and-grey grid)
- Blue outline around selected cell(s)

---

## 2. APPLICATION INTERFACE (UI ELEMENTS)

### 2.1 Top-Level Layout
- **Spreadsheet title** (editable, top left) with star icon (favorite) and folder icon (move to folder)
- **Menu bar**: File, Edit, View, Insert, Format, Data, Tools, Add-ons, Help
- **Toolbar** (icon row below menu bar)
- **Formula bar** (fx label + input field, below toolbar)
- **Column headers** (A, B, C, ... letters across top of grid)
- **Row headers** (1, 2, 3, ... numbers down left side of grid)
- **Grid area** (main content area)
- **Sheet tabs** (bottom of screen) with + button and hamburger menu icon
- **Status bar** (bottom right) showing Sum, Avg, Min, Max, Count for selected range
- **Comments button** (top right)
- **Share button** (blue, top right)
- **"All changes saved in Drive"** indicator (top bar)
- **User avatar/email** (top right corner)

### 2.2 Toolbar Buttons (Left to Right)
1. **Print** button
2. **Undo** button
3. **Redo** button
4. **Copy cell formatting** (paint format / format painter)
5. **Format as currency ($)** button
6. **Format as percent (%)** button
7. **Decrease decimal places** button
8. **Increase decimal places** button
9. **Format as number (123)** dropdown
10. **Font family** dropdown (default: Arial)
11. **Font size** dropdown (default: 10)
12. **Bold** button (B)
13. **Italic** button (I)
14. **Strikethrough** button (S)
15. **Text color** button (A with color bar)
16. **Fill color / background color** button (paint bucket)
17. **Borders** button (grid icon) with dropdown
18. **Merge cells** button with dropdown
19. **Horizontal alignment** button (left/center/right) with dropdown
20. **Vertical alignment** button (top/middle/bottom) with dropdown
21. **Text wrapping** button with dropdown
22. **Text rotation** button with dropdown
23. **Insert link** button (chain icon)
24. **Insert comment** button
25. **Insert chart** button (bar chart icon)
26. **Filter** button (funnel icon)
27. **Functions** button (sigma icon) with dropdown: SUM, AVERAGE, COUNT, MAX, MIN, More functions...

### 2.3 Formula Bar
- **fx** label on the left
- **Formula/value input field** stretching across the width
- Shows the formula of the selected cell (e.g., `=SUM(B2:B8)`)
- Clicking a cell shows its underlying formula here
- Formula help tooltip: blue highlight and question mark icon appear when typing formulas
- Formula autocomplete/suggestions while typing

### 2.4 Sheet Tabs (Bottom)
- **+ button** to add new sheet
- **Hamburger menu** (three horizontal lines) icon for sheet list
- **Sheet tabs** with names (default: "Sheet1")
- **Tab dropdown arrow** on each tab for sheet options
- **Tab right-click context menu** (rename, delete, duplicate, move, etc.)
- **Double-click** tab to rename
- **Drag tabs** to reorder sheets
- **Tab color** can be changed

### 2.5 Status Bar (Bottom Right)
- Displays computed values for selected range:
  - **Sum**: total of selected numbers
  - **Avg**: average of selected numbers
  - **Min**: minimum value
  - **Max**: maximum value
  - **Count**: count of non-empty cells
- Click SUM label to change default formula
- Auto-displays SUM for pure number selections, COUNT for mixed selections

---

## 3. MENUS AND THEIR ITEMS

### 3.1 File Menu
- Share...
- New > (Document, Spreadsheet, Presentation, Form, Drawing, From template)
- Open... (shortcut displayed)
- Rename...
- Make a copy...
- Move to folder...
- Move to trash
- Import...
- See revision history (Ctrl+Option+Shift+G / shortcut displayed)
- Spreadsheet settings...
- Download as > (Microsoft Excel .xlsx, OpenDocument format .ods, PDF document .pdf, Comma-separated values .csv current sheet, Tab-separated values .tsv current sheet, Web page .zip)
- Publish to the web...
- Email collaborators...
- Email as attachment...
- Print (shortcut displayed)

### 3.2 Edit Menu
- Undo / Redo
- Cut / Copy / Paste
- Paste special > (Paste values only, Paste format only, Paste all except borders, Paste formula only, Paste data validation only, Paste conditional formatting only, Paste transpose)
- Find and replace
- Delete values / Delete rows / Delete columns

### 3.3 View Menu
- **Freeze** > (No rows, 1 row, 2 rows, Up to current row (N), No columns, 1 column, 2 columns, Up to current column (X))
- **Gridlines** (toggle on/off)
- **Protected ranges**
- **Formula bar** (toggle visibility)
- **All formulas** (Ctrl+` to toggle showing all formulas)
- **Hidden sheets** > (submenu showing hidden sheet names)
- **Compact controls** (Ctrl+Shift+F)
- **Full screen**

### 3.4 Insert Menu
- Row above
- Row below
- Column left
- Column right
- New sheet
- Comment (Cmd+Option+M / Ctrl+Alt+M)
- Note (Shift+F2)
- Function > (submenu)
- Chart...
- Image...
- Link... (Cmd+K / Ctrl+K)
- Form...
- Drawing...

### 3.5 Format Menu
- Number formatting (various number formats)
- Text formatting (bold, italic, underline, strikethrough)
- Alignment
- Conditional Formatting... (opens side panel)
- Alternating colors
- Cell size (row height, column width)
- Text rotation
- Text wrapping

### 3.6 Data Menu
- Sort sheet by column (A-Z / Z-A)
- Sort range...
- Filter / Filter views
- Data validation...
- Pivot table (mentioned in context)
- Named ranges
- Protected sheets and ranges

### 3.7 Tools Menu
- Script editor (opens Google Apps Script)
- Notification rules
- Spelling

### 3.8 Add-ons Menu
- Get add-ons...
- Manage add-ons...
- (Installed add-ons appear as submenu items, e.g., Google Analytics > Create new report, Run reports, Schedule reports, Help)

### 3.9 Form Menu (appears when form is attached)
- Go To Live Form
- Edit form
- Send form
- Unlink form
- Show form responses

---

## 4. RIGHT-CLICK CONTEXT MENUS

### 4.1 Cell Right-Click Menu
- Cut
- Copy
- Paste
- Paste special > (Paste values only, Paste format only, Paste all except borders, Paste formula only, Paste data validation only, Paste conditional formatting only, Paste transpose)
- Sort range...
- Convert to link / Unlink
- Define named range...
- Protect range...
- Insert comment
- Insert note
- Clear notes
- Conditional formatting...
- Data validation...

### 4.2 Row Header Right-Click Menu
- Insert row above / below
- Delete row
- Clear row
- Hide row
- Unhide rows (when adjacent rows are hidden)
- Resize row

### 4.3 Column Header Right-Click Menu
- Cut
- Copy
- Paste
- Paste special >
- Insert N left / Insert N right (where N = number of selected columns)
- Delete columns A-B (range)
- Clear columns A-B
- Hide columns A-B
- Unhide columns (when adjacent hidden)
- Resize column
- Sort sheet A-Z / Sort sheet Z-A

### 4.4 Sheet Tab Right-Click Menu
- Rename
- Duplicate
- Delete
- Copy to... (another spreadsheet)
- Hide / Unhide
- Change color
- Protect sheet
- Move left / Move right

---

## 5. DATA ENTRY AND INPUT METHODS

### 5.1 Direct Typing
- Click a cell and start typing (no double-click needed for new data)
- Data populates immediately in the selected cell
- One value/word/piece of data per cell (best practice)
- Formula bar also updates as you type

### 5.2 Confirming/Navigating After Entry
- **Enter**: Save data and move to the beginning of the next row below
- **Tab**: Save data and move to the right in the same row
- **Arrow Keys** (Up/Down/Left/Right): Save and move 1 cell in that direction
- **Click any cell**: Jump directly to that cell (saves current entry)
- **Escape**: Cancel current entry

### 5.3 Copy and Paste
- Copy/paste text or numbers from any source
- Copy/paste HTML tables from websites (preserves column structure)
- Single-click on destination cell before pasting = data distributes into individual cells
- Double-click on destination cell before pasting = all data goes into one cell
- **Paste Special** options:
  - Paste values only (strips formulas, keeps text results)
  - Paste format only
  - Paste all except borders
  - Paste formula only
  - Paste data validation only
  - Paste conditional formatting only
  - Paste transpose (swap rows and columns)

### 5.4 Import
- **File > Import > Upload**: Import files from local computer
- Supported formats: **CSV** (comma-separated values), **XLS** (Excel), **XLSX** (Excel)
- Import actions:
  - Create new spreadsheet
  - Insert new sheet(s)
  - Replace spreadsheet
  - Replace current sheet
  - Append rows to current sheet
  - Replace data starting at selected cell
- Separator character options: Detect automatically, Tab, Comma, Custom
- Import from Google Drive: search files within import window

### 5.5 Drag-Fill (Auto-Fill)
- Small blue dot (fill handle) in the bottom-right corner of selected cell
- **Click and drag** the blue dot down or across to fill neighboring cells
- Behaviors:
  - Copy a cell's data to neighboring cells (including formatting)
  - Copy a cell's formula to neighboring cells (with relative reference shifting)
  - Create an ordered/incrementing list (e.g., "Contestant 1" dragged creates "Contestant 2, 3, 4...")
  - Increment numbers: if text ends with a number, Sheets auto-increments by +1
  - Without a trailing number, dragging copies the exact text

### 5.6 Data from External Sources
- **IMPORTXML()**: Import a site's XML markup
- **IMPORTDATA()**: Import content of a page saved as .csv or .tsv
- **IMPORTFEED()**: Import an RSS or ATOM feed
- **IMPORTHTML()**: Import a page's HTML tables
- **IMPORTRANGE()**: Import data from another Google Sheet (requires spreadsheet ID and range)
- **=FINANCE()**: Built-in function for financial data
- **Google Forms**: Responses auto-populate in a linked sheet
- **Add-ons**: Google Analytics, etc. can push data into sheets
- **Copy/paste from websites**: HTML table data preserves structure

### 5.7 Linking to External Content
- **Cmd+K / Ctrl+K**: Insert link dialog
  - Shows search results from Google
  - Enter URL manually
  - Link text field
  - Apply button
- **=HYPERLINK(url, label)**: Create clickable link via formula
- Hover over linked cell to see URL popup

---

## 6. CELL SELECTION AND NAVIGATION

### 6.1 Single Cell Selection
- **Click** on a cell to select it
- **Arrow keys** to move one cell in any direction
- Selected cell has blue outline

### 6.2 Range Selection
- **Click and drag** across cells
- **Click first cell, hold Shift, click last cell** to select a range (e.g., A1 to A10)
- **Click column header letter** to select entire column
- **Click row header number** to select entire row
- **Click gray box in top-left corner** (intersection of row/column headers) to select ALL cells in sheet
- **Ctrl/Cmd + Click** to select non-contiguous cells (implied)

### 6.3 Navigation
- **Arrow keys**: Move one cell at a time
- **Enter**: Move down one row
- **Tab**: Move right one cell
- **Scroll**: Mouse wheel or trackpad to scroll the grid
- **Click any cell**: Jump to that cell
- **Ctrl+G / Cmd+G or Name Box**: Navigate to specific cell reference

---

## 7. TEXT AND CELL FORMATTING

### 7.1 Font Formatting
- **Font family** selector (default: Arial, also shows Open Sans in examples)
- **Font size** selector (default: 10, examples show 12)
- **Bold** (B button / Ctrl+B / Cmd+B)
- **Italic** (I button / Ctrl+I / Cmd+I)
- **Strikethrough** (S button)
- **Text color** (A with colored underline, color picker)
- **Underline** (mentioned in formatting style options)

### 7.2 Cell Formatting
- **Fill/background color** (paint bucket icon, color picker)
- **Borders** (grid icon with dropdown):
  - All borders
  - Inner borders
  - Outer borders
  - Top/Bottom/Left/Right border individually
  - Clear borders
  - Border color
  - Border style (solid, dashed, dotted, thick)
- **Merge cells** button with options:
  - Merge all
  - Merge horizontally
  - Merge vertically
  - Unmerge
- **Cell outlines/borders** (toolbar section labeled "Cell Outlines")

### 7.3 Alignment
- **Horizontal alignment**: Left, Center, Right
- **Vertical alignment**: Top, Middle, Bottom
- **Text wrapping**: Overflow, Wrap, Clip
- **Text rotation**: options for angled/vertical text
- **Indentation**: Increase indent, Decrease indent

### 7.4 Number Formatting
- **Format as currency ($)**: Adds dollar sign and decimal places
- **Format as percent (%)**: Multiplies by 100 and adds % sign
- **Increase/Decrease decimal places**: Adjust decimal precision
- **Number format (123)** dropdown:
  - Automatic
  - Plain text
  - Number (1,000.12)
  - Percent (10.12%)
  - Scientific (1.01E+03)
  - Accounting
  - Financial
  - Currency ($1,000.12)
  - Currency (rounded) ($1,000)
  - Date (9/26/2008)
  - Time (3:59:00 PM)
  - Date time (9/26/2008 15:59:00)
  - Duration (24:01:00)
  - Custom number format
- Formatting applied to a whole row/column applies to future values in that row/column

### 7.5 Conditional Formatting
- Access: **Format > Conditional Formatting** or **Right-click > Conditional formatting**
- Opens **Conditional format rules** side panel on the right
- Two tabs: **Single color** and **Color scale**
- **Single color** options:
  - Apply to range (editable cell reference, e.g., A15, or E1:E10,E13:E1000)
  - Format cells if... (dropdown conditions):
    - Cell is empty
    - Cell is not empty
    - Text contains
    - Text does not contain
    - Text starts with
    - Text ends with
    - Text is exactly
    - Date is / Date is before / Date is after
    - Greater than / Greater than or equal to
    - Less than / Less than or equal to
    - Is equal to / Is not equal to
    - Is between / Is not between
    - Custom formula is (enter any formula returning TRUE/FALSE)
  - Formatting style: Default dropdown, plus manual:
    - Bold, Italic, Underline, Strikethrough
    - Text color picker
    - Background color picker
  - Done / Cancel buttons
  - Add another rule
- **Color scale** options:
  - Minpoint / Midpoint / Maxpoint
  - Type: Number, Percent, Percentile, None (for midpoint)
  - Color pickers for each point
  - Preview bar showing gradient
- Multiple rules can be stacked (evaluated top to bottom)
- Custom formula example: `=AND($C:$C<>"", $K:$K="")` -- uses $ for absolute column references
- Conditional formatting applies to whole rows when using column references with $

### 7.6 Copy Cell Formatting (Paint Format)
- **Paint format** button in toolbar (paint roller icon)
- Click to copy formatting from selected cell(s)
- Then click/drag on destination cells to apply

---

## 8. FORMULAS AND FUNCTIONS

### 8.1 Formula Entry Methods
1. Select a range, then click formula from toolbar dropdown (result placed adjacent)
2. Select result cell, click formula from toolbar, then select input range
3. Type the formula directly into a cell (start with `=` sign)
- Formula auto-fill/suggestions appear as you type based on formula name
- Question mark icon appears next to cell for formula help tooltip
- Help section shows parameter types and examples

### 8.2 Basic Mathematical Functions
- **SUM(range)**: Adds up all values in a range (e.g., `=SUM(B2:B8)`)
- **AVERAGE(range)**: Finds the average of a range
- **COUNT(range)**: Counts cells containing numeric values in a range
- **MAX(range)**: Finds the highest value in a range
- **MIN(range)**: Finds the lowest value in a range
- **Basic arithmetic**: `+`, `-`, `*`, `/` operators directly in cells (e.g., `=A1*B1+C1`)

### 8.3 Logical Functions
- **IF(condition, true_value, false_value)**: Conditional logic
- **AND(condition1, condition2, ...)**: Returns TRUE if all conditions are true
- **IFERROR(formula, fallback_value)**: Returns fallback if formula errors

### 8.4 Lookup and Reference Functions
- **INDIRECT("SheetName!Range")**: Reference data from another sheet dynamically
- **IMPORTRANGE("spreadsheet_id", "SheetName!Range")**: Import data from another spreadsheet document
- **ROW(reference)**: Returns the row number of a reference
- **CHOOSE(index, value1, value2, ...)**: Select from a list based on index

### 8.5 Text Functions
- **CONCATENATE(text1, text2, ...)**: Combine text from multiple cells
- **& operator**: Alternative for combining text (`=A1&" "&B1`)
- **LEN(text)**: Count characters in a cell
- **HYPERLINK(url, label)**: Create a clickable link
- **IMAGE(url, [mode], [height], [width])**: Display image from URL in a cell
  - Mode: fit to cell, or set specific height/width
- **GOOGLETRANSLATE()**: Translate text between languages (mentioned)
- **DETECTLANGUAGE()**: Identify the language of text (mentioned)

### 8.6 Array and Advanced Functions
- **ARRAYFORMULA(formula)**: Apply a formula to an entire range/column at once
  - e.g., `=ARRAYFORMULA(if(name<>"", ROW(name)-3, ""))`
  - Eliminates need to copy formula to each row
- **TRANSPOSE(range)**: Flip vertical data to horizontal and vice versa
- **SUM(ARRAYFORMULA(range1*range2))**: Multiply arrays then sum

### 8.7 Import Functions
- **IMPORTXML(url, xpath)**: Import data from XML/HTML using XPath selectors
  - XPath examples: `"//title"`, `"//meta[@name='description']/@content"`
  - Limit: 50 IMPORTXML() instances per worksheet
  - Can combine XPath queries with `|` (OR operator)
- **IMPORTDATA(url)**: Import .csv or .tsv from a URL
- **IMPORTFEED(url, "items title", false, count)**: Import RSS/ATOM feed titles
  - Also: `"items URL"` for links
- **IMPORTHTML(url, "table"|"list", index)**: Import HTML tables or lists
- **IMPORTRANGE("spreadsheet_id", "range")**: Import from another Google Sheet
  - Requires "Allow access" authorization on first use

### 8.8 Date/Time Functions
- **TODAY()**: Returns current date
- **DATE()**: Create a date value
- Used in dashboard date range configuration

### 8.9 Financial Functions
- **=FINANCE()**: Import financial data (mentioned)

### 8.10 Formula Behavior
- **Relative references**: When copying a formula, cell references shift automatically (e.g., SUM(B2:B8) becomes SUM(C2:C8) when copied right)
- **Absolute references ($)**: `$` locks a column or row (e.g., `$I2` always references column I)
- **Mixed references**: e.g., `$I2` (locked column, relative row)
- **Cross-sheet references**: `SheetName!CellRange` (e.g., `'Current Month Blog Metrics'!B11`)
- **= sign** to start referencing another cell/sheet (type `=` then navigate to cell)
- **Error types**: `#REF!`, `#ERROR!`, `#NAME?`, `#N/A` -- Sheets shows error tooltip on hover
- **Formula help context**: Question mark icon toggles formula documentation on/off

---

## 9. CHARTS AND VISUALIZATIONS

### 9.1 Inserting Charts
- **Insert > Charts** from top menu
- First highlight the data range to chart
- Chart appears floating on top of the spreadsheet

### 9.2 Chart Editor Dialog
- Three tabs: **Recommendations**, **Chart types**, **Customization**
- **Recommendations**: Sheets suggests chart types based on data
- **Chart types**: Full list of available chart formats
- **Customization**: Edit data, labels, chart style
- Data range selector
- **Insert** and **Cancel** buttons

### 9.3 Chart Types Mentioned
- **Line chart** (line volume chart)
- **Bar chart** (horizontal bars)
- **Column chart** (vertical bars)
- **Pie chart** / Donut chart (shown in recommendations)
- **Area chart**
- **Histogram** (shown in recommendations)
- **Scatter chart** (implied)

### 9.4 Chart Interactions
- **Click and drag** chart to reposition it on the sheet
- **Drag corner squares** to resize the chart
- **Pencil icon** (edit button) on chart to modify data, labels, or chart style
- **Three-dot menu** on chart for additional options
- Charts update automatically when underlying data changes

---

## 10. SORTING AND FILTERING

### 10.1 Sorting
- **Data > Sort sheet by column X, A-Z**: Sort entire sheet ascending by a column
- **Data > Sort sheet by column X, Z-A**: Sort entire sheet descending
- **Right-click column header > Sort sheet A-Z / Z-A**
- **Data > Sort range...**: Sort a specific selected range (preserves other data)
- **Right-click > Sort range...**: Same as above

### 10.2 Filtering
- **Data > Filter**: Adds filter dropdowns to header row of data
- **Filter button** in toolbar (funnel icon)
- Filter dropdown on each column header allows:
  - Sort A-Z / Sort Z-A
  - Filter by condition
  - Filter by values (checkboxes for each unique value)
  - Search within filter
- **Data > Filter views**: Create named filter views
  - Filter views are shareable -- other users see the filtered view when they open it
  - Multiple filter views can be saved
  - Set which part of the spreadsheet should be interactive
  - Readers can click on a column, filter by name or text, see only desired results
  - Preset filters can be created to show specific data subsets

---

## 11. SHARING AND COLLABORATION

### 11.1 Share Dialog
- Access: **File > Share** or blue **Share** button (top right)
- **Sharing settings** dialog:
  - **Link to share**: URL copied to clipboard by default
  - **Who has access** section:
    - Lists current collaborators with roles (owner, editor, viewer)
    - **Change...** button to modify access level
  - **Invite people**: Email input field + **Can edit** dropdown (Can edit / Can comment / Can view)
  - **Owner settings**:
    - Prevent editors from changing access and adding new people
    - Disable options to download, print, and copy for commenters and viewers
  - **Done** button

### 11.2 Link Sharing Options
- **On - Public on the web**: Anyone on the internet can find and access
- **On - Anyone with the link**: No sign-in required
- **On - [Organization Name]**: Anyone at organization can find and access
- **On - Anyone at [Organization Name] with the link**
- **Off - Specific people**: Shared with specific people only

### 11.3 Permission Levels
- **Can edit**: Full editing access
- **Can comment**: Can add comments but not edit data
- **Can view**: Read-only access

### 11.4 Collaboration Features
- Real-time co-editing (multiple users editing simultaneously)
- "All changes saved in Drive" indicator
- **Revision history**: File > See revision history (Ctrl+Option+Shift+G)
- Automatic saving (no manual save needed)
- **Comments**: Insert > Comment (Cmd+Option+M)
  - Comments button in top right shows comment thread

---

## 12. DATA PROTECTION AND VALIDATION

### 12.1 Protected Ranges and Sheets
- **Data > Protected sheets and ranges**
- **Right-click > Protect range...**
- **View > Protected ranges**
- Restrict who can edit specific cells, ranges, or entire sheets
- Set permissions per range (specific users or warning-only)

### 12.2 Data Validation
- **Data > Data validation...** or **Right-click > Data validation...**
- Set rules for what data a cell can accept
- Create dropdown lists in cells
- Show validation warnings or reject invalid input

---

## 13. SHEET MANAGEMENT

### 13.1 Creating Sheets
- **+ button** at bottom left to add a new sheet
- **Insert > New sheet** from menu
- Default naming: Sheet1, Sheet2, etc.

### 13.2 Renaming Sheets
- **Double-click** on sheet tab name
- Right-click tab > Rename

### 13.3 Sheet Operations
- Drag tabs to reorder (drag to the left to make it the first tab seen)
- Delete sheets
- Duplicate sheets
- Copy to another spreadsheet
- Hide / Unhide sheets (View > Hidden sheets to see hidden ones)
- Change tab color
- Protect sheet

### 13.4 Cross-Sheet References
- Use `!` to reference another sheet: `SheetName!A2:H`
- Use `=INDIRECT("SheetName!Range")` for dynamic references
- Use `=IMPORTRANGE()` for cross-spreadsheet references

---

## 14. VIEW OPTIONS

### 14.1 Freeze Panes
- **View > Freeze**: Lock rows/columns in place while scrolling
  - Freeze no rows / 1 row / 2 rows / Up to current row
  - Freeze no columns / 1 column / 2 columns / Up to current column
- **Drag the dark grey bar** at top-left of spreadsheet between rows to freeze
  - Bar appears as a thick grey line; cursor changes to a hand when hovering
- Frozen rows/columns stay visible while scrolling through data

### 14.2 Gridlines
- **View > Gridlines**: Toggle gridlines on/off
- Removing gridlines for cleaner dashboard/report look

### 14.3 Full Screen
- **View > Full screen**: Hide menus and toolbars for maximum grid visibility

### 14.4 Compact Controls
- **View > Compact controls** (Ctrl+Shift+F): Reduce toolbar size

### 14.5 Show All Formulas
- **View > All formulas** (Ctrl+`): Toggle between showing values vs. formulas in all cells

### 14.6 Formula Bar Visibility
- **View > Formula bar**: Toggle formula bar on/off

---

## 15. ROWS AND COLUMNS OPERATIONS

### 15.1 Insert
- **Insert > Row above / Row below**
- **Insert > Column left / Column right**
- Right-click row/column header for insert options
- Insert N rows/columns at once (when multiple selected)

### 15.2 Delete
- Right-click > Delete row / Delete column
- Delete selected columns A-B (range delete)

### 15.3 Clear
- Right-click > Clear row / Clear column (removes content but keeps row/column)
- Clear columns A-B

### 15.4 Hide/Unhide
- Right-click row number > **Hide row**
- Right-click column header > **Hide columns A-B**
- Hidden rows/columns: small arrow indicators appear at the border
- To unhide: select surrounding rows/columns > Right-click > **Unhide rows** / **Unhide columns**
- Hidden columns still work in functions/formulas

### 15.5 Resize
- **Drag column border** in header to resize column width
- **Drag row border** in header to resize row height
- Right-click > Resize row / Resize column (enter exact pixel value)
- **Double-click** column border to auto-fit to content width

### 15.6 Drag to Reorder
- Click and drag on the left-hand row number to rearrange rows (move data)

---

## 16. FIND AND REPLACE

- **Edit > Find and replace** (Ctrl+H / Cmd+Shift+H)
- Find field
- Replace field
- Match case option
- Match entire cell contents option
- Search using regular expressions option
- Search within: All sheets / Current sheet

---

## 17. COMMENTS AND NOTES

### 17.1 Comments
- **Insert > Comment** (Cmd+Option+M / Ctrl+Alt+M)
- **Right-click > Insert comment**
- Comments are threaded (can have replies)
- Comments button (top right) shows all comments
- Visible to collaborators
- Can @mention people to notify them
- Resolve/delete comments

### 17.2 Notes
- **Insert > Note** (Shift+F2)
- **Right-click > Insert note**
- **Right-click > Clear notes**: Remove notes from selected cells
- Notes are simpler than comments (no threading)
- Appear as small triangle in cell corner
- Hover to view note content

---

## 18. LINKS AND HYPERLINKS

### 18.1 Insert Link Dialog
- **Cmd+K / Ctrl+K**: Open insert link dialog
- Google search integration: type text and see search results
- Can open search result in new tab or link selected text to that URL
- Text field and Link field
- Apply button

### 18.2 HYPERLINK Function
- `=HYPERLINK("url", "display_text")`
- Hover over cell to see URL popup; click to follow link
- Can combine with other functions for dynamic URLs

### 18.3 Link Behavior
- **Convert to link / Unlink**: Right-click menu option
- Links appear as blue underlined text
- Clickable in the cell

---

## 19. IMAGES

### 19.1 IMAGE Function
- `=IMAGE(url)`: Display an image from a URL inside a cell
- `=IMAGE(url, mode, height, width)`: Set display mode and dimensions
- Mode options: auto-fit to cell height/width, or set manually
- Combined with HYPERLINK: `=HYPERLINK(A1, IMAGE(B1))` -- clickable image
- Only works in Google Sheets (not saved if exported to Excel)

### 19.2 Insert Image (Menu)
- **Insert > Image**: Full-sized image floating on spreadsheet
- **Insert Image dialog** tabs:
  - Upload
  - Take a snapshot
  - By URL
  - Your albums
  - Google Drive
  - Search (Google Image Search)
    - Color filter options
    - License filter ("labeled for commercial reuse with modification")
- Inserted images:
  - Not tied to any specific cell
  - Can be resized and repositioned freely
  - Better for reports and visual layouts

---

## 20. DOWNLOADING AND EXPORTING

### 20.1 Download As (File > Download as)
- **Microsoft Excel (.xlsx)**
- **OpenDocument format (.ods)**
- **PDF document (.pdf)**
- **Comma-separated values (.csv, current sheet)**
- **Tab-separated values (.tsv, current sheet)**
- **Web page (.zip)**

### 20.2 Publish to the Web
- **File > Publish to the Web**
- **Link** tab: Get a public URL
  - Entire Document or specific sheet
  - Format: Web page
  - Publish button
- **Embed** tab: Get HTML embed code
  - Copy code to paste into blog/website
  - Embeds a live copy of the spreadsheet
- Published content & settings section

### 20.3 Print
- **File > Print** or printer icon in toolbar
- Print options (mentioned): print sheet, save as PDF to Google Drive

### 20.4 Email
- **File > Email collaborators...**: Send email notification to collaborators
- **File > Email as attachment...**: Send spreadsheet as email attachment

---

## 21. SPREADSHEET CREATION METHODS

### 21.1 Three Ways to Create
1. Google Drive dashboard > red "NEW" button > Google Sheets
2. Inside a spreadsheet: File > New > Spreadsheet
3. Google Sheets homepage (sheets.google.com): Click "Blank" or select a template

### 21.2 Templates
- Available from Google Sheets homepage
- Examples mentioned: Blank, To-do list, 2016 Calendar, Invoice
- **File > Make a copy**: Duplicate an existing spreadsheet as a template

### 21.3 Spreadsheet Settings
- **File > Spreadsheet settings...**
- Locale, timezone, calculation settings, etc.

---

## 22. OFFLINE MODE

- Google Sheets supports **Offline Mode**
- Automatically syncs changes when reconnected to the internet
- Requires setup (Chrome extension mentioned)
- Allows spreadsheet use without internet connection

---

## 23. NAMED RANGES

- **Right-click > Define named range...**
- **Data > Named ranges**
- Opens **Named ranges** panel on the right side
- Fields:
  - **Name Your Range**: Text input for alias name (e.g., "CRMC1:27")
  - **Range**: Cell range reference (auto-populated or editable, e.g., CRM!C3:G738)
  - **Done** button
- Shows list of all named ranges with their sheet and range
- Case-sensitive names
- Can be used in formulas: `=ARRAYFORMULA(twitter)` instead of `=ARRAYFORMULA(B2:B)`
- Simplifies complex formulas and cross-sheet references

---

## 24. CONDITIONAL FORMATTING (DETAILED)

### 24.1 Single Color Rules
- Format cells if conditions (full list):
  - Cell is empty / Cell is not empty
  - Text contains / does not contain / starts with / ends with / is exactly
  - Date is / is before / is after
  - Greater than / Greater than or equal to
  - Less than / Less than or equal to
  - Is equal to / Is not equal to
  - Is between / Is not between
  - **Custom formula is**: Enter any formula (e.g., `=AND($C:$C<>"", $K:$K="")`)
- Formatting style options:
  - Bold / Italic / Underline / Strikethrough
  - Text color
  - Background/fill color
- Can add multiple rules (processed in order)

### 24.2 Color Scale Rules
- Continuous gradient coloring based on values
- Min point / Midpoint (optional) / Max point
- Each point has:
  - Type: Number, Percent, Percentile, None
  - Value
  - Color picker
- Creates heat-map style coloring across a range

### 24.3 Example Uses from Book
- Highlight rows where data is missing (custom formula checking multiple columns)
- Color cells red if value < 0, green if value > 0 (growth/decline indicators)
- Color budget cells based on spend thresholds (min/mid/max color scale)
- Highlight cells containing specific text (e.g., "Published" status)

---

## 25. PASTE SPECIAL OPTIONS (FULL LIST)

- **Paste values only**: Strip formulas, keep computed values
- **Paste format only**: Apply formatting without data
- **Paste all except borders**: Paste content and formatting without border styles
- **Paste formula only**: Paste only the formula
- **Paste data validation only**: Paste only validation rules
- **Paste conditional formatting only**: Paste only conditional formatting rules
- **Paste transpose**: Swap rows and columns during paste

---

## 26. FUNCTIONS MENTIONED ACROSS ALL CHAPTERS

### Mathematical
- SUM(), AVERAGE(), COUNT(), MAX(), MIN()
- Basic arithmetic: +, -, *, /

### Logical
- IF(), AND(), IFERROR()

### Lookup/Reference
- INDIRECT(), IMPORTRANGE(), ROW(), CHOOSE()

### Text
- CONCATENATE(), & (concatenation operator), LEN(), HYPERLINK()

### Array
- ARRAYFORMULA(), TRANSPOSE()

### Import/Web
- IMPORTXML(), IMPORTDATA(), IMPORTFEED(), IMPORTHTML(), IMPORTRANGE()

### Image
- IMAGE()

### Date/Time
- TODAY(), DATE()

### Financial
- FINANCE()

### Translation
- GOOGLETRANSLATE(), DETECTLANGUAGE()

---

## 27. KEYBOARD SHORTCUTS MENTIONED

| Action | Mac | Windows/Chrome |
|--------|-----|----------------|
| Insert link | Cmd+K | Ctrl+K |
| Insert comment | Cmd+Option+M | Ctrl+Alt+M |
| Insert note | Shift+F2 | Shift+F2 |
| See revision history | Cmd+Option+Shift+G | Ctrl+Alt+Shift+G |
| Compact controls | Cmd+Shift+F | Ctrl+Shift+F |
| Show all formulas | Ctrl+` | Ctrl+` |
| Print | Cmd+P | Ctrl+P |
| Undo | Cmd+Z | Ctrl+Z |
| Redo | Cmd+Y | Ctrl+Y |
| Bold | Cmd+B | Ctrl+B |
| Italic | Cmd+I | Ctrl+I |
| Open file | Cmd+O | Ctrl+O |
| Navigate cells | Arrow keys | Arrow keys |
| Confirm entry (next row) | Enter | Enter |
| Confirm entry (next column) | Tab | Tab |

---

## 28. MOUSE ACTIONS AND INTERACTIONS

| Action | What It Does |
|--------|-------------|
| Click cell | Select the cell |
| Double-click cell | Enter edit mode for that cell |
| Click and drag across cells | Select a range |
| Click column header letter | Select entire column |
| Click row header number | Select entire row |
| Click top-left gray box | Select all cells |
| Right-click cell/row/column | Open context menu |
| Click and drag fill handle (blue dot) | Auto-fill / copy data downward or across |
| Drag column border in header | Resize column width |
| Drag row border in header | Resize row height |
| Double-click column border | Auto-fit column width to content |
| Drag freeze bar (dark grey line at top-left) | Freeze/unfreeze rows or columns |
| Click and drag chart | Move chart position |
| Drag chart corner squares | Resize chart |
| Click chart pencil icon | Edit chart data/style |
| Double-click sheet tab | Rename sheet |
| Drag sheet tabs | Reorder sheets |
| Drag row numbers (left side) | Reorder/move rows |
| Hover over icon in toolbar | See description and shortcut key |
| Hover over linked cell | See URL popup |
| Hover over cell with note | See note content |
| Click and drag to select, then check bottom-right corner | See Sum/Avg/Count/Min/Max in status bar |
| Shift+Click last cell | Extend selection to that cell |
| Click formula question mark | Toggle formula help tooltip |

---

## 29. DASHBOARD FEATURES (CHAPTER 5)

### 29.1 Dashboard Layout Techniques
- Use a dedicated "Dashboard" sheet (rename Sheet1)
- Drag Dashboard tab to the left to make it the first visible sheet
- Enter metric labels in column A
- Populate data cells with `=` references to other sheets
- Apply background colors for each metric section (e.g., red headers)
- Number formatting to make values stand out
- Remove gridlines for cleaner appearance (View > Gridlines off)
- Arrange charts in a grid layout on the right side

### 29.2 Cross-Sheet Cell References for Dashboards
- Type `=` in dashboard cell, navigate to data sheet, click source cell, press Enter
- Formula created: `='Sheet Name'!CellRef` (e.g., `='Current Month Blog Metrics'!B11`)
- Data auto-updates when source sheet changes

### 29.3 Growth/Change Calculations
- Difference: `=SUM(current_cell - previous_cell)` or `=cell1 - cell2`
- Percentage change: `=(Value1 - Value2) / Value2`
- Apply conditional formatting to growth cells (green for positive, red for negative)

### 29.4 Report Scheduling
- Add-ons > Google Analytics > Schedule reports
- Automate data refresh on a schedule
- TODAY() and DATE() functions for dynamic date ranges

---

## 30. GOOGLE FORMS INTEGRATION

### 30.1 Creating Forms from Sheets
- **Insert > Form**: Opens Google Forms editor in new window
- Form responses automatically saved in a new sheet tab ("Form Responses 1")
- Linked form shows **Form** menu in the menu bar

### 30.2 Form Response Sheet
- Automatic timestamp column
- One column per form question
- New responses appended as new rows
- Can rename the response sheet tab (double-click)

### 30.3 Form Menu (appears when form linked)
- Go To Live Form
- Edit form
- Send form
- Unlink form
- Show form responses

---

## 31. ADD-ONS SYSTEM

### 31.1 Add-ons Menu
- **Add-ons > Get add-ons...**: Opens Google Sheets Add-ons Store
- **Add-ons > Manage add-ons...**: View/remove installed add-ons
- Installed add-ons appear as submenu items

### 31.2 Add-on Integration Points
- Side panel on right side of editor
- Menu items under Add-ons menu
- Can add custom functions
- Can add custom menus
- Can run scheduled operations

---

## 32. MISCELLANEOUS FEATURES

### 32.1 Revision History
- **File > See revision history** (Ctrl+Option+Shift+G)
- View previous versions of the spreadsheet
- See who made changes and when
- Restore previous versions

### 32.2 Spreadsheet Settings
- **File > Spreadsheet settings...**
- Locale settings
- Timezone
- Calculation settings
- Iteration settings

### 32.3 Templates
- Available from Google Sheets homepage
- Pre-populated spreadsheets for common use cases
- **File > Make a copy** to duplicate any spreadsheet

### 32.4 Mobile App
- Google Sheets mobile app for iOS and Android
- View, edit, share spreadsheets on the go
- Companion to (not replacement for) the web app

### 32.5 Google Drive Integration
- Spreadsheets stored in Google Drive
- Folder organization
- Star/favorite spreadsheets
- Move to folder (File > Move to folder)
- File > Move to trash

### 32.6 Error Handling
- Error types displayed in cells: `#REF!`, `#ERROR!`, `#NAME?`, `#N/A`
- Hover/click on error cell to see error description tooltip
- `=IFERROR(formula, fallback)` to handle errors gracefully

### 32.7 Automatic URL Detection
- URLs typed or pasted are automatically converted to clickable hyperlinks
- Can convert to link or unlink via right-click menu

### 32.8 Cell Reference Types
- **Relative** (A1): Shifts when formula is copied
- **Absolute** ($A$1): Locked row and column
- **Mixed** ($A1 or A$1): One dimension locked

---

## 33. FEATURES MENTIONED BUT NOT DETAILED IN BOOK

These features are referenced or visible in screenshots but not fully explained. They are
part of Google Sheets and should be included in a production clone:

- **Data validation dropdowns**: Create dropdown menus in cells with preset values
- **Alternating colors** (Format > Alternating colors): Zebra-stripe rows
- **Pivot tables** (Data > Pivot table): Summarize data dynamically
- **Explore button** (bottom right, sparkle icon): AI-powered data insights
- **Version history**: Detailed change tracking with named versions
- **Notification rules** (Tools > Notification rules): Email alerts on changes
- **Spelling check** (Tools > Spelling)
- **Drawing** (Insert > Drawing): Create shapes and diagrams
- **Checkbox** data type: Boolean checkboxes in cells
- **Text rotation**: Angle text within cells
- **Merge cells**: Various merge options (all, horizontal, vertical, unmerge)
- **Cell borders**: Full border customization (style, color, thickness)
- **Frozen columns**: Same as frozen rows but for columns
- **Custom number formats**: User-defined number display patterns
- **Data range selection widget**: Grid icon next to range inputs in dialogs
- **"More" button** in toolbar (overflow for additional tools)
- **Last edit timestamp**: "Last edit was yesterday at 8:38 PM" displayed in header area
- **Saving indicator**: "Saving..." appears during saves, changes to "All changes saved in Drive"

---

## SUMMARY: FEATURE COUNT BY CATEGORY

| Category | Approximate Feature Count |
|----------|--------------------------|
| UI Elements (toolbar, menus, dialogs, panels) | ~85 |
| Keyboard Shortcuts | ~16 |
| Mouse Actions/Interactions | ~25 |
| Formulas/Functions | ~30+ |
| Number Format Options | ~15 |
| Data Entry Methods | ~8 |
| Cell/Text Formatting Options | ~25 |
| Conditional Formatting Conditions | ~18 |
| Paste Special Options | ~7 |
| Chart Types | ~7 |
| Export Formats | ~6 |
| Sharing/Permission Options | ~8 |
| View Options | ~8 |
| Sheet Management Operations | ~10 |
| Right-Click Context Menu Items | ~35 |
| **TOTAL DISTINCT FEATURES** | **~300+** |
