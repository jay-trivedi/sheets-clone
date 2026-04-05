# Google Sheets Exact Styles (from production CSS)

Extracted from `waffle_k_ltr.css` (3.8MB, 23,010 rules)

## Key Colors

| Usage | Color | Hex |
|-------|-------|-----|
| Text primary | `--gm3-sys-color-on-surface` | `#1f1f1f` |
| Text secondary | `--gm3-sys-color-on-surface-variant` | `#444746` |
| Text tertiary | | `#5e5e5e` |
| Text disabled | | `#747775` |
| Surface/background | `--gm3-sys-color-surface` | `#fff` |
| Surface tint | | `#f0f4f9` |
| Surface container | | `#f8fafd` |
| Toolbar background | | `#f0f4f9` |
| Outline/border | `--gm3-sys-color-outline-variant` | `#c4c7c5` |
| Grid line | | `#dadada` |
| Divider | | `#e3e3e3` |
| Active/primary | | `#0b57d0` |
| Selection blue | Active cell border | `#3271ea` |
| Link blue | | `#4285f4` |
| Selection fill | | `#d3e3fd` |
| Selection highlight | | `#c2e7ff` |
| Error/red | | `#b3261e` / `#d14836` |
| Success/green | | `#198639` / `#146c2e` |
| Green container | | `#c4eed0` / `#e7f8ed` |
| Row/col header bg | | `#efeded` |
| Tab border | | `rgba(0,0,0,.1)` |
| Tab text | | `#8f8f8f` |
| Tab selected bg | | `#ababab` |
| Tab selected text | | `#fff` |
| Hover bg | | `rgba(0,0,0,.06)` |
| Active/pressed bg | | `rgba(0,0,0,.12)` |

## Fonts

| Context | Font stack |
|---------|-----------|
| Primary (most common) | `Google Sans, Roboto, Arial, sans-serif` |
| Legacy | `Roboto, RobotoDraft, Helvetica, Arial, sans-serif` |
| Tabs | `Google Sans, Roboto, RobotoDraft, Helvetica, Arial, sans-serif` |
| Name box | `Google Sans, Roboto, sans-serif` (13px) |
| Cell content | `Arial` (default, user-changeable) |

## Font Sizes (by frequency)

| Size | Usage count |
|------|-------------|
| 14px | 544 (most common — toolbar buttons, menus) |
| 12px | 307 (secondary text, labels) |
| 13px | 169 (name box, formula bar) |
| 11px | 102 (small labels) |
| 16px | 83 (titles) |

## Component Specifics

### Row/Column Headers
```css
background-color: #efeded;
width: 45px;  /* row header width */
text-align: center;
vertical-align: middle;
font-size: 8pt;
color: #1f1f1f;
```

### Active Cell Border
```css
border-color: #3271ea;
z-index: 7;
```

### Name Box
```css
font-size: 13px;
width: 69px;
height: 19px;
padding: 0 8px 0 6px;
border-radius: 4px;
background: color-mix(in srgb, #1f1f1f 8%, #fff);
```

### Cell Input (editor)
```css
color: #1f1f1f;
font-family: Google Sans, Roboto, sans-serif;
white-space: pre-wrap;
word-wrap: break-word;
```

### Toolbar
```css
color: #5e5e5e;
font-size: 12px;
line-height: 1;
/* Button hover */
background-color: rgba(0,0,0,.06);
border-radius: 2px;
/* Button active */
background-color: rgba(0,0,0,.12);
```

### Sheet Tabs
```css
font-family: Google Sans, Roboto, RobotoDraft, Helvetica, Arial, sans-serif;
font-weight: 500;
border: 1px solid rgba(0,0,0,.1);
color: #8f8f8f;
padding: 2px 4px;
/* Selected */
background: #ababab;
color: #fff;
```

### Grid Lines
```css
border-color: #dadada;
```
