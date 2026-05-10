# xlsx-interact

CLI tool for reading `.xlsx` files. REPL by default, one-shot for scripting.

## Setup

```bash
cd ~/dev/github/xlsx-interact
uv sync
```

## Usage

### REPL (interactive)

```bash
uv run xlsx data.xlsx
xlsx> help
xlsx> cell B2
[Manual DO] B2 = 95 (number)
xlsx> cell B2 --computed
[Manual DO] B2 = 95 (number)
xlsx> cell B2 --style
[Manual DO] B2 = 95 (number)
  Font: bold 11.0 Calibri #FF000000
  Fill: fg=#none bg=#none
  Border: top=thin bottom=thin left=thin right=thin
  Align: h=general v=bottom wrap=no
  Format: General
xlsx> cell A2
[Manual DO] A2 = =IF(B2=0,"Hide","Show") (formula)
xlsx> cell A2 --computed
[Manual DO] A2 = Show (formula)
xlsx> cell B2 --table
[Manual DO] B2 = 95 (number)
  Table: TransactionData (A1:Z1000)
  Column: "Amount"
xlsx> cell C5 --table   (with header=both)
  Table: ProductData (B2:E20)
  Column: "Price"
  Row: "SKU-001"
xlsx> cell E1           (merged cell D1:E1)
[Sheet1] E1 = Sales Projection (text)
  Merged: D1:E1
xlsx> cell E1 --computed
[Sheet1] E1 = Sales Projection (text)
  Merged: D1:E1
xlsx> cell E3 --table   (merged column header)
  Table: Data (A1:E3)
  Column: "Sales Projection"
xlsx> trace C5
[Sheet1] C5 = =(A3-B2)*C1/100 (formula)
  A3 = 50000 (number)
  B2 = 12000 (number)
  C1 = 15 (number)
  C5 = 5700 (computed)
xlsx> cell B2 -s Sheet1
xlsx> sheets
xlsx> info
xlsx> sheet Meta
xlsx> exit
```

### One-shot (scripting)

```bash
uv run xlsx data.xlsx cell B2
uv run xlsx data.xlsx cell B2 -f json
uv run xlsx data.xlsx "cell A2 --computed --style"
uv run xlsx data.xlsx sheets
uv run xlsx data.xlsx info
```

## Commands

| Command | Description |
|---------|-------------|
| `cell <ref>` | Get cell value with auto-detected type (number/text/formula/date/boolean/error/empty). Auto-detects merged cells — if cell is in a merged range, shows value from top-left and the merged range |
| `cell <ref> --computed` | Show computed value instead of formula string |
| `cell <ref> --style` | Show font, fill, border, alignment, number format |
| `cell <ref> --table` | Show containing Excel table name, range, column header, and row label (depends on header type) |
| `cell <ref> -s <sheet>` | Target a specific sheet |
| `trace <ref>` | Trace formula dependencies — extract all cell references in the formula and show their current values (1 level flat) |
| `trace <ref> -s <sheet>` | Trace on a specific sheet |
| `find <query>` | Search cells in current sheet by value or reference (case-insensitive) |
| `find <query> --type number\|text\|formula` | Filter by cell type |
| `find <query> --empty` | Find all empty cells (shortcut for `--type empty`) |
| `find <query> --formula <pattern>` | Find formulas containing a function or pattern |
| `find <query> --merged` | Find cells that belong to merged ranges |
| `doc` | Generate summary report for current sheet |
| `doc --all` | Generate summary report for all sheets |
| `table add <name> <range>` | Register a manual table range for `--table` context |
| `table add --header column\|both\|none` | Set header type (default: row) |
| `table list` | List all manual tables for current sheet |
| `table remove <name>` | Remove a manual table |
| `table clear` | Remove all manual tables |
| `table detect` | Auto-detect table regions via border-line scanning (dry-run) |
| `table detect --register` | Auto-detect and persist to `.xlsx.tables.json` |
| `table detect --register --prefix <name>` | Custom table name prefix |

## Doc Report

```
$ uv run xlsx data.xlsx doc
File: data.xlsx (50,000 bytes)

Sheet1 — 100 rows, 20 cols
  Total cells: 2,000
  Formulas: 150
  Merged ranges: 5
  Tables: (none)

$ uv run xlsx data.xlsx "doc --all"
File: data.xlsx (50,000 bytes)
Sheets: 3

Sheet1 — 100 rows, 20 cols
  Total cells: 2,000
  Formulas: 150
  Merged ranges: 5
  Table: TransactionData (A1:Z1000)

Sheet2 — 50 rows, 10 cols
  Total cells: 500
  Formulas: 30
  Merged ranges: 0
  Tables: (none)

Named ranges: (none)
```

## `table detect` ⚠️ NEEDS IMPROVEMENT

> **Known issues:** columns preview still includes non-descriptive values (numbers like `1.0`, truncation collisions). The detection logic and output formatting need refinement.

Auto-detect table regions by scanning for border separator rows. A separator is a row where a visual border line spans most columns AND at least one adjacent row is mostly empty. The detection checks both the cell's own borders and adjacent cells' borders to catch all visual lines.

```
$ uv run xlsx data.xlsx "table detect"
Detected 2 table(s) in Sheet1:

  SalesReport — A1:D10 (10 rows, 4 cols)
    Header: row (row 1)
    Columns: Product, Qty, Price, Total

  Inventory — A12:D25 (14 rows, 4 cols)
    Header: row (row 12)
    Columns: SKU, Stock, Location, Status

Use `table detect --register` to persist these tables.

$ uv run xlsx data.xlsx "table detect --register"
Registered 2 table(s)

$ uv run xlsx data.xlsx "table detect --register --prefix Report"
(registers as Report_1, Report_2, ...)
```

## Architecture

```
main.py       → session.Session(path)
                    │
                    ├── REPL:       repl.run(session) → commands.parse() → fn()
                    └── one-shot:                     → commands.parse() → fn()
                                                              │
                                                          formatter.render()
```

## Cell Type Detection

Every `cell` output includes the cell's intrinsic type:

| Type | Meaning |
|------|---------|
| `number` | Numeric value |
| `text` | String/label |
| `formula` | Formula expression |
| `date` | Date/time value |
| `boolean` | True/False |
| `error` | Error value (`#N/A`, `#REF!`, etc.) |
| `empty` | Blank cell |

## Number Formats

Excel number formats control how values are displayed without changing the underlying data.

### Structure

```
positive;negative;zero;text
```

A semicolon-separated sequence of up to four sections:

| Section | Applies when value is | Example |
|---------|----------------------|---------|
| 1st | Positive or zero | `#,##0.00` |
| 2nd | Negative | `\-#,##0.00` |
| 3rd | Zero | `"Zero"` |
| 4th | Text | `@` |

If only 2 sections: 1st = positive & zero, 2nd = negative.
If only 1 section: applies to all numbers.

### Placeholder Characters

| Char | Meaning | Example |
|------|---------|---------|
| `0` | Digit placeholder — shows leading/trailing zeros | `00.0` → `05.0` |
| `#` | Digit placeholder — no leading/trailing zeros | `#.#` → `5.` |
| `?` | Space placeholder — aligns decimals with spaces | `??.??` |
| `.` | Decimal point | |
| `,` | Thousands separator or scale | `#,##0` = 1,000; `#,` = 1 (scales by 1000) |

### Text & Literals

| Char | Meaning |
|------|---------|
| `@` | The raw cell text (text placeholder) |
| `\` | Escape next character (display literally) |
| `"text"` | Literal text inside quotes |
| `_` | Space equal to width of next character (for alignment) |
| `*` | Repeat next char to fill cell width |

### Color & Conditions

Prefix a section with a color name in brackets:

```
[Red]#,##0;[Blue]-#,##0
```

Conditional formats use comparison in brackets:

```
[>100]"High";[>0]"Low";"Zero"
```

Supported colors: Black, Blue, Cyan, Green, Magenta, Red, White, Yellow, Color1–Color56.

### Currency & Locale

```
[$RM-4409]
```

- `$RM` = currency symbol (Ringgit Malaysia)
- `4409` = locale LCID (Malay/Malaysia)

Common LCIDs: `409` (en-US), `809` (en-GB), `404` (zh-TW), `804` (zh-CN), `412` (ms-MY), `416` (id-ID).

### Date & Time

| Code | Meaning | Example |
|------|---------|---------|
| `d` | Day (1–31) | |
| `dd` | Day (01–31) | |
| `ddd` | Day abbr | Mon |
| `dddd` | Day full | Monday |
| `m` | Month (1–12) or minute | |
| `mm` | Month (01–12) or minute (00–59) | |
| `mmm` | Month abbr | Jan |
| `mmmm` | Month full | January |
| `yy` | Year (2-digit) | 26 |
| `yyyy` | Year (4-digit) | 2026 |
| `h` / `hh` | Hour | |
| `m` / `mm` | Minute (after h) | |
| `s` / `ss` | Second | |
| `AM/PM` | 12-hour clock | |

### Examples

| Format | 1234.5 | -0.25 | 0 | Text |
|--------|--------|-------|---|------|
| `#,##0` | 1,235 | -0 | 0 | Text |
| `#,##0.00` | 1,234.50 | -0.25 | 0.00 | Text |
| `_-[$RM-4409]* #,##0_-;\-[$RM-4409]* #,##0_-;_-[$RM-4409]* "-"??_-;_-@` | RM 1,235 | -RM 0 | RM - | Text |
| `0%` | 123450% | -25% | 0% | Text |
| `[Red]#,##0;[Blue]-#,##0` | 1,235 (red) | -0 (blue) | 0 (red) | Text |
| `yyyy-mm-dd` | 1903-08-04 | 1899-12-30 | 1900-01-00 | Text |
| `0.0E+00` | 1.2E+03 | -2.5E-01 | 0.0E+00 | Text |

## JSON Output

All commands support `-f json` for machine-readable output:

```bash
uv run xlsx data.xlsx "cell A2 --computed --style --table" -f json
```

```json
{
  "ref": "A2",
  "sheet": "Manual DO",
  "value": "Show",
  "type": "formula",
  "table": {
    "name": "Table at A1",
    "range": "A1:AK646",
    "column": "Status"
  },
  "style": {
    "font": {"bold": false, "italic": false, "size": 16.0, "name": "Century Gothic", "color": "FF000000"},
    "fill": {"fg": null, "bg": null},
    "border": {"top": "thin", "bottom": "thin", "left": "thin", "right": "thin"},
    "align": {"h": "center", "v": "center", "wrap": false},
    "number_format": "General"
  }
}
```

**AI/MCP integration:** Call `commands.COMMANDS["cell"]["fn"](session, ref="A2", computed=True, style=True, table=True)` directly — result is a plain dict.
