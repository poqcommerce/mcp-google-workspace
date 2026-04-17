# Example Style Guide: Google Sheets Formatting

A ready-to-adapt style guide for producing well-formatted, readable Google Sheets with the MCP tools. Drop this into your `~/.claude/CLAUDE.md` (or a project-level `CLAUDE.md`) and edit the brand colours, widths, and conventions to match your team's preferences. See the [README](../README.md#style-guides) for how style guides are loaded.

---

## When to apply formatting

Whenever you create a new sheet with `gsheets_create_and_populate`, `gsheets_add_sheet`, or populate new data with `gsheets_append_rows`/`gsheets_batch_update`, immediately follow up with formatting calls. An unstyled sheet of raw values is rarely the desired output.

The two tools that cover most needs:
- `gsheets_format_cells` — text styling, background colours, alignment, wrapping, borders, number formats
- `gsheets_format_dimensions` — column widths, row heights, merge cells, auto-resize

## Brand colours

Replace these with your own. Google Sheets colours are `{red, green, blue}` in the 0–1 range.

- **Header background:** `{red: 0.106, green: 0.078, blue: 0.392}` (dark purple)
- **Header text:** `{red: 1, green: 1, blue: 1}` (white)
- **Section separator background:** `{red: 0.93, green: 0.93, blue: 0.93}` (light grey)
- **Body text:** `{red: 0.15, green: 0.15, blue: 0.15}` (near-black, softer than pure black)
- **Borders:** `{red: 0.8, green: 0.8, blue: 0.8}` (light grey)

## Header rows (row 1)

- **Bold** text, 11pt, white foreground
- **Dark purple** background
- `wrapStrategy: "WRAP"` so long column labels break across two lines rather than overflowing or truncating
- `horizontalAlignment: "CENTER"` for short labels, `"LEFT"` for longer descriptive ones
- `verticalAlignment: "MIDDLE"`
- Row height: 40–50 pixels (tall enough for two wrapped lines)
- Freeze the header row via `gsheets_rename_sheet` with `frozenRowCount: 1`

## Column widths

Always set explicit widths — never rely on Google's defaults, which produce narrow, truncated columns.

- **ID / reference columns:** 60–80 pixels
- **Short labels (status, type, yes/no):** 100–130 pixels
- **Dates:** 100–120 pixels
- **Names / short text:** 150–180 pixels
- **Descriptions / free text:** 250–400 pixels
- **Notes / long-form:** 400+ pixels

When in doubt, use `autoResizeColumns` first, then manually widen any that look cramped.

## Section separators

For sheets with multiple logical sections (e.g. instructions block + data block), use a visually distinct separator row:
- **Bold** text, light grey background
- Horizontal rule: thin bottom border on the row above
- Leave one blank row above and below for breathing room

## Body rows

- Default font size: 10pt
- Near-black text, white background
- `wrapStrategy: "WRAP"` for any column that may contain sentences
- `verticalAlignment: "TOP"` for rows with wrapped content
- No borders between body rows (they add noise); rely on alternating colours if separation is needed

## Number formats

Apply number formats for any column with dates, currency, percentages, or thousands:
- Dates: `{type: "DATE", pattern: "yyyy-mm-dd"}`
- Currency (GBP): `{type: "CURRENCY", pattern: "[$£-809]#,##0.00"}`
- Percentages: `{type: "PERCENT", pattern: "0.0%"}`
- Large numbers: `{type: "NUMBER", pattern: "#,##0"}`

## Data validation (dropdowns)

Not currently exposed in the MCP — if you need dropdown cells, create the sheet with text lookups listed in an unused area and reference them manually, or note this as a follow-up.

## Typical workflow

1. Create sheet and populate data (`gsheets_create_and_populate` or equivalent)
2. `gsheets_format_cells` — apply header styling, body text formatting, number formats, wrapping
3. `gsheets_format_dimensions` — set column widths, row heights, merge any grouped cells
4. `gsheets_rename_sheet` — freeze the header row, set a tab colour if helpful

All four steps are fast — do them proactively rather than waiting for the user to ask.

## What NOT to do

- Don't leave a header row unformatted — bold + background at minimum
- Don't use pure black text — it's visually harsh; use near-black (rgb 0.15)
- Don't rely on default column widths — they always produce truncated text
- Don't add borders everywhere — use them sparingly for emphasis
- Don't use more than two background colours in a sheet — it looks chaotic
