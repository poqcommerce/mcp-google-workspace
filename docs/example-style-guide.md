# Example Style Guide: Slides Table Formatting

This is an example style guide for use with the Google Workspace MCP tools. It tells Claude how to create well-formatted, on-brand presentations. See the [README](../README.md#style-guides) for how to use style guides in your own projects.

---

## Brand Colours

- **Dark purple (primary text):** rgb(0.106, 0.078, 0.392)
- **White (#FFFFFF):** content slide backgrounds
- **Dark purple on white:** high-contrast body text

## Table Layout

### Sizing
- Standard slide: 720pt x 405pt
- Tables: ~620pt wide, centered with ~50pt margin each side
- Position at approximately x=50, y=80 to leave room for the slide title

### Column Widths
- Always set explicit column widths — never rely on equal-width defaults
- 2-column label/description: ~25%/75% split (e.g. 155pt / 465pt)
- 3+ columns: allocate width proportional to content length
- Short label columns: 100–180pt; description columns: fill remaining space

### Font Sizing
- 8–12 rows: 10–11pt body, 11–12pt headers
- 13–18 rows: 9pt body, 10pt headers
- 19+ rows: split across two slides
- Headers: always bold, 2–3pt larger than body

### Content Alignment
- `TOP` for multi-line description cells
- `MIDDLE` for single-line tables

### Header Row
- Background: dark purple with white bold text
- Alternative: white background with bold dark purple text

### Body Rows
- White or very light grey background
- All text in dark purple — never black

## Table Creation Workflow

1. `gslides_add_table` — create with position and size
2. `gslides_insert_table_text` — populate cells
3. `gslides_format_table` — set column widths, alignment, header background
4. `gslides_batch_format_text` — apply font size, colour, bold in one call

## General Slide Principles

- One idea per slide
- White background on content slides; branded dark slides for section dividers
- Consistent font sizing across all slides in a deck
- Use a branded template via `gdrive_copy_file` rather than setting colours from scratch
