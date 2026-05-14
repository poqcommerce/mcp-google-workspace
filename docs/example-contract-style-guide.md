# Example Style Guide: Google Docs Contracts

A ready-to-adapt style guide for producing consistently formatted contracts, change orders, and other legal/commercial documents with the MCP tools. Drop this into your `~/.claude/CLAUDE.md` (or a project-level `CLAUDE.md`) and edit the fonts, colours, and layout conventions to match your house style. See the [README](../README.md#style-guides) for how style guides are loaded.

---

## When to apply formatting

Contract documents have predictable structure (title, recitals, defined terms, signature block, schedules). Apply formatting in this order:

1. **Set document defaults first** with `gdocs_set_document_defaults` — body font, body size, page margins. Everything inserted afterwards inherits these.
2. **Insert content** — `gdocs_create_document` with initial body text, or `gdocs_create_from_template` if you have a template.
3. **Apply paragraph styles** — `gdocs_format_paragraph` to set heading levels, alignment, spacing, indents.
4. **Apply character styles** — `gdocs_format_text` for bold labels, link styling, defined-term emphasis.
5. **Insert and format tables** — `gdocs_insert_table` + `gdocs_update_table` for field lists, fee schedules, signature blocks.

The best workflow for repeated contract types is templating: build a polished template once in Google Docs (with `{{PLACEHOLDER}}` strings), then `gdocs_create_from_template` for every new instance.

## Page setup

- **Margins:** 72pt top/bottom (1"), 72pt left/right (1") — standard letter/A4.
- For dense legal templates with lots of clauses, tighten to 54pt (0.75") on the sides.

## Body font

Replace these with your house style:

- **Family:** "Red Hat Display" (or "Arial" / "Calibri" / "Times New Roman" — whatever your firm uses)
- **Size:** 11pt
- **Colour:** near-black `{red: 0.15, green: 0.15, blue: 0.15}` (softer than pure black)
- **Line spacing:** 115 (1.15x) — comfortable for reading without wasting page space
- **Space below paragraphs:** 8pt

## Heading hierarchy

Use `namedStyleType` via `gdocs_format_paragraph` (not just font sizing) — this puts headings in the document outline / navigation pane and applies consistently.

| Element | Style | Size | Weight | Notes |
|---|---|---|---|---|
| Document title | `TITLE` | 22pt | Bold | Centred, on its own line at top |
| Section heading | `HEADING_1` | 16pt | Bold | "Terms and Conditions", "Schedule 1", etc. |
| Subsection | `HEADING_2` | 13pt | Bold | "Service Credits", "Termination" |
| Sub-subsection | `HEADING_3` | 11pt | Bold | Numbered clauses if needed |

Set `spaceAbove: 18` / `spaceBelow: 6` on HEADING_1, `spaceAbove: 12` / `spaceBelow: 4` on HEADING_2 — gives clean visual hierarchy without huge gaps.

## Defined terms and emphasis

- **Defined terms** (e.g. "Services", "Customer", "Effective Date"): bold on first occurrence inside the definitions section, then plain text thereafter.
- **Capitalised defined terms** anywhere in the body: leave as-is — the capitalisation alone is the legal signal.
- **Field labels in forms / change orders** (e.g. "Order Start Date:", "Order Term:"): bold, followed by the value in plain text.

## Field-list tables (change orders, order forms)

For a list of fields like "Start Date / Term / Fee / Notes", use a 2-column table rather than running prose:

- **2 columns**: label (narrow) + value (wide)
- **Column widths**: 150pt / 350pt for a standard 6.5"-wide page
- **Borders**: `NONE` — labels and values read cleanly without grid lines
- **Cell padding**: 6pt
- **Header row background**: usually unnecessary for 2-column field lists — only use one if the table has more than 2 columns

Example call:
```
gdocs_insert_table(rows=N, columns=2, cellContent=[["Order Start Date", "1 Jan 2026"], ...])
gdocs_update_table(tableStartIndex=X, rows=N, columns=2, columnWidths=[150, 350], borders="NONE", cellPadding=6)
```

## Fee / schedule tables (multi-column)

For multi-column tables (e.g. fee breakdowns, deliverables × cost × timeline):

- **Header row**: bold text, brand-colour background, white text
- **Header background**: e.g. `{red: 0.106, green: 0.078, blue: 0.392}` (dark purple)
- **Header text colour**: `{red: 1, green: 1, blue: 1}` (white)
- **Borders**: `ALL` (thin grey) for tables that look like data; `NONE` for tables that just structure prose
- **Cell padding**: 6–8pt
- **Content alignment**: `MIDDLE` vertically for short cells, `TOP` for long-text cells
- **Number columns**: right-aligned (via `gdocs_format_paragraph` on the cell range with `alignment: "END"`)

## Signature blocks

Use a 2-column table:
- **Left column**: signer 1 (name, title, company, signature line, date)
- **Right column**: signer 2 (same)
- **Borders**: `NONE`
- **Column widths**: equal (e.g. 250pt each)
- **Cell padding**: 12pt
- Bold the labels ("Name:", "Title:", "Signed:", "Date:") — leave the values plain

## Lists

For terms or deliverable lists:
- **Bullets**: `gdocs_create_bullets` with `BULLET_DISC_CIRCLE_SQUARE` (default).
- **Numbered clauses**: `NUMBERED_DECIMAL_ALPHA_ROMAN` for legal-style nesting (1, a, i).
- Set `indentStart: 36` (0.5") on bullet/list paragraphs for clean hanging indents.

## Hyperlinks

Apply via `gdocs_format_text` with `link: {url: "..."}`. Google Docs auto-styles them (blue + underlined) — don't manually colour them blue or you'll override the link style on update.

## Typical contract workflow

1. `gdocs_create_from_template(templateId, title, replacements, parentFolderId)` — bulk replace placeholders ({{CLIENT}}, {{DATE}}, {{FEE}}).
2. `gdocs_get_document` — read it back, check structure, get table positions if you need to style them.
3. `gdocs_format_paragraph` — adjust any heading styles or alignment that the template didn't cover.
4. `gdocs_update_table` — apply column widths and styling to any tables.
5. `gdocs_get_document` again to verify — and check via `gdrive_get_file_info` that it's in the right folder.

## What NOT to do

- **Don't use `gdocs_format_text` to fake headings** — set `namedStyleType` via `gdocs_format_paragraph` so the doc outline works.
- **Don't manually colour links** — let the default link style win.
- **Don't use multi-paragraph text in a single cell** when a multi-row table would be clearer.
- **Don't apply ALL borders to a layout table** — borders should only appear where they aid readability (data grids), not on field lists.
- **Don't mix font families** — pick one family for the whole document via `gdocs_set_document_defaults` and only deviate for code samples (monospace) if needed.
- **Don't use pure black text** — `rgb(0.15, 0.15, 0.15)` is gentler on the eye and matches how most modern legal templates print.
- **Don't insert tables and forget to set column widths** — Docs defaults to equal-width columns which rarely matches what you need.
