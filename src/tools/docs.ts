import { docs_v1, drive_v3 } from 'googleapis';
import type {
  CreateDocumentRequest,
  InsertTextRequest,
  ReplaceTextRequest,
  FormatTextRequest,
  FormatParagraphRequest,
  SetHeadingRequest,
  CreateBulletsRequest,
  InsertTableRequest,
  UpdateTableRequest,
  SetDocumentDefaultsRequest,
  CreateFromTemplateRequest,
  FindTextRequest,
  DeleteRangeRequest,
  GetStyleProfileRequest,
  FillTableCellRequest,
  InsertTableRowRequest,
  DeleteTableRowRequest,
  InsertTableColumnRequest,
  DeleteTableColumnRequest,
  DeleteTableRequest,
  ToolDefinition,
  ToolResponse,
} from '../types.js';
import { successResponse, textResponse, errorResponse, normaliseText } from '../utils.js';

// ── Tool definitions ───────────────────────────────────────────────────────────

export function getDocsToolDefinitions(): ToolDefinition[] {
  return [
    {
      name: 'gdocs_create_document',
      description: 'Create a new Google Document',
      inputSchema: {
        type: 'object',
        properties: {
          title: { type: 'string', description: 'Title for the new document' },
          content: { type: 'string', description: 'Initial content for the document (optional)' },
          parentFolderId: {
            type: 'string',
            description: 'ID of the parent folder to create the document in (optional, defaults to root)',
          },
        },
        required: ['title'],
      },
    },
    {
      name: 'gdocs_get_document',
      description:
        'Get the content of a Google Document. Includes detected styles, suggestion summary, and table positions (tableStartIndex) needed for gdocs_update_table.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document to retrieve' },
        },
        required: ['documentId'],
      },
    },
    {
      name: 'gdocs_insert_text',
      description: 'Insert text into a Google Document',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          text: { type: 'string', description: 'Text to insert' },
          index: { type: 'number', description: 'Position to insert text (optional, defaults to end)' },
        },
        required: ['documentId', 'text'],
      },
    },
    {
      name: 'gdocs_append_text',
      description: 'Append text to the end of a Google Document',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          text: { type: 'string', description: 'Text to append' },
        },
        required: ['documentId', 'text'],
      },
    },
    {
      name: 'gdocs_replace_text',
      description: 'Find and replace text in a Google Document',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          find: { type: 'string', description: 'Text to find' },
          replace: { type: 'string', description: 'Text to replace with' },
        },
        required: ['documentId', 'find', 'replace'],
      },
    },
    {
      name: 'gdocs_format_text',
      description:
        'Apply character-level formatting to a text range. Supports bold, italic, underline, strikethrough, fontSize (pt), ' +
        'fontFamily (e.g. "Red Hat Display", "Arial"), foregroundColor, backgroundColor, and link. ' +
        'Colours use {red, green, blue} in 0-1 range. For paragraph-level styling (alignment, headings, spacing), use gdocs_format_paragraph.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          startIndex: { type: 'number', description: 'Start index of text to format' },
          endIndex: { type: 'number', description: 'End index of text to format' },
          format: {
            type: 'object',
            description: 'Formatting options',
            properties: {
              bold: { type: 'boolean' },
              italic: { type: 'boolean' },
              underline: { type: 'boolean' },
              strikethrough: { type: 'boolean' },
              fontSize: { type: 'number', description: 'Font size in points' },
              fontFamily: { type: 'string', description: 'Font family name (e.g. "Red Hat Display", "Arial")' },
              foregroundColor: {
                type: 'object',
                description: '{ red, green, blue } values 0-1',
                properties: {
                  red: { type: 'number' },
                  green: { type: 'number' },
                  blue: { type: 'number' },
                },
              },
              backgroundColor: {
                type: 'object',
                description: '{ red, green, blue } values 0-1',
                properties: {
                  red: { type: 'number' },
                  green: { type: 'number' },
                  blue: { type: 'number' },
                },
              },
              link: {
                type: 'object',
                description: 'Hyperlink target',
                properties: { url: { type: 'string' } },
              },
            },
          },
        },
        required: ['documentId', 'startIndex', 'endIndex', 'format'],
      },
    },
    {
      name: 'gdocs_format_paragraph',
      description:
        'Apply paragraph-level styling to a range: namedStyleType (NORMAL_TEXT/TITLE/SUBTITLE/HEADING_1..6), alignment, ' +
        'lineSpacing (e.g. 115 for 1.15x), spaceAbove/spaceBelow (pt), indents (pt). Setting namedStyleType applies the named ' +
        'style and also makes the paragraph appear in the document outline.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          startIndex: { type: 'number', description: 'Start index of range' },
          endIndex: { type: 'number', description: 'End index of range' },
          style: {
            type: 'object',
            description: 'Paragraph styling options',
            properties: {
              namedStyleType: {
                type: 'string',
                description: 'Named style for this paragraph (controls outline + default appearance)',
                enum: [
                  'NORMAL_TEXT',
                  'TITLE',
                  'SUBTITLE',
                  'HEADING_1',
                  'HEADING_2',
                  'HEADING_3',
                  'HEADING_4',
                  'HEADING_5',
                  'HEADING_6',
                ],
              },
              alignment: {
                type: 'string',
                description: 'Horizontal alignment',
                enum: ['START', 'CENTER', 'END', 'JUSTIFIED'],
              },
              lineSpacing: { type: 'number', description: 'Line spacing as a percentage (115 = 1.15x)' },
              spaceAbove: { type: 'number', description: 'Space before paragraph in points' },
              spaceBelow: { type: 'number', description: 'Space after paragraph in points' },
              indentFirstLine: { type: 'number', description: 'First-line indent in points' },
              indentStart: { type: 'number', description: 'Left indent in points' },
              indentEnd: { type: 'number', description: 'Right indent in points' },
              keepWithNext: { type: 'boolean', description: 'Keep paragraph with the next one (avoid page break between)' },
            },
          },
        },
        required: ['documentId', 'startIndex', 'endIndex', 'style'],
      },
    },
    {
      name: 'gdocs_set_heading',
      description: 'Convert a text range to a heading (sets namedStyleType to HEADING_N). Same effect as gdocs_format_paragraph with namedStyleType, kept for convenience.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          startIndex: { type: 'number', description: 'Start index of text to make a heading' },
          endIndex: { type: 'number', description: 'End index of text to make a heading' },
          headingLevel: {
            type: 'number',
            description: 'Heading level (1-6)',
            minimum: 1,
            maximum: 6,
          },
        },
        required: ['documentId', 'startIndex', 'endIndex', 'headingLevel'],
      },
    },
    {
      name: 'gdocs_create_bullets',
      description:
        'Apply a bulleted or numbered list to a paragraph range. Use preset to choose style: ' +
        'BULLET_DISC_CIRCLE_SQUARE (default), BULLET_DIAMONDX_ARROW3D_SQUARE, BULLET_CHECKBOX, ' +
        'NUMBERED_DECIMAL_ALPHA_ROMAN, NUMBERED_DECIMAL_NESTED, NUMBERED_UPPERALPHA_ALPHA_ROMAN, etc.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          startIndex: { type: 'number', description: 'Start index of range' },
          endIndex: { type: 'number', description: 'End index of range' },
          preset: {
            type: 'string',
            description: 'BulletGlyphPreset value (default: BULLET_DISC_CIRCLE_SQUARE)',
          },
        },
        required: ['documentId', 'startIndex', 'endIndex'],
      },
    },
    {
      name: 'gdocs_insert_table',
      description:
        'Insert a table at the given index and optionally populate cells. cellContent is a 2D array (rows × columns) — missing rows/cells stay empty. ' +
        'Returns the table\'s startIndex which you pass to gdocs_update_table to style it. ' +
        'Tip: use gdocs_get_document afterwards to see exact table cell indices if you need to format individual cells.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          index: { type: 'number', description: 'Position to insert the table' },
          rows: { type: 'number', description: 'Number of rows' },
          columns: { type: 'number', description: 'Number of columns' },
          cellContent: {
            type: 'array',
            description: 'Optional 2D array of cell text. Rows × columns. Cells beyond the array stay empty.',
            items: {
              type: 'array',
              items: { type: 'string' },
            },
          },
        },
        required: ['documentId', 'index', 'rows', 'columns'],
      },
    },
    {
      name: 'gdocs_insert_table_row',
      description:
        'Insert a new row into a table above or below an existing row. tableStartIndex identifies the table (from gdocs_get_document); rowIndex picks the anchor row. Set insertBelow=true (default) to add the new row after the anchor, or false to add it before. Useful for growing tables like fee schedules or SOW deliverable lists. After insertion, indices for content after this table will shift — re-fetch the document or operate top-to-bottom.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          tableStartIndex: { type: 'number', description: 'startIndex of the table (from gdocs_get_document)' },
          rowIndex: { type: 'number', description: '0-based row index of the anchor row' },
          insertBelow: {
            type: 'boolean',
            description: 'true (default) to insert below the anchor row, false to insert above',
          },
        },
        required: ['documentId', 'tableStartIndex', 'rowIndex'],
      },
    },
    {
      name: 'gdocs_delete_table_row',
      description:
        'Delete a row from a table by its 0-based row index. tableStartIndex identifies the table; rowIndex picks the row to remove. The Docs API will refuse to delete the last remaining row of a table — use gdocs_delete_table to remove a single-row table entirely. Indices for content after the table will shift after deletion.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          tableStartIndex: { type: 'number', description: 'startIndex of the table (from gdocs_get_document)' },
          rowIndex: { type: 'number', description: '0-based row index to delete' },
        },
        required: ['documentId', 'tableStartIndex', 'rowIndex'],
      },
    },
    {
      name: 'gdocs_insert_table_column',
      description:
        'Insert a new column into a table to the left or right of an existing column. tableStartIndex identifies the table; columnIndex picks the anchor column. Set insertRight=true (default) to add the new column after the anchor, or false to add before. The new column starts empty in every row.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          tableStartIndex: { type: 'number', description: 'startIndex of the table (from gdocs_get_document)' },
          columnIndex: { type: 'number', description: '0-based column index of the anchor column' },
          insertRight: {
            type: 'boolean',
            description: 'true (default) to insert right of the anchor, false to insert left',
          },
        },
        required: ['documentId', 'tableStartIndex', 'columnIndex'],
      },
    },
    {
      name: 'gdocs_delete_table_column',
      description:
        'Delete a column from a table by its 0-based column index. tableStartIndex identifies the table; columnIndex picks the column to remove. The Docs API will refuse to delete the last remaining column — use gdocs_delete_table to remove a single-column table entirely.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          tableStartIndex: { type: 'number', description: 'startIndex of the table (from gdocs_get_document)' },
          columnIndex: { type: 'number', description: '0-based column index to delete' },
        },
        required: ['documentId', 'tableStartIndex', 'columnIndex'],
      },
    },
    {
      name: 'gdocs_delete_table',
      description:
        'Delete an entire table from the document. tableStartIndex identifies which table (from gdocs_get_document). Removes the table and all its content in one batchUpdate. Useful for stripping template skeletons (e.g. style-legend tables) from a copy. Indices for content after the table will shift.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          tableStartIndex: { type: 'number', description: 'startIndex of the table to delete (from gdocs_get_document)' },
        },
        required: ['documentId', 'tableStartIndex'],
      },
    },
    {
      name: 'gdocs_fill_table_cell',
      description:
        'Insert text into a specific table cell by row/column index. Handles cell content-range calculation internally — you don\'t need to know the cell\'s text indices. Use mode="replace" (default) to overwrite the cell\'s existing content, or mode="append" to add to the end. Pair with gdocs_find_text (which returns tableContext with rowIndex/columnIndex) to fill the cell next to a labelled row.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          tableStartIndex: { type: 'number', description: 'startIndex of the table element (from gdocs_get_document)' },
          rowIndex: { type: 'number', description: '0-based row index within the table' },
          columnIndex: { type: 'number', description: '0-based column index within the row' },
          text: { type: 'string', description: 'Text to insert into the cell' },
          mode: {
            type: 'string',
            description: 'replace (overwrite existing content, default) or append (add to end)',
            enum: ['replace', 'append'],
          },
        },
        required: ['documentId', 'tableStartIndex', 'rowIndex', 'columnIndex', 'text'],
      },
    },
    {
      name: 'gdocs_update_table',
      description:
        'Style a table: column widths (pt), header/body row background colours, cell padding (pt, applied to all cells), borders (NONE for invisible layout tables / ALL for thin grey lines), vertical content alignment. ' +
        'tableStartIndex is the index of the table element — get it from gdocs_get_document.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          tableStartIndex: { type: 'number', description: 'startIndex of the table element (from gdocs_get_document)' },
          rows: { type: 'number', description: 'Total rows in the table' },
          columns: { type: 'number', description: 'Total columns in the table' },
          columnWidths: {
            type: 'array',
            description: 'Column widths in points. Length must match columns count.',
            items: { type: 'number' },
          },
          headerRowBackgroundColor: {
            type: 'object',
            description: 'Background colour for row 0 — { red, green, blue } 0-1',
            properties: {
              red: { type: 'number' },
              green: { type: 'number' },
              blue: { type: 'number' },
            },
          },
          bodyRowBackgroundColor: {
            type: 'object',
            description: 'Background colour for non-header rows — { red, green, blue } 0-1',
            properties: {
              red: { type: 'number' },
              green: { type: 'number' },
              blue: { type: 'number' },
            },
          },
          cellPadding: { type: 'number', description: 'Uniform cell padding in points (applied top/bottom/left/right)' },
          borders: {
            type: 'string',
            description: 'NONE = invisible borders (good for layout tables), ALL = thin grey borders on every edge',
            enum: ['NONE', 'ALL'],
          },
          contentAlignment: {
            type: 'string',
            description: 'Vertical alignment for all cells',
            enum: ['TOP', 'MIDDLE', 'BOTTOM'],
          },
        },
        required: ['documentId', 'tableStartIndex', 'rows', 'columns'],
      },
    },
    {
      name: 'gdocs_set_document_defaults',
      description:
        'Apply document-wide defaults: body font family/size/colour (applied to NORMAL_TEXT paragraphs only — heading/title sizes are left alone) and page margins (points). ' +
        'Best called AFTER inserting body content — the Docs API has no global "default text style", so this applies the style to existing text. ' +
        'Newly inserted text adjacent to styled text will then inherit it. Margins persist regardless of content order. ' +
        'Apply per-range overrides (e.g. bolded labels) AFTER this call, since the defaults will overwrite any prior textStyle on NORMAL_TEXT.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          fontFamily: { type: 'string', description: 'Default body font family (e.g. "Red Hat Display", "Arial")' },
          fontSize: { type: 'number', description: 'Default body font size in points' },
          foregroundColor: {
            type: 'object',
            description: 'Default text colour — { red, green, blue } 0-1',
            properties: {
              red: { type: 'number' },
              green: { type: 'number' },
              blue: { type: 'number' },
            },
          },
          marginTop: { type: 'number', description: 'Top margin in points (72pt = 1 inch)' },
          marginBottom: { type: 'number', description: 'Bottom margin in points' },
          marginLeft: { type: 'number', description: 'Left margin in points' },
          marginRight: { type: 'number', description: 'Right margin in points' },
        },
        required: ['documentId'],
      },
    },
    {
      name: 'gdocs_get_style_profile',
      description:
        'Extract a portable JSON profile of a document\'s styling: page margins, every named style (TITLE, HEADING_1..6, NORMAL_TEXT) with its textStyle + paragraphStyle, the dominant body font, every table pattern (rows×columns, widths, borders, cell backgrounds, padding, alignment), and a list of "suspected headings" — short bold or larger paragraphs that look like headings but lack a namedStyleType. Save the output to disk to reuse as a style reference across multiple documents.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document to extract styles from' },
        },
        required: ['documentId'],
      },
    },
    {
      name: 'gdocs_find_text',
      description:
        'Find one or more occurrences of a text string and return their indices. For each match, returns the matched range (startIndex, endIndex) AND the enclosing paragraph range (paragraphStartIndex, paragraphEndIndex) — use the paragraph range to apply paragraph styles like headings. Also indicates whether the match is inside a table cell. Essential for working with documents after structural edits have shifted indices.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          query: { type: 'string', description: 'Text to find (exact match)' },
          matchCase: {
            type: 'boolean',
            description: 'Case-sensitive search (default: true)',
          },
          firstOnly: {
            type: 'boolean',
            description: 'Return only the first match (default: false — returns all)',
          },
        },
        required: ['documentId', 'query'],
      },
    },
    {
      name: 'gdocs_delete_range',
      description:
        'Delete a precise range of content. Pair with gdocs_find_text to delete a specific paragraph (find the text, take paragraphStartIndex/paragraphEndIndex, delete that range). Safer than gdocs_replace_text with an empty replacement, which matches every occurrence including inside tables.',
      inputSchema: {
        type: 'object',
        properties: {
          documentId: { type: 'string', description: 'The ID of the document' },
          startIndex: { type: 'number', description: 'Start of range to delete' },
          endIndex: { type: 'number', description: 'End of range to delete (exclusive)' },
        },
        required: ['documentId', 'startIndex', 'endIndex'],
      },
    },
    {
      name: 'gdocs_create_from_template',
      description:
        'Copy an existing Google Doc as a template and apply text replacements. Replacements is an object map: each key is a placeholder to find ' +
        '(e.g. "{{CLIENT_NAME}}"), each value is the replacement text. All replacements run in a single batchUpdate. Use parentFolderId to file the new doc.',
      inputSchema: {
        type: 'object',
        properties: {
          templateId: { type: 'string', description: 'The ID of the template document to copy' },
          title: { type: 'string', description: 'Title for the new document' },
          replacements: {
            type: 'object',
            description: 'Map of placeholder strings to replacement values. Example: { "{{CLIENT_NAME}}": "Acme Corp" }',
            additionalProperties: { type: 'string' },
          },
          parentFolderId: {
            type: 'string',
            description: 'Optional folder ID to place the new document in',
          },
        },
        required: ['templateId', 'title'],
      },
    },
  ];
}

// ── Style analysis helpers ────────────────────────────────────────────────────

interface StyleSample {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  foregroundColor?: { red: number; green: number; blue: number };
  charCount: number;
}

function extractDocStyles(doc: docs_v1.Schema$Document): {
  dominantStyle: Omit<StyleSample, 'charCount'> | null;
  headingStyle: Omit<StyleSample, 'charCount'> | null;
} {
  const bodySamples: StyleSample[] = [];
  const headingSamples: StyleSample[] = [];

  const collectFromElements = (elements: docs_v1.Schema$StructuralElement[]) => {
    for (const element of elements) {
      if (element.paragraph) {
        const para = element.paragraph;
        if (!para.elements) continue;

        const isHeading = para.paragraphStyle?.namedStyleType?.startsWith('HEADING');

        for (const elem of para.elements) {
          const run = elem.textRun;
          if (!run?.content || !run.textStyle) continue;

          const text = run.content.replace(/\n/g, '');
          if (text.length === 0) continue;

          const ts = run.textStyle;
          const sample: StyleSample = {
            fontFamily: ts.weightedFontFamily?.fontFamily || undefined,
            fontSize: ts.fontSize?.magnitude || undefined,
            bold: ts.bold || undefined,
            italic: ts.italic || undefined,
            charCount: text.length,
          };

          const fg = ts.foregroundColor?.color?.rgbColor;
          if (fg) {
            sample.foregroundColor = {
              red: Math.round((fg.red || 0) * 1000) / 1000,
              green: Math.round((fg.green || 0) * 1000) / 1000,
              blue: Math.round((fg.blue || 0) * 1000) / 1000,
            };
          }

          if (isHeading) {
            headingSamples.push(sample);
          } else {
            bodySamples.push(sample);
          }
        }
      } else if (element.table) {
        for (const row of element.table.tableRows || []) {
          for (const cell of row.tableCells || []) {
            if (cell.content) collectFromElements(cell.content);
          }
        }
      }
    }
  };

  collectFromElements(doc.body?.content || []);

  return {
    dominantStyle: pickDominant(bodySamples),
    headingStyle: pickDominant(headingSamples),
  };
}

function pickDominant(samples: StyleSample[]): Omit<StyleSample, 'charCount'> | null {
  if (samples.length === 0) return null;

  // Weight by character count to find the most common style
  const fontMap = new Map<string, number>();
  const sizeMap = new Map<number, number>();
  const colorMap = new Map<string, { color: { red: number; green: number; blue: number }; count: number }>();
  let boldChars = 0;
  let italicChars = 0;
  let totalChars = 0;

  for (const s of samples) {
    totalChars += s.charCount;
    if (s.fontFamily) fontMap.set(s.fontFamily, (fontMap.get(s.fontFamily) || 0) + s.charCount);
    if (s.fontSize) sizeMap.set(s.fontSize, (sizeMap.get(s.fontSize) || 0) + s.charCount);
    if (s.bold) boldChars += s.charCount;
    if (s.italic) italicChars += s.charCount;
    if (s.foregroundColor) {
      const key = `${s.foregroundColor.red},${s.foregroundColor.green},${s.foregroundColor.blue}`;
      const existing = colorMap.get(key);
      if (existing) {
        existing.count += s.charCount;
      } else {
        colorMap.set(key, { color: s.foregroundColor, count: s.charCount });
      }
    }
  }

  const result: Record<string, any> = {};

  // Most common font
  let maxFont = 0;
  for (const [font, count] of fontMap) {
    if (count > maxFont) { maxFont = count; result.fontFamily = font; }
  }

  // Most common size
  let maxSize = 0;
  for (const [size, count] of sizeMap) {
    if (count > maxSize) { maxSize = count; result.fontSize = size; }
  }

  // Bold/italic if majority
  if (boldChars > totalChars * 0.5) result.bold = true;
  if (italicChars > totalChars * 0.5) result.italic = true;

  // Most common color
  let maxColor = 0;
  for (const [, entry] of colorMap) {
    if (entry.count > maxColor) { maxColor = entry.count; result.foregroundColor = entry.color; }
  }

  return Object.keys(result).length > 0 ? result : null;
}

// ── Handler class ──────────────────────────────────────────────────────────────

export class DocsHandler {
  constructor(
    private docs: docs_v1.Docs,
    private drive: drive_v3.Drive,
  ) {}

  /** Route a tool call to the appropriate handler. Returns null if not handled. */
  async handleTool(name: string, args: any): Promise<ToolResponse | null> {
    switch (name) {
      case 'gdocs_create_document':
        return this.handleCreateDocument(this.validateCreateDocArgs(args));
      case 'gdocs_get_document':
        return this.handleGetDocument(this.validateDocumentIdArgs(args));
      case 'gdocs_insert_text':
        return this.handleInsertText(this.validateInsertTextArgs(args));
      case 'gdocs_append_text':
        return this.handleAppendText(this.validateInsertTextArgs(args));
      case 'gdocs_replace_text':
        return this.handleReplaceText(this.validateReplaceTextArgs(args));
      case 'gdocs_format_text':
        return this.handleFormatText(this.validateFormatTextArgs(args));
      case 'gdocs_format_paragraph':
        return this.handleFormatParagraph(this.validateFormatParagraphArgs(args));
      case 'gdocs_set_heading':
        return this.handleSetHeading(this.validateSetHeadingArgs(args));
      case 'gdocs_create_bullets':
        return this.handleCreateBullets(this.validateCreateBulletsArgs(args));
      case 'gdocs_insert_table':
        return this.handleInsertTable(this.validateInsertTableArgs(args));
      case 'gdocs_update_table':
        return this.handleUpdateTable(this.validateUpdateTableArgs(args));
      case 'gdocs_fill_table_cell':
        return this.handleFillTableCell(this.validateFillTableCellArgs(args));
      case 'gdocs_insert_table_row':
        return this.handleInsertTableRow(this.validateInsertTableRowArgs(args));
      case 'gdocs_delete_table_row':
        return this.handleDeleteTableRow(this.validateDeleteTableRowArgs(args));
      case 'gdocs_insert_table_column':
        return this.handleInsertTableColumn(this.validateInsertTableColumnArgs(args));
      case 'gdocs_delete_table_column':
        return this.handleDeleteTableColumn(this.validateDeleteTableColumnArgs(args));
      case 'gdocs_delete_table':
        return this.handleDeleteTable(this.validateDeleteTableArgs(args));
      case 'gdocs_set_document_defaults':
        return this.handleSetDocumentDefaults(this.validateSetDocumentDefaultsArgs(args));
      case 'gdocs_create_from_template':
        return this.handleCreateFromTemplate(this.validateCreateFromTemplateArgs(args));
      case 'gdocs_find_text':
        return this.handleFindText(this.validateFindTextArgs(args));
      case 'gdocs_delete_range':
        return this.handleDeleteRange(this.validateDeleteRangeArgs(args));
      case 'gdocs_get_style_profile':
        return this.handleGetStyleProfile(this.validateGetStyleProfileArgs(args));
      default:
        return null;
    }
  }

  // ── Validators ─────────────────────────────────────────────────────────────

  private validateCreateDocArgs(args: any): CreateDocumentRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.title || typeof args.title !== 'string') {
      throw new Error('Invalid title: expected non-empty string');
    }
    if (args.parentFolderId && typeof args.parentFolderId !== 'string') {
      throw new Error('Invalid parentFolderId: expected string');
    }
    return {
      title: args.title,
      content: args.content ? normaliseText(args.content) : '',
      parentFolderId: args.parentFolderId,
    };
  }

  private validateDocumentIdArgs(args: any) {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    return { documentId: args.documentId };
  }

  private validateInsertTextArgs(args: any): InsertTextRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    if (!args.text || typeof args.text !== 'string') {
      throw new Error('Invalid text: expected non-empty string');
    }
    return {
      documentId: args.documentId,
      text: normaliseText(args.text),
      index: args.index || undefined,
    };
  }

  private validateReplaceTextArgs(args: any): ReplaceTextRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    if (!args.find || typeof args.find !== 'string') {
      throw new Error('Invalid find: expected non-empty string');
    }
    if (typeof args.replace !== 'string') {
      throw new Error('Invalid replace: expected string');
    }
    return {
      documentId: args.documentId,
      find: normaliseText(args.find),
      replace: normaliseText(args.replace),
    };
  }

  /** Shared range validation: documentId, startIndex, endIndex. */
  private validateRangeArgs(args: any): { documentId: string; startIndex: number; endIndex: number } {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    if (typeof args.startIndex !== 'number' || args.startIndex < 0) {
      throw new Error('Invalid startIndex: expected non-negative number');
    }
    if (typeof args.endIndex !== 'number' || args.endIndex <= args.startIndex) {
      throw new Error('Invalid endIndex: expected number greater than startIndex');
    }
    return { documentId: args.documentId, startIndex: args.startIndex, endIndex: args.endIndex };
  }

  private validateFormatTextArgs(args: any): FormatTextRequest {
    const range = this.validateRangeArgs(args);
    if (!args.format || typeof args.format !== 'object') {
      throw new Error('Invalid format: expected object');
    }
    return { ...range, format: args.format };
  }

  private validateFormatParagraphArgs(args: any): FormatParagraphRequest {
    const range = this.validateRangeArgs(args);
    if (!args.style || typeof args.style !== 'object') {
      throw new Error('Invalid style: expected object');
    }
    return { ...range, style: args.style };
  }

  private validateSetHeadingArgs(args: any): SetHeadingRequest {
    const range = this.validateRangeArgs(args);
    const headingLevel = args?.headingLevel;
    if (typeof headingLevel !== 'number' || headingLevel < 1 || headingLevel > 6) {
      throw new Error('Invalid headingLevel: expected number between 1 and 6');
    }
    return { ...range, headingLevel };
  }

  private validateCreateBulletsArgs(args: any): CreateBulletsRequest {
    const range = this.validateRangeArgs(args);
    const preset = args?.preset;
    if (preset !== undefined && typeof preset !== 'string') {
      throw new Error('Invalid preset: expected string');
    }
    return { ...range, preset };
  }

  private validateInsertTableArgs(args: any): InsertTableRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    if (typeof args.index !== 'number' || args.index < 0) {
      throw new Error('Invalid index: expected non-negative number');
    }
    if (typeof args.rows !== 'number' || args.rows < 1) {
      throw new Error('Invalid rows: expected positive number');
    }
    if (typeof args.columns !== 'number' || args.columns < 1) {
      throw new Error('Invalid columns: expected positive number');
    }
    if (args.cellContent !== undefined) {
      if (!Array.isArray(args.cellContent)) throw new Error('Invalid cellContent: expected 2D array');
      for (const row of args.cellContent) {
        if (!Array.isArray(row)) throw new Error('Invalid cellContent: expected 2D array of strings');
      }
    }
    return {
      documentId: args.documentId,
      index: args.index,
      rows: args.rows,
      columns: args.columns,
      cellContent: args.cellContent,
    };
  }

  private validateUpdateTableArgs(args: any): UpdateTableRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    if (typeof args.tableStartIndex !== 'number' || args.tableStartIndex < 0) {
      throw new Error('Invalid tableStartIndex: expected non-negative number');
    }
    if (typeof args.rows !== 'number' || args.rows < 1) {
      throw new Error('Invalid rows: expected positive number');
    }
    if (typeof args.columns !== 'number' || args.columns < 1) {
      throw new Error('Invalid columns: expected positive number');
    }
    if (args.columnWidths !== undefined) {
      if (!Array.isArray(args.columnWidths) || args.columnWidths.length !== args.columns) {
        throw new Error(`Invalid columnWidths: expected array of length ${args.columns}`);
      }
    }
    if (args.borders !== undefined && !['NONE', 'ALL'].includes(args.borders)) {
      throw new Error('Invalid borders: expected NONE or ALL');
    }
    if (args.contentAlignment !== undefined && !['TOP', 'MIDDLE', 'BOTTOM'].includes(args.contentAlignment)) {
      throw new Error('Invalid contentAlignment: expected TOP, MIDDLE or BOTTOM');
    }
    return {
      documentId: args.documentId,
      tableStartIndex: args.tableStartIndex,
      rows: args.rows,
      columns: args.columns,
      columnWidths: args.columnWidths,
      headerRowBackgroundColor: args.headerRowBackgroundColor,
      bodyRowBackgroundColor: args.bodyRowBackgroundColor,
      cellPadding: args.cellPadding,
      borders: args.borders,
      contentAlignment: args.contentAlignment,
    };
  }

  private validateSetDocumentDefaultsArgs(args: any): SetDocumentDefaultsRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    return {
      documentId: args.documentId,
      fontFamily: args.fontFamily,
      fontSize: args.fontSize,
      foregroundColor: args.foregroundColor,
      marginTop: args.marginTop,
      marginBottom: args.marginBottom,
      marginLeft: args.marginLeft,
      marginRight: args.marginRight,
    };
  }

  private validateCreateFromTemplateArgs(args: any): CreateFromTemplateRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.templateId || typeof args.templateId !== 'string') {
      throw new Error('Invalid templateId: expected non-empty string');
    }
    if (!args.title || typeof args.title !== 'string') {
      throw new Error('Invalid title: expected non-empty string');
    }
    if (args.replacements !== undefined && (typeof args.replacements !== 'object' || Array.isArray(args.replacements))) {
      throw new Error('Invalid replacements: expected object map of placeholder → value');
    }
    return {
      templateId: args.templateId,
      title: args.title,
      replacements: args.replacements,
      parentFolderId: args.parentFolderId,
    };
  }

  private validateFindTextArgs(args: any): FindTextRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    if (!args.query || typeof args.query !== 'string') {
      throw new Error('Invalid query: expected non-empty string');
    }
    return {
      documentId: args.documentId,
      query: args.query,
      matchCase: args.matchCase !== false,
      firstOnly: args.firstOnly === true,
    };
  }

  private validateFillTableCellArgs(args: any): FillTableCellRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    if (typeof args.tableStartIndex !== 'number' || args.tableStartIndex < 0) {
      throw new Error('Invalid tableStartIndex: expected non-negative number');
    }
    if (typeof args.rowIndex !== 'number' || args.rowIndex < 0) {
      throw new Error('Invalid rowIndex: expected non-negative number');
    }
    if (typeof args.columnIndex !== 'number' || args.columnIndex < 0) {
      throw new Error('Invalid columnIndex: expected non-negative number');
    }
    if (typeof args.text !== 'string') {
      throw new Error('Invalid text: expected string');
    }
    if (args.mode !== undefined && args.mode !== 'replace' && args.mode !== 'append') {
      throw new Error('Invalid mode: expected "replace" or "append"');
    }
    return {
      documentId: args.documentId,
      tableStartIndex: args.tableStartIndex,
      rowIndex: args.rowIndex,
      columnIndex: args.columnIndex,
      text: args.text,
      mode: args.mode || 'replace',
    };
  }

  private validateInsertTableRowArgs(args: any): InsertTableRowRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    if (typeof args.tableStartIndex !== 'number' || args.tableStartIndex < 0) {
      throw new Error('Invalid tableStartIndex: expected non-negative number');
    }
    if (typeof args.rowIndex !== 'number' || args.rowIndex < 0) {
      throw new Error('Invalid rowIndex: expected non-negative number');
    }
    if (args.insertBelow !== undefined && typeof args.insertBelow !== 'boolean') {
      throw new Error('Invalid insertBelow: expected boolean');
    }
    return {
      documentId: args.documentId,
      tableStartIndex: args.tableStartIndex,
      rowIndex: args.rowIndex,
      insertBelow: args.insertBelow !== false, // default true
    };
  }

  private validateDeleteTableRowArgs(args: any): DeleteTableRowRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    if (typeof args.tableStartIndex !== 'number' || args.tableStartIndex < 0) {
      throw new Error('Invalid tableStartIndex: expected non-negative number');
    }
    if (typeof args.rowIndex !== 'number' || args.rowIndex < 0) {
      throw new Error('Invalid rowIndex: expected non-negative number');
    }
    return {
      documentId: args.documentId,
      tableStartIndex: args.tableStartIndex,
      rowIndex: args.rowIndex,
    };
  }

  private validateInsertTableColumnArgs(args: any): InsertTableColumnRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    if (typeof args.tableStartIndex !== 'number' || args.tableStartIndex < 0) {
      throw new Error('Invalid tableStartIndex: expected non-negative number');
    }
    if (typeof args.columnIndex !== 'number' || args.columnIndex < 0) {
      throw new Error('Invalid columnIndex: expected non-negative number');
    }
    if (args.insertRight !== undefined && typeof args.insertRight !== 'boolean') {
      throw new Error('Invalid insertRight: expected boolean');
    }
    return {
      documentId: args.documentId,
      tableStartIndex: args.tableStartIndex,
      columnIndex: args.columnIndex,
      insertRight: args.insertRight !== false, // default true
    };
  }

  private validateDeleteTableColumnArgs(args: any): DeleteTableColumnRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    if (typeof args.tableStartIndex !== 'number' || args.tableStartIndex < 0) {
      throw new Error('Invalid tableStartIndex: expected non-negative number');
    }
    if (typeof args.columnIndex !== 'number' || args.columnIndex < 0) {
      throw new Error('Invalid columnIndex: expected non-negative number');
    }
    return {
      documentId: args.documentId,
      tableStartIndex: args.tableStartIndex,
      columnIndex: args.columnIndex,
    };
  }

  private validateDeleteTableArgs(args: any): DeleteTableRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    if (typeof args.tableStartIndex !== 'number' || args.tableStartIndex < 0) {
      throw new Error('Invalid tableStartIndex: expected non-negative number');
    }
    return {
      documentId: args.documentId,
      tableStartIndex: args.tableStartIndex,
    };
  }

  private validateGetStyleProfileArgs(args: any): GetStyleProfileRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    return { documentId: args.documentId };
  }

  private validateDeleteRangeArgs(args: any): DeleteRangeRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    if (typeof args.startIndex !== 'number' || args.startIndex < 1) {
      throw new Error('Invalid startIndex: expected number >= 1');
    }
    if (typeof args.endIndex !== 'number' || args.endIndex <= args.startIndex) {
      throw new Error('Invalid endIndex: expected number greater than startIndex');
    }
    return { documentId: args.documentId, startIndex: args.startIndex, endIndex: args.endIndex };
  }

  // ── Handlers ───────────────────────────────────────────────────────────────

  /** Get the end-of-document index for insert/append operations. */
  private async getDocEndIndex(documentId: string): Promise<number> {
    const doc = await this.docs.documents.get({ documentId });
    let endIndex = 1;
    if (doc.data.body?.content) {
      for (const element of doc.data.body.content) {
        if (element.endIndex) {
          endIndex = Math.max(endIndex, element.endIndex - 1);
        }
      }
    }
    return endIndex;
  }

  private async handleCreateDocument(args: CreateDocumentRequest): Promise<ToolResponse> {
    try {
      const response = await this.docs.documents.create({
        requestBody: { title: args.title },
      });

      const documentId = response.data.documentId!;

      if (args.content && args.content.trim()) {
        await this.docs.documents.batchUpdate({
          documentId,
          requestBody: {
            requests: [{ insertText: { location: { index: 1 }, text: args.content } }],
          },
        });
      }

      if (args.parentFolderId) {
        const file = await this.drive.files.get({ fileId: documentId, fields: 'parents', supportsAllDrives: true });
        const previousParents = file.data.parents?.join(',');
        await this.drive.files.update({
          fileId: documentId,
          addParents: args.parentFolderId,
          removeParents: previousParents,
          fields: 'id, parents',
          supportsAllDrives: true,
        });
      }

      return successResponse({
        success: true,
        documentId,
        url: `https://docs.google.com/document/d/${documentId}/edit`,
        title: args.title,
        parentFolderId: args.parentFolderId || 'root',
      });
    } catch (error) {
      return errorResponse('creating document', error);
    }
  }

  private async handleGetDocument(args: { documentId: string }): Promise<ToolResponse> {
    try {
      const response = await this.docs.documents.get({
        documentId: args.documentId,
      });
      const doc = response.data;

      let content = '';
      const suggestions: {
        id: string;
        type: 'insertion' | 'deletion' | 'format';
        text?: string;
        context?: string;
      }[] = [];
      const tablePositions: { startIndex: number; rows: number; columns: number; context?: string }[] = [];

      // Track the current section heading for suggestion context
      let currentHeading = '';

      /** Extract text and suggestions from paragraph elements */
      const processParagraph = (elements: docs_v1.Schema$ParagraphElement[], headingText?: string): string => {
        // Update current heading if this paragraph is a heading
        if (headingText !== undefined) {
          currentHeading = headingText;
        }
        let text = '';
        for (const elem of elements) {
          const run = elem.textRun;
          if (!run?.content) continue;

          const runText = run.content;
          const insertionIds = run.suggestedInsertionIds || [];
          const deletionIds = run.suggestedDeletionIds || [];

          if (insertionIds.length > 0) {
            for (const id of insertionIds) {
              suggestions.push({ id, type: 'insertion', text: runText.replace(/\n$/, ''), context: currentHeading });
            }
          } else if (deletionIds.length > 0) {
            for (const id of deletionIds) {
              suggestions.push({ id, type: 'deletion', text: runText.replace(/\n$/, ''), context: currentHeading });
            }
          }

          if (run.suggestedTextStyleChanges) {
            for (const [id] of Object.entries(run.suggestedTextStyleChanges)) {
              const existing = suggestions.find((s) => s.id === id && s.type === 'format');
              if (!existing) {
                suggestions.push({ id, type: 'format', text: runText.replace(/\n$/, ''), context: currentHeading });
              }
            }
          }

          text += runText;
        }
        return text;
      };

      /** Render a table as pipe-delimited rows */
      const processTable = (table: docs_v1.Schema$Table, startIndex?: number): string => {
        const rows: string[][] = [];
        for (const row of table.tableRows || []) {
          const cells: string[] = [];
          for (const cell of row.tableCells || []) {
            let cellText = '';
            for (const cellElement of cell.content || []) {
              if (cellElement.paragraph?.elements) {
                cellText += processParagraph(cellElement.paragraph.elements);
              } else if (cellElement.table) {
                cellText += processTable(cellElement.table);
              }
            }
            // Trim trailing newlines from cell content and collapse internal newlines
            cells.push(cellText.replace(/\n+$/, '').replace(/\n/g, ' '));
          }
          rows.push(cells);
        }

        if (startIndex !== undefined) {
          tablePositions.push({
            startIndex,
            rows: table.rows || rows.length,
            columns: table.columns || (rows[0]?.length ?? 0),
            context: currentHeading || undefined,
          });
        }

        if (rows.length === 0) return '';

        let result = '';
        for (let i = 0; i < rows.length; i++) {
          result += '| ' + rows[i].join(' | ') + ' |\n';
          if (i === 0) {
            result += '| ' + rows[i].map(() => '---').join(' | ') + ' |\n';
          }
        }
        return result;
      };

      /** Process structural elements (paragraphs, tables, etc.) */
      const processStructuralElements = (elements: docs_v1.Schema$StructuralElement[]): string => {
        let text = '';
        for (const element of elements) {
          if (element.paragraph?.elements) {
            const isHeading = element.paragraph.paragraphStyle?.namedStyleType?.startsWith('HEADING');
            // Extract plain text for heading context (before processing suggestions)
            const headingLabel = isHeading
              ? element.paragraph.elements.map((e) => e.textRun?.content || '').join('').trim()
              : undefined;
            text += processParagraph(element.paragraph.elements, headingLabel);
          } else if (element.table) {
            text += processTable(element.table, element.startIndex ?? undefined);
          }
        }
        return text;
      };

      if (doc.body?.content) {
        content = processStructuralElements(doc.body.content);
      }

      const styles = extractDocStyles(doc);
      let styleInfo = '';
      if (styles.dominantStyle) {
        styleInfo += `\n\nDominant body style: ${JSON.stringify(styles.dominantStyle)}`;
      }
      if (styles.headingStyle) {
        styleInfo += `\nHeading style: ${JSON.stringify(styles.headingStyle)}`;
      }
      if (styleInfo) {
        styleInfo = '\n\n--- Detected Styles (match these when adding content) ---' + styleInfo;
      }

      let tablesInfo = '';
      if (tablePositions.length > 0) {
        tablesInfo = `\n\n--- Tables (${tablePositions.length}) ---`;
        tablePositions.forEach((t, i) => {
          const section = t.context ? ` [${t.context}]` : '';
          tablesInfo += `\n${i + 1}. tableStartIndex=${t.startIndex}, ${t.rows}×${t.columns}${section}`;
        });
      }

      // Summarise suggestions if any
      let suggestionsInfo = '';
      if (suggestions.length > 0) {
        // Group by suggestion ID to pair insertions and deletions
        const grouped = new Map<string, { insertions: string[]; deletions: string[]; formats: string[]; context: string }>();
        for (const s of suggestions) {
          if (!grouped.has(s.id)) grouped.set(s.id, { insertions: [], deletions: [], formats: [], context: s.context || '' });
          const group = grouped.get(s.id)!;
          if (s.type === 'insertion' && s.text) group.insertions.push(s.text);
          else if (s.type === 'deletion' && s.text) group.deletions.push(s.text);
          else if (s.type === 'format' && s.text) group.formats.push(s.text);
          // Keep the most specific context (first non-empty wins)
          if (!group.context && s.context) group.context = s.context;
        }

        const insertionCount = [...grouped.values()].filter((g) => g.insertions.length > 0).length;
        const deletionCount = [...grouped.values()].filter((g) => g.deletions.length > 0).length;
        const formatCount = [...grouped.values()].filter((g) => g.formats.length > 0 && g.insertions.length === 0 && g.deletions.length === 0).length;

        suggestionsInfo = `\n\n--- Suggested Changes (${grouped.size} suggestions: ${insertionCount} insertions, ${deletionCount} deletions, ${formatCount} format changes) ---`;

        let i = 1;
        for (const [, group] of grouped) {
          if (group.deletions.length > 0 || group.insertions.length > 0) {
            const section = group.context ? `[${group.context}] ` : '';
            suggestionsInfo += `\n${i}. ${section}`;
            if (group.deletions.length > 0) {
              suggestionsInfo += `DELETE: "${group.deletions.join('')}"`;
            }
            if (group.insertions.length > 0) {
              if (group.deletions.length > 0) suggestionsInfo += ' → ';
              suggestionsInfo += `INSERT: "${group.insertions.join('')}"`;
            }
            i++;
          }
        }
        if (formatCount > 0) {
          suggestionsInfo += `\n(+ ${formatCount} formatting-only changes)`;
        }
      }

      return textResponse(`Document: ${doc.title}\n\nContent:\n${content}${tablesInfo}${suggestionsInfo}${styleInfo}`);
    } catch (error) {
      return errorResponse('getting document', error);
    }
  }

  private async handleInsertText(args: InsertTextRequest): Promise<ToolResponse> {
    try {
      let insertIndex = args.index;
      if (insertIndex === undefined) {
        insertIndex = await this.getDocEndIndex(args.documentId);
      }

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [{ insertText: { location: { index: insertIndex }, text: args.text } }],
        },
      });

      return successResponse({
        success: true,
        insertedAt: insertIndex,
        textLength: args.text.length,
      });
    } catch (error) {
      return errorResponse('inserting text', error);
    }
  }

  private async handleAppendText(args: InsertTextRequest): Promise<ToolResponse> {
    try {
      const endIndex = await this.getDocEndIndex(args.documentId);

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [{ insertText: { location: { index: endIndex }, text: args.text } }],
        },
      });

      return successResponse({
        success: true,
        appendedAt: endIndex,
        textLength: args.text.length,
      });
    } catch (error) {
      return errorResponse('appending text', error);
    }
  }

  private async handleReplaceText(args: ReplaceTextRequest): Promise<ToolResponse> {
    try {
      const response = await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              replaceAllText: {
                containsText: { text: args.find, matchCase: true },
                replaceText: args.replace,
              },
            },
          ],
        },
      });

      const replaceCount = response.data.replies?.[0]?.replaceAllText?.occurrencesChanged || 0;

      return successResponse({
        success: true,
        replacements: replaceCount,
        find: args.find,
        replace: args.replace,
      });
    } catch (error) {
      return errorResponse('replacing text', error);
    }
  }

  private async handleFormatText(args: FormatTextRequest): Promise<ToolResponse> {
    try {
      const textStyle: docs_v1.Schema$TextStyle = {};
      const fields: string[] = [];
      const f = args.format;

      if (f.bold !== undefined) { textStyle.bold = f.bold; fields.push('bold'); }
      if (f.italic !== undefined) { textStyle.italic = f.italic; fields.push('italic'); }
      if (f.underline !== undefined) { textStyle.underline = f.underline; fields.push('underline'); }
      if (f.strikethrough !== undefined) { textStyle.strikethrough = f.strikethrough; fields.push('strikethrough'); }
      if (f.fontSize !== undefined) {
        textStyle.fontSize = { magnitude: f.fontSize, unit: 'PT' };
        fields.push('fontSize');
      }
      if (f.fontFamily !== undefined) {
        textStyle.weightedFontFamily = { fontFamily: f.fontFamily };
        fields.push('weightedFontFamily');
      }
      if (f.foregroundColor !== undefined) {
        textStyle.foregroundColor = { color: { rgbColor: f.foregroundColor } };
        fields.push('foregroundColor');
      }
      if (f.backgroundColor !== undefined) {
        textStyle.backgroundColor = { color: { rgbColor: f.backgroundColor } };
        fields.push('backgroundColor');
      }
      if (f.link !== undefined) {
        textStyle.link = { url: f.link.url };
        fields.push('link');
      }

      if (fields.length === 0) {
        throw new Error('No formatting properties provided in format object');
      }

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              updateTextStyle: {
                range: { startIndex: args.startIndex, endIndex: args.endIndex },
                textStyle,
                fields: fields.join(','),
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        formattedRange: `${args.startIndex}-${args.endIndex}`,
        appliedFormat: args.format,
      });
    } catch (error) {
      return errorResponse('formatting text', error);
    }
  }

  private async handleFormatParagraph(args: FormatParagraphRequest): Promise<ToolResponse> {
    try {
      const paragraphStyle: docs_v1.Schema$ParagraphStyle = {};
      const fields: string[] = [];
      const s = args.style;

      if (s.namedStyleType !== undefined) { paragraphStyle.namedStyleType = s.namedStyleType; fields.push('namedStyleType'); }
      if (s.alignment !== undefined) { paragraphStyle.alignment = s.alignment; fields.push('alignment'); }
      if (s.lineSpacing !== undefined) { paragraphStyle.lineSpacing = s.lineSpacing; fields.push('lineSpacing'); }
      if (s.spaceAbove !== undefined) {
        paragraphStyle.spaceAbove = { magnitude: s.spaceAbove, unit: 'PT' };
        fields.push('spaceAbove');
      }
      if (s.spaceBelow !== undefined) {
        paragraphStyle.spaceBelow = { magnitude: s.spaceBelow, unit: 'PT' };
        fields.push('spaceBelow');
      }
      if (s.indentFirstLine !== undefined) {
        paragraphStyle.indentFirstLine = { magnitude: s.indentFirstLine, unit: 'PT' };
        fields.push('indentFirstLine');
      }
      if (s.indentStart !== undefined) {
        paragraphStyle.indentStart = { magnitude: s.indentStart, unit: 'PT' };
        fields.push('indentStart');
      }
      if (s.indentEnd !== undefined) {
        paragraphStyle.indentEnd = { magnitude: s.indentEnd, unit: 'PT' };
        fields.push('indentEnd');
      }
      if (s.keepWithNext !== undefined) { paragraphStyle.keepWithNext = s.keepWithNext; fields.push('keepWithNext'); }

      if (fields.length === 0) {
        throw new Error('No style properties provided in style object');
      }

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              updateParagraphStyle: {
                range: { startIndex: args.startIndex, endIndex: args.endIndex },
                paragraphStyle,
                fields: fields.join(','),
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        formattedRange: `${args.startIndex}-${args.endIndex}`,
        appliedStyle: args.style,
      });
    } catch (error) {
      return errorResponse('formatting paragraph', error);
    }
  }

  private async handleSetHeading(args: SetHeadingRequest): Promise<ToolResponse> {
    try {
      const namedStyleType = `HEADING_${args.headingLevel}`;

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              updateParagraphStyle: {
                range: { startIndex: args.startIndex, endIndex: args.endIndex },
                paragraphStyle: { namedStyleType },
                fields: 'namedStyleType',
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        headingLevel: args.headingLevel,
        formattedRange: `${args.startIndex}-${args.endIndex}`,
      });
    } catch (error) {
      return errorResponse('setting heading', error);
    }
  }

  private async handleCreateBullets(args: CreateBulletsRequest): Promise<ToolResponse> {
    try {
      const preset = args.preset || 'BULLET_DISC_CIRCLE_SQUARE';

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              createParagraphBullets: {
                range: { startIndex: args.startIndex, endIndex: args.endIndex },
                bulletPreset: preset,
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        range: `${args.startIndex}-${args.endIndex}`,
        preset,
      });
    } catch (error) {
      return errorResponse('creating bullets', error);
    }
  }

  private async handleInsertTable(args: InsertTableRequest): Promise<ToolResponse> {
    try {
      // Step 1: insert the table
      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              insertTable: {
                location: { index: args.index },
                rows: args.rows,
                columns: args.columns,
              },
            },
          ],
        },
      });

      // Step 2: locate the table to get its true startIndex (Docs API may
      // place it 1 position later than the requested index if a newline was inserted first).
      const doc = await this.docs.documents.get({ documentId: args.documentId });
      const tables: { startIndex: number; table: docs_v1.Schema$Table }[] = [];
      for (const element of doc.data.body?.content || []) {
        if (element.table && element.startIndex !== undefined && element.startIndex !== null) {
          tables.push({ startIndex: element.startIndex, table: element.table });
        }
      }

      // Pick the table closest to (and at or after) the requested index — the one we just inserted.
      const inserted = tables
        .filter((t) => t.startIndex >= args.index)
        .sort((a, b) => a.startIndex - b.startIndex)[0];

      if (!inserted) {
        throw new Error('Could not locate the inserted table');
      }

      // Step 3: populate cells if requested.
      if (args.cellContent && args.cellContent.length > 0) {
        // Walk cells in reverse so earlier insertText calls don't shift later cell indices.
        const inserts: docs_v1.Schema$Request[] = [];
        const rows = inserted.table.tableRows || [];
        for (let r = rows.length - 1; r >= 0; r--) {
          const cells = rows[r].tableCells || [];
          for (let c = cells.length - 1; c >= 0; c--) {
            const cellText = args.cellContent[r]?.[c];
            if (!cellText) continue;
            // Insert at the first paragraph's startIndex — this places text before the cell's
            // trailing newline, i.e. inside the cell. Adding +1 would land at the cell's endIndex
            // (outside the cell).
            const firstParagraph = cells[c].content?.[0];
            const cellTextStart = firstParagraph?.startIndex;
            if (cellTextStart === undefined || cellTextStart === null) continue;
            inserts.push({
              insertText: {
                location: { index: cellTextStart },
                text: normaliseText(cellText),
              },
            });
          }
        }

        if (inserts.length > 0) {
          await this.docs.documents.batchUpdate({
            documentId: args.documentId,
            requestBody: { requests: inserts },
          });
        }
      }

      return successResponse({
        success: true,
        tableStartIndex: inserted.startIndex,
        rows: args.rows,
        columns: args.columns,
        populatedCells: args.cellContent
          ? args.cellContent.reduce((n, row) => n + row.filter((c) => c).length, 0)
          : 0,
      });
    } catch (error) {
      return errorResponse('inserting table', error);
    }
  }

  private async handleUpdateTable(args: UpdateTableRequest): Promise<ToolResponse> {
    try {
      const requests: docs_v1.Schema$Request[] = [];
      const tableStartLocation = { index: args.tableStartIndex };

      // Column widths
      if (args.columnWidths && args.columnWidths.length > 0) {
        for (let i = 0; i < args.columnWidths.length; i++) {
          requests.push({
            updateTableColumnProperties: {
              tableStartLocation,
              columnIndices: [i],
              tableColumnProperties: {
                widthType: 'FIXED_WIDTH',
                width: { magnitude: args.columnWidths[i], unit: 'PT' },
              },
              fields: 'widthType,width',
            },
          });
        }
      }

      // Cell-level styling (background, padding, borders, alignment) is applied per cell range.
      const cellStyle: docs_v1.Schema$TableCellStyle = {};
      const cellFields: string[] = [];

      if (args.cellPadding !== undefined) {
        cellStyle.paddingTop = { magnitude: args.cellPadding, unit: 'PT' };
        cellStyle.paddingBottom = { magnitude: args.cellPadding, unit: 'PT' };
        cellStyle.paddingLeft = { magnitude: args.cellPadding, unit: 'PT' };
        cellStyle.paddingRight = { magnitude: args.cellPadding, unit: 'PT' };
        cellFields.push('paddingTop', 'paddingBottom', 'paddingLeft', 'paddingRight');
      }

      if (args.contentAlignment !== undefined) {
        cellStyle.contentAlignment = args.contentAlignment;
        cellFields.push('contentAlignment');
      }

      if (args.borders === 'NONE') {
        const noBorder = { color: { color: {} }, width: { magnitude: 0, unit: 'PT' }, dashStyle: 'SOLID' };
        cellStyle.borderTop = noBorder;
        cellStyle.borderBottom = noBorder;
        cellStyle.borderLeft = noBorder;
        cellStyle.borderRight = noBorder;
        cellFields.push('borderTop', 'borderBottom', 'borderLeft', 'borderRight');
      } else if (args.borders === 'ALL') {
        const grey = {
          color: { color: { rgbColor: { red: 0.8, green: 0.8, blue: 0.8 } } },
          width: { magnitude: 0.75, unit: 'PT' },
          dashStyle: 'SOLID',
        };
        cellStyle.borderTop = grey;
        cellStyle.borderBottom = grey;
        cellStyle.borderLeft = grey;
        cellStyle.borderRight = grey;
        cellFields.push('borderTop', 'borderBottom', 'borderLeft', 'borderRight');
      }

      if (cellFields.length > 0) {
        // Apply to all cells via tableRange covering the whole table.
        requests.push({
          updateTableCellStyle: {
            tableRange: {
              tableCellLocation: {
                tableStartLocation,
                rowIndex: 0,
                columnIndex: 0,
              },
              rowSpan: args.rows,
              columnSpan: args.columns,
            },
            tableCellStyle: cellStyle,
            fields: cellFields.join(','),
          },
        });
      }

      // Header row background (applied after general cell styling so it wins for row 0).
      if (args.headerRowBackgroundColor) {
        requests.push({
          updateTableCellStyle: {
            tableRange: {
              tableCellLocation: {
                tableStartLocation,
                rowIndex: 0,
                columnIndex: 0,
              },
              rowSpan: 1,
              columnSpan: args.columns,
            },
            tableCellStyle: {
              backgroundColor: { color: { rgbColor: args.headerRowBackgroundColor } },
            },
            fields: 'backgroundColor',
          },
        });
      }

      // Body row background (rows 1..N-1).
      if (args.bodyRowBackgroundColor && args.rows > 1) {
        requests.push({
          updateTableCellStyle: {
            tableRange: {
              tableCellLocation: {
                tableStartLocation,
                rowIndex: 1,
                columnIndex: 0,
              },
              rowSpan: args.rows - 1,
              columnSpan: args.columns,
            },
            tableCellStyle: {
              backgroundColor: { color: { rgbColor: args.bodyRowBackgroundColor } },
            },
            fields: 'backgroundColor',
          },
        });
      }

      if (requests.length === 0) {
        throw new Error('No table updates provided — pass at least one of columnWidths, headerRowBackgroundColor, bodyRowBackgroundColor, cellPadding, borders, contentAlignment');
      }

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: { requests },
      });

      return successResponse({
        success: true,
        tableStartIndex: args.tableStartIndex,
        appliedOperations: requests.length,
      });
    } catch (error) {
      return errorResponse('updating table', error);
    }
  }

  /**
   * Find a table by its structural element startIndex. Searches top-level body
   * and recurses into nested tables. Throws if not found.
   */
  private async findTableElementByIndex(
    documentId: string,
    tableStartIndex: number,
  ): Promise<{ table: docs_v1.Schema$Table; element: docs_v1.Schema$StructuralElement }> {
    const doc = await this.docs.documents.get({ documentId });
    let result: { table: docs_v1.Schema$Table; element: docs_v1.Schema$StructuralElement } | undefined;
    const walk = (elements: docs_v1.Schema$StructuralElement[]): void => {
      for (const element of elements) {
        if (result) return;
        if (element.table) {
          if (element.startIndex === tableStartIndex) {
            result = { table: element.table, element };
            return;
          }
          for (const row of element.table.tableRows || []) {
            for (const cell of row.tableCells || []) {
              if (cell.content) walk(cell.content);
            }
          }
        }
      }
    };
    walk(doc.data.body?.content || []);
    if (!result) {
      throw new Error(`No table found at tableStartIndex ${tableStartIndex}`);
    }
    return result;
  }

  /** Validate that the given row/column indices are in bounds for the given table. */
  private assertTableBounds(
    table: docs_v1.Schema$Table,
    rowIndex?: number,
    columnIndex?: number,
  ): void {
    const rowCount = table.tableRows?.length || 0;
    if (rowIndex !== undefined && rowIndex >= rowCount) {
      throw new Error(`rowIndex ${rowIndex} out of bounds (table has ${rowCount} rows)`);
    }
    const columnCount = table.tableRows?.[0]?.tableCells?.length || 0;
    if (columnIndex !== undefined && columnIndex >= columnCount) {
      throw new Error(`columnIndex ${columnIndex} out of bounds (table has ${columnCount} columns)`);
    }
  }

  private async handleFillTableCell(args: FillTableCellRequest): Promise<ToolResponse> {
    try {
      const { table } = await this.findTableElementByIndex(args.documentId, args.tableStartIndex);
      this.assertTableBounds(table, args.rowIndex, args.columnIndex);

      const rows = table.tableRows || [];
      const cells = rows[args.rowIndex].tableCells || [];
      const cell = cells[args.columnIndex];
      const content = cell.content || [];
      if (content.length === 0) {
        throw new Error('Cell has no content paragraphs');
      }

      const firstStart = content[0].startIndex;
      const lastEnd = content[content.length - 1].endIndex;
      if (
        firstStart === undefined ||
        firstStart === null ||
        lastEnd === undefined ||
        lastEnd === null
      ) {
        throw new Error('Cell content indices missing — cannot resolve insert position');
      }

      const text = normaliseText(args.text);
      const mode = args.mode || 'replace';
      const requests: docs_v1.Schema$Request[] = [];

      // The cell's trailing newline sits at lastEnd - 1. We preserve it.
      // Replace: delete [firstStart, lastEnd - 1) (everything except the trailing \n), then insert at firstStart.
      // Append: insert at lastEnd - 1 (just before the trailing \n).
      const trailingNewlineIndex = lastEnd - 1;

      if (mode === 'replace') {
        if (trailingNewlineIndex > firstStart) {
          requests.push({
            deleteContentRange: {
              range: { startIndex: firstStart, endIndex: trailingNewlineIndex },
            },
          });
        }
        if (text.length > 0) {
          requests.push({
            insertText: { location: { index: firstStart }, text },
          });
        }
      } else {
        // append
        if (text.length > 0) {
          requests.push({
            insertText: { location: { index: trailingNewlineIndex }, text },
          });
        }
      }

      if (requests.length === 0) {
        // Empty text in append mode, or empty replace on already-empty cell — no-op
        return successResponse({
          success: true,
          tableStartIndex: args.tableStartIndex,
          rowIndex: args.rowIndex,
          columnIndex: args.columnIndex,
          mode,
          charactersInserted: 0,
          note: 'No-op (empty text or empty cell)',
        });
      }

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: { requests },
      });

      return successResponse({
        success: true,
        tableStartIndex: args.tableStartIndex,
        rowIndex: args.rowIndex,
        columnIndex: args.columnIndex,
        mode,
        charactersInserted: text.length,
      });
    } catch (error) {
      return errorResponse('filling table cell', error);
    }
  }

  private async handleInsertTableRow(args: InsertTableRowRequest): Promise<ToolResponse> {
    try {
      const { table } = await this.findTableElementByIndex(args.documentId, args.tableStartIndex);
      this.assertTableBounds(table, args.rowIndex);

      const insertBelow = args.insertBelow !== false;
      const oldRowCount = table.tableRows?.length || 0;

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              insertTableRow: {
                tableCellLocation: {
                  tableStartLocation: { index: args.tableStartIndex },
                  rowIndex: args.rowIndex,
                  columnIndex: 0,
                },
                insertBelow,
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        tableStartIndex: args.tableStartIndex,
        anchorRowIndex: args.rowIndex,
        insertBelow,
        rowsBefore: oldRowCount,
        rowsAfter: oldRowCount + 1,
      });
    } catch (error) {
      return errorResponse('inserting table row', error);
    }
  }

  private async handleDeleteTableRow(args: DeleteTableRowRequest): Promise<ToolResponse> {
    try {
      const { table } = await this.findTableElementByIndex(args.documentId, args.tableStartIndex);
      this.assertTableBounds(table, args.rowIndex);

      const oldRowCount = table.tableRows?.length || 0;
      if (oldRowCount <= 1) {
        throw new Error(
          'Cannot delete the last row of a table — use gdocs_delete_table to remove the table entirely',
        );
      }

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              deleteTableRow: {
                tableCellLocation: {
                  tableStartLocation: { index: args.tableStartIndex },
                  rowIndex: args.rowIndex,
                  columnIndex: 0,
                },
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        tableStartIndex: args.tableStartIndex,
        deletedRowIndex: args.rowIndex,
        rowsBefore: oldRowCount,
        rowsAfter: oldRowCount - 1,
      });
    } catch (error) {
      return errorResponse('deleting table row', error);
    }
  }

  private async handleInsertTableColumn(args: InsertTableColumnRequest): Promise<ToolResponse> {
    try {
      const { table } = await this.findTableElementByIndex(args.documentId, args.tableStartIndex);
      this.assertTableBounds(table, undefined, args.columnIndex);

      const insertRight = args.insertRight !== false;
      const oldColumnCount = table.tableRows?.[0]?.tableCells?.length || 0;

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              insertTableColumn: {
                tableCellLocation: {
                  tableStartLocation: { index: args.tableStartIndex },
                  rowIndex: 0,
                  columnIndex: args.columnIndex,
                },
                insertRight,
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        tableStartIndex: args.tableStartIndex,
        anchorColumnIndex: args.columnIndex,
        insertRight,
        columnsBefore: oldColumnCount,
        columnsAfter: oldColumnCount + 1,
      });
    } catch (error) {
      return errorResponse('inserting table column', error);
    }
  }

  private async handleDeleteTableColumn(args: DeleteTableColumnRequest): Promise<ToolResponse> {
    try {
      const { table } = await this.findTableElementByIndex(args.documentId, args.tableStartIndex);
      this.assertTableBounds(table, undefined, args.columnIndex);

      const oldColumnCount = table.tableRows?.[0]?.tableCells?.length || 0;
      if (oldColumnCount <= 1) {
        throw new Error(
          'Cannot delete the last column of a table — use gdocs_delete_table to remove the table entirely',
        );
      }

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              deleteTableColumn: {
                tableCellLocation: {
                  tableStartLocation: { index: args.tableStartIndex },
                  rowIndex: 0,
                  columnIndex: args.columnIndex,
                },
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        tableStartIndex: args.tableStartIndex,
        deletedColumnIndex: args.columnIndex,
        columnsBefore: oldColumnCount,
        columnsAfter: oldColumnCount - 1,
      });
    } catch (error) {
      return errorResponse('deleting table column', error);
    }
  }

  private async handleDeleteTable(args: DeleteTableRequest): Promise<ToolResponse> {
    try {
      const { element } = await this.findTableElementByIndex(args.documentId, args.tableStartIndex);
      const startIndex = element.startIndex;
      const endIndex = element.endIndex;
      if (
        startIndex === undefined ||
        startIndex === null ||
        endIndex === undefined ||
        endIndex === null
      ) {
        throw new Error('Table element indices missing — cannot delete');
      }

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              deleteContentRange: {
                range: { startIndex, endIndex },
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        deletedRange: `${startIndex}-${endIndex}`,
        charactersDeleted: endIndex - startIndex,
      });
    } catch (error) {
      return errorResponse('deleting table', error);
    }
  }

  private async handleSetDocumentDefaults(args: SetDocumentDefaultsRequest): Promise<ToolResponse> {
    try {
      const requests: docs_v1.Schema$Request[] = [];

      // Page margins via updateDocumentStyle
      const documentStyle: docs_v1.Schema$DocumentStyle = {};
      const docFields: string[] = [];
      if (args.marginTop !== undefined) {
        documentStyle.marginTop = { magnitude: args.marginTop, unit: 'PT' };
        docFields.push('marginTop');
      }
      if (args.marginBottom !== undefined) {
        documentStyle.marginBottom = { magnitude: args.marginBottom, unit: 'PT' };
        docFields.push('marginBottom');
      }
      if (args.marginLeft !== undefined) {
        documentStyle.marginLeft = { magnitude: args.marginLeft, unit: 'PT' };
        docFields.push('marginLeft');
      }
      if (args.marginRight !== undefined) {
        documentStyle.marginRight = { magnitude: args.marginRight, unit: 'PT' };
        docFields.push('marginRight');
      }
      if (docFields.length > 0) {
        requests.push({
          updateDocumentStyle: {
            documentStyle,
            fields: docFields.join(','),
          },
        });
      }

      // Default body text style — apply to the whole body range
      const textStyle: docs_v1.Schema$TextStyle = {};
      const textFields: string[] = [];
      if (args.fontFamily !== undefined) {
        textStyle.weightedFontFamily = { fontFamily: args.fontFamily };
        textFields.push('weightedFontFamily');
      }
      if (args.fontSize !== undefined) {
        textStyle.fontSize = { magnitude: args.fontSize, unit: 'PT' };
        textFields.push('fontSize');
      }
      if (args.foregroundColor !== undefined) {
        textStyle.foregroundColor = { color: { rgbColor: args.foregroundColor } };
        textFields.push('foregroundColor');
      }

      if (textFields.length > 0) {
        // Apply textStyle only to NORMAL_TEXT paragraphs (including inside table cells)
        // so we don't trample heading/title font sizes from their named styles.
        const doc = await this.docs.documents.get({ documentId: args.documentId });

        const collectNormalTextRanges = (
          elements: docs_v1.Schema$StructuralElement[],
          out: { startIndex: number; endIndex: number }[],
        ) => {
          for (const element of elements) {
            if (element.paragraph) {
              const ns = element.paragraph.paragraphStyle?.namedStyleType || 'NORMAL_TEXT';
              if (
                ns === 'NORMAL_TEXT' &&
                element.startIndex !== undefined &&
                element.startIndex !== null &&
                element.endIndex !== undefined &&
                element.endIndex !== null
              ) {
                out.push({ startIndex: element.startIndex, endIndex: element.endIndex });
              }
            } else if (element.table) {
              for (const row of element.table.tableRows || []) {
                for (const cell of row.tableCells || []) {
                  if (cell.content) collectNormalTextRanges(cell.content, out);
                }
              }
            }
          }
        };

        const ranges: { startIndex: number; endIndex: number }[] = [];
        collectNormalTextRanges(doc.data.body?.content || [], ranges);

        for (const r of ranges) {
          requests.push({
            updateTextStyle: {
              range: r,
              textStyle,
              fields: textFields.join(','),
            },
          });
        }
      }

      if (requests.length === 0) {
        throw new Error('No defaults provided — pass at least one of fontFamily, fontSize, foregroundColor, margins');
      }

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: { requests },
      });

      return successResponse({
        success: true,
        appliedOperations: requests.length,
      });
    } catch (error) {
      return errorResponse('setting document defaults', error);
    }
  }

  private async handleCreateFromTemplate(args: CreateFromTemplateRequest): Promise<ToolResponse> {
    try {
      // Step 1: copy the template via Drive API
      const copyBody: drive_v3.Schema$File = { name: args.title };
      if (args.parentFolderId) copyBody.parents = [args.parentFolderId];

      const copyResponse = await this.drive.files.copy({
        fileId: args.templateId,
        requestBody: copyBody,
        fields: 'id,name,parents',
        supportsAllDrives: true,
      });

      const newDocId = copyResponse.data.id;
      if (!newDocId) throw new Error('Drive copy did not return a file ID');

      // Step 2: apply replacements (if any) in a single batchUpdate
      let replacementsApplied = 0;
      if (args.replacements && Object.keys(args.replacements).length > 0) {
        const requests: docs_v1.Schema$Request[] = Object.entries(args.replacements).map(
          ([find, replace]) => ({
            replaceAllText: {
              containsText: { text: find, matchCase: true },
              replaceText: normaliseText(replace),
            },
          }),
        );

        const response = await this.docs.documents.batchUpdate({
          documentId: newDocId,
          requestBody: { requests },
        });

        for (const reply of response.data.replies || []) {
          replacementsApplied += reply.replaceAllText?.occurrencesChanged || 0;
        }
      }

      return successResponse({
        success: true,
        documentId: newDocId,
        url: `https://docs.google.com/document/d/${newDocId}/edit`,
        title: args.title,
        parentFolderId: args.parentFolderId || 'root',
        replacementsApplied,
      });
    } catch (error) {
      return errorResponse('creating document from template', error);
    }
  }

  private async handleFindText(args: FindTextRequest): Promise<ToolResponse> {
    try {
      const doc = await this.docs.documents.get({ documentId: args.documentId });
      const matchCase = args.matchCase !== false;
      const needle = matchCase ? args.query : args.query.toLowerCase();

      type Match = {
        startIndex: number;
        endIndex: number;
        paragraphStartIndex: number;
        paragraphEndIndex: number;
        inTable: boolean;
        tableContext?: { tableStartIndex: number; rowIndex: number; columnIndex: number };
        sectionHeading?: string;
        matchedText: string;
      };

      const matches: Match[] = [];
      let currentHeading = '';
      let done = false;

      /** Search a paragraph's text runs for matches. */
      const searchParagraph = (
        para: docs_v1.Schema$Paragraph,
        paraStart: number,
        paraEnd: number,
        tableContext?: { tableStartIndex: number; rowIndex: number; columnIndex: number },
      ) => {
        if (done) return;
        // Concatenate runs while building an index map: paraText[i] → doc index
        let paraText = '';
        const paraIndices: number[] = [];
        for (const elem of para.elements || []) {
          const run = elem.textRun;
          if (!run?.content || elem.startIndex === undefined || elem.startIndex === null) continue;
          for (let i = 0; i < run.content.length; i++) {
            paraIndices.push(elem.startIndex + i);
          }
          paraText += run.content;
        }
        if (paraText.length === 0) return;

        const haystack = matchCase ? paraText : paraText.toLowerCase();
        let pos = 0;
        while ((pos = haystack.indexOf(needle, pos)) !== -1) {
          const start = paraIndices[pos];
          const end = paraIndices[pos + needle.length - 1] + 1;
          matches.push({
            startIndex: start,
            endIndex: end,
            paragraphStartIndex: paraStart,
            paragraphEndIndex: paraEnd,
            inTable: tableContext !== undefined,
            tableContext,
            sectionHeading: currentHeading || undefined,
            matchedText: paraText.substring(pos, pos + needle.length),
          });
          if (args.firstOnly) {
            done = true;
            return;
          }
          pos += needle.length;
        }
      };

      const walk = (elements: docs_v1.Schema$StructuralElement[]) => {
        for (const element of elements) {
          if (done) return;
          if (element.paragraph?.elements) {
            const para = element.paragraph;
            const paraStart = element.startIndex ?? 0;
            const paraEnd = element.endIndex ?? 0;
            const isHeading = para.paragraphStyle?.namedStyleType?.startsWith('HEADING');
            if (isHeading) {
              currentHeading = (para.elements || [])
                .map((e) => e.textRun?.content || '')
                .join('')
                .trim();
            }
            searchParagraph(para, paraStart, paraEnd);
          } else if (element.table) {
            const tableStart = element.startIndex ?? 0;
            const rows = element.table.tableRows || [];
            for (let r = 0; r < rows.length; r++) {
              const cells = rows[r].tableCells || [];
              for (let c = 0; c < cells.length; c++) {
                const cell = cells[c];
                for (const cellElement of cell.content || []) {
                  if (done) return;
                  if (cellElement.paragraph?.elements) {
                    searchParagraph(
                      cellElement.paragraph,
                      cellElement.startIndex ?? 0,
                      cellElement.endIndex ?? 0,
                      { tableStartIndex: tableStart, rowIndex: r, columnIndex: c },
                    );
                  }
                }
              }
            }
          }
        }
      };

      walk(doc.data.body?.content || []);

      return successResponse({
        success: true,
        query: args.query,
        matchCount: matches.length,
        matches,
      });
    } catch (error) {
      return errorResponse('finding text', error);
    }
  }

  private async handleGetStyleProfile(args: GetStyleProfileRequest): Promise<ToolResponse> {
    try {
      const doc = (await this.docs.documents.get({ documentId: args.documentId })).data;

      // ── Document defaults (page setup) ──────────────────────────────
      const documentDefaults: Record<string, any> = {};
      const ds = doc.documentStyle;
      if (ds) {
        const dim = (d: docs_v1.Schema$Dimension | undefined | null) =>
          d?.magnitude !== undefined && d.magnitude !== null ? d.magnitude : undefined;
        documentDefaults.margins = {
          top: dim(ds.marginTop),
          bottom: dim(ds.marginBottom),
          left: dim(ds.marginLeft),
          right: dim(ds.marginRight),
        };
        if (ds.pageSize) {
          documentDefaults.pageSize = {
            width: dim(ds.pageSize.width),
            height: dim(ds.pageSize.height),
          };
        }
      }

      // ── Named styles ────────────────────────────────────────────────
      const namedStyles: Record<string, any> = {};
      const dim = (d: docs_v1.Schema$Dimension | undefined | null) =>
        d?.magnitude !== undefined && d.magnitude !== null ? d.magnitude : undefined;

      const summariseTextStyle = (ts: docs_v1.Schema$TextStyle | undefined | null) => {
        if (!ts) return undefined;
        const result: Record<string, any> = {};
        if (ts.weightedFontFamily?.fontFamily) result.fontFamily = ts.weightedFontFamily.fontFamily;
        if (ts.fontSize?.magnitude) result.fontSize = ts.fontSize.magnitude;
        if (ts.bold !== undefined && ts.bold !== null) result.bold = ts.bold;
        if (ts.italic !== undefined && ts.italic !== null) result.italic = ts.italic;
        if (ts.underline !== undefined && ts.underline !== null) result.underline = ts.underline;
        const fg = ts.foregroundColor?.color?.rgbColor;
        if (fg) result.foregroundColor = {
          red: Math.round((fg.red || 0) * 1000) / 1000,
          green: Math.round((fg.green || 0) * 1000) / 1000,
          blue: Math.round((fg.blue || 0) * 1000) / 1000,
        };
        const bg = ts.backgroundColor?.color?.rgbColor;
        if (bg) result.backgroundColor = {
          red: Math.round((bg.red || 0) * 1000) / 1000,
          green: Math.round((bg.green || 0) * 1000) / 1000,
          blue: Math.round((bg.blue || 0) * 1000) / 1000,
        };
        return Object.keys(result).length > 0 ? result : undefined;
      };

      const summariseParagraphStyle = (ps: docs_v1.Schema$ParagraphStyle | undefined | null) => {
        if (!ps) return undefined;
        const result: Record<string, any> = {};
        if (ps.alignment) result.alignment = ps.alignment;
        if (ps.lineSpacing) result.lineSpacing = ps.lineSpacing;
        const sa = dim(ps.spaceAbove); if (sa !== undefined) result.spaceAbove = sa;
        const sb = dim(ps.spaceBelow); if (sb !== undefined) result.spaceBelow = sb;
        const ifl = dim(ps.indentFirstLine); if (ifl !== undefined) result.indentFirstLine = ifl;
        const is_ = dim(ps.indentStart); if (is_ !== undefined) result.indentStart = is_;
        const ie = dim(ps.indentEnd); if (ie !== undefined) result.indentEnd = ie;
        if (ps.keepWithNext !== undefined && ps.keepWithNext !== null) result.keepWithNext = ps.keepWithNext;
        return Object.keys(result).length > 0 ? result : undefined;
      };

      for (const style of doc.namedStyles?.styles || []) {
        if (!style.namedStyleType) continue;
        const entry: Record<string, any> = {};
        const ts = summariseTextStyle(style.textStyle);
        const ps = summariseParagraphStyle(style.paragraphStyle);
        if (ts) entry.textStyle = ts;
        if (ps) entry.paragraphStyle = ps;
        if (Object.keys(entry).length > 0) {
          namedStyles[style.namedStyleType] = entry;
        }
      }

      // ── Dominant body style (reusing existing extractDocStyles) ────
      const styles = extractDocStyles(doc);

      // ── Walk body for: tables, suspected headings ──────────────────
      const tablePatterns: Array<Record<string, any>> = [];
      const suspectedHeadings: Array<Record<string, any>> = [];
      const bodyFontSize = styles.dominantStyle?.fontSize || 11;

      const summariseCellStyle = (cs: docs_v1.Schema$TableCellStyle | undefined | null) => {
        if (!cs) return undefined;
        const result: Record<string, any> = {};
        const bg = cs.backgroundColor?.color?.rgbColor;
        if (bg) result.backgroundColor = {
          red: Math.round((bg.red || 0) * 1000) / 1000,
          green: Math.round((bg.green || 0) * 1000) / 1000,
          blue: Math.round((bg.blue || 0) * 1000) / 1000,
        };
        const padTop = dim(cs.paddingTop);
        const padBottom = dim(cs.paddingBottom);
        const padLeft = dim(cs.paddingLeft);
        const padRight = dim(cs.paddingRight);
        if (padTop !== undefined || padBottom !== undefined || padLeft !== undefined || padRight !== undefined) {
          if (padTop === padBottom && padTop === padLeft && padTop === padRight) {
            result.padding = padTop;
          } else {
            result.padding = { top: padTop, bottom: padBottom, left: padLeft, right: padRight };
          }
        }
        if (cs.contentAlignment) result.contentAlignment = cs.contentAlignment;
        const borderTop = cs.borderTop;
        if (borderTop?.width?.magnitude === 0) {
          result.borders = 'NONE';
        } else if (borderTop?.width?.magnitude) {
          result.borders = 'ALL';
        }
        return Object.keys(result).length > 0 ? result : undefined;
      };

      const summariseTable = (table: docs_v1.Schema$Table, startIndex: number) => {
        const rows = table.tableRows || [];
        const columnCount = table.columns || (rows[0]?.tableCells?.length ?? 0);
        const columnWidths: number[] = [];
        for (const colProp of table.tableStyle?.tableColumnProperties || []) {
          const w = dim(colProp.width);
          if (w !== undefined) columnWidths.push(w);
        }

        // Sample cell style from the first cell (representative)
        const firstCellStyle = rows[0]?.tableCells?.[0]?.tableCellStyle;
        const firstRowFirstCell = summariseCellStyle(firstCellStyle);

        // Header vs body: compare first row's first cell background to second row's
        let headerRowBackgroundColor;
        let bodyRowBackgroundColor;
        const row0Bg = rows[0]?.tableCells?.[0]?.tableCellStyle?.backgroundColor?.color?.rgbColor;
        const row1Bg = rows[1]?.tableCells?.[0]?.tableCellStyle?.backgroundColor?.color?.rgbColor;
        if (row0Bg) {
          headerRowBackgroundColor = {
            red: Math.round((row0Bg.red || 0) * 1000) / 1000,
            green: Math.round((row0Bg.green || 0) * 1000) / 1000,
            blue: Math.round((row0Bg.blue || 0) * 1000) / 1000,
          };
        }
        if (row1Bg) {
          bodyRowBackgroundColor = {
            red: Math.round((row1Bg.red || 0) * 1000) / 1000,
            green: Math.round((row1Bg.green || 0) * 1000) / 1000,
            blue: Math.round((row1Bg.blue || 0) * 1000) / 1000,
          };
        }

        const pattern: Record<string, any> = {
          tableStartIndex: startIndex,
          rows: table.rows || rows.length,
          columns: columnCount,
        };
        if (columnWidths.length > 0) pattern.columnWidths = columnWidths;
        if (headerRowBackgroundColor) pattern.headerRowBackgroundColor = headerRowBackgroundColor;
        if (bodyRowBackgroundColor) pattern.bodyRowBackgroundColor = bodyRowBackgroundColor;
        if (firstRowFirstCell?.padding !== undefined) pattern.cellPadding = firstRowFirstCell.padding;
        if (firstRowFirstCell?.contentAlignment) pattern.contentAlignment = firstRowFirstCell.contentAlignment;
        if (firstRowFirstCell?.borders) pattern.borders = firstRowFirstCell.borders;

        // Classify purpose heuristically
        if (columnCount === 2 && !headerRowBackgroundColor && firstRowFirstCell?.borders === 'NONE') {
          pattern.purpose = 'field-list-borderless';
        } else if (columnCount === 2 && !headerRowBackgroundColor) {
          pattern.purpose = 'field-list';
        } else if (headerRowBackgroundColor) {
          pattern.purpose = 'data-table';
        }

        return pattern;
      };

      const detectSuspectedHeading = (
        para: docs_v1.Schema$Paragraph,
        startIndex: number,
        endIndex: number,
      ) => {
        const ns = para.paragraphStyle?.namedStyleType;
        if (ns && ns !== 'NORMAL_TEXT') return; // already a heading

        const text = (para.elements || [])
          .map((e) => e.textRun?.content || '')
          .join('')
          .replace(/\n+$/, '')
          .trim();

        if (text.length === 0 || text.length > 80) return;
        if (text.endsWith('.') || text.endsWith(',') || text.endsWith(':')) {
          // colons sometimes used for headings — only reject if there's a clear sentence-end
          if (text.endsWith('.') || text.endsWith(',')) return;
        }

        // Analyse text run styles weighted by character count
        let boldChars = 0;
        let totalChars = 0;
        let weightedSize = 0;
        for (const elem of para.elements || []) {
          const run = elem.textRun;
          if (!run?.content || !run.textStyle) continue;
          const len = run.content.replace(/\n/g, '').length;
          totalChars += len;
          if (run.textStyle.bold) boldChars += len;
          if (run.textStyle.fontSize?.magnitude) {
            weightedSize += run.textStyle.fontSize.magnitude * len;
          }
        }
        if (totalChars === 0) return;
        const avgSize = weightedSize / totalChars;
        const mostlyBold = boldChars / totalChars > 0.6;
        const largerThanBody = avgSize > bodyFontSize * 1.2;
        const allCaps = text === text.toUpperCase() && /[A-Z]/.test(text);

        const reasons: string[] = [];
        if (mostlyBold) reasons.push('mostly bold');
        if (largerThanBody) reasons.push(`larger than body (${avgSize.toFixed(1)}pt vs ${bodyFontSize}pt)`);
        if (allCaps) reasons.push('ALL CAPS');

        if (reasons.length >= 2 || (reasons.length === 1 && largerThanBody)) {
          // Suggest a level based on size and presence of cues
          let suggested = 'HEADING_2';
          if (avgSize >= bodyFontSize * 1.8 || (allCaps && largerThanBody)) suggested = 'HEADING_1';
          if (avgSize >= bodyFontSize * 2.2) suggested = 'TITLE';

          suspectedHeadings.push({
            text,
            startIndex,
            endIndex,
            currentNamedStyleType: ns || 'NORMAL_TEXT',
            suggestedNamedStyleType: suggested,
            reason: reasons.join('; '),
          });
        }
      };

      const walkBody = (elements: docs_v1.Schema$StructuralElement[]) => {
        for (const element of elements) {
          if (element.paragraph && element.startIndex !== undefined && element.startIndex !== null) {
            detectSuspectedHeading(
              element.paragraph,
              element.startIndex,
              element.endIndex ?? element.startIndex,
            );
          } else if (element.table && element.startIndex !== undefined && element.startIndex !== null) {
            tablePatterns.push(summariseTable(element.table, element.startIndex));
            // Recurse into cells (in case of nested tables, though rare)
            for (const row of element.table.tableRows || []) {
              for (const cell of row.tableCells || []) {
                if (cell.content) walkBody(cell.content);
              }
            }
          }
        }
      };

      walkBody(doc.body?.content || []);

      const profile = {
        documentId: args.documentId,
        title: doc.title || '',
        extractedAt: new Date().toISOString(),
        documentDefaults,
        namedStyles,
        dominantBodyStyle: styles.dominantStyle || null,
        tablePatterns,
        suspectedHeadings,
      };

      return successResponse(profile);
    } catch (error) {
      return errorResponse('extracting style profile', error);
    }
  }

  private async handleDeleteRange(args: DeleteRangeRequest): Promise<ToolResponse> {
    try {
      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              deleteContentRange: {
                range: { startIndex: args.startIndex, endIndex: args.endIndex },
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        deletedRange: `${args.startIndex}-${args.endIndex}`,
        charactersDeleted: args.endIndex - args.startIndex,
      });
    } catch (error) {
      return errorResponse('deleting range', error);
    }
  }
}
