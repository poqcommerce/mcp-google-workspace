import { slides_v1, drive_v3 } from 'googleapis';
import type { ToolDefinition, ToolResponse } from '../types.js';
import { successResponse, textResponse, errorResponse, normaliseText } from '../utils.js';

// ── Tool definitions ───────────────────────────────────────────────────────────

const PREDEFINED_LAYOUTS = [
  'BLANK',
  'CAPTION_ONLY',
  'TITLE',
  'TITLE_AND_BODY',
  'TITLE_AND_TWO_COLUMNS',
  'TITLE_ONLY',
  'SECTION_HEADER',
  'SECTION_TITLE_AND_DESCRIPTION',
  'ONE_COLUMN_TEXT',
  'MAIN_POINT',
  'BIG_NUMBER',
];

export function getSlidesToolDefinitions(): ToolDefinition[] {
  return [
    // ── Phase 4: Foundation ──────────────────────────────────────────────────
    {
      name: 'gslides_create_presentation',
      description: 'Create a new Google Slides presentation',
      inputSchema: {
        type: 'object',
        properties: {
          title: { type: 'string', description: 'Title for the new presentation' },
          parentFolderId: {
            type: 'string',
            description: 'ID of the parent folder (optional, defaults to root)',
          },
        },
        required: ['title'],
      },
    },
    {
      name: 'gslides_get_presentation',
      description: 'Get presentation structure, slide content, and speaker notes',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          includeText: {
            type: 'boolean',
            description: 'Extract text content from all slides (default: true)',
          },
        },
        required: ['presentationId'],
      },
    },

    // ── Phase 5: Content ─────────────────────────────────────────────────────
    {
      name: 'gslides_add_slide',
      description: 'Add a new slide to a presentation',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          layout: {
            type: 'string',
            description: `Predefined layout: ${PREDEFINED_LAYOUTS.join(', ')}`,
            enum: PREDEFINED_LAYOUTS,
          },
          insertionIndex: {
            type: 'number',
            description: 'Position to insert the slide (0-based, optional, defaults to end)',
          },
        },
        required: ['presentationId'],
      },
    },
    {
      name: 'gslides_insert_text',
      description: 'Insert text into a shape or placeholder on a slide',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          objectId: {
            type: 'string',
            description: 'Object ID of the shape/placeholder (from gslides_get_presentation)',
          },
          text: { type: 'string', description: 'Text to insert' },
          insertionIndex: {
            type: 'number',
            description: 'Character index to insert at (default: 0)',
          },
        },
        required: ['presentationId', 'objectId', 'text'],
      },
    },
    {
      name: 'gslides_replace_text',
      description: 'Find and replace text across the entire presentation (ideal for template filling)',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          find: { type: 'string', description: 'Text to find' },
          replace: { type: 'string', description: 'Text to replace with' },
          matchCase: { type: 'boolean', description: 'Case-sensitive match (default: true)' },
          pageObjectIds: {
            type: 'array',
            description: 'Limit replacement to specific slides (optional)',
            items: { type: 'string' },
          },
        },
        required: ['presentationId', 'find', 'replace'],
      },
    },
    {
      name: 'gslides_speaker_notes',
      description: 'Read or write speaker notes on a slide',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          slideObjectId: { type: 'string', description: 'Object ID of the slide' },
          notes: {
            type: 'string',
            description: 'Speaker notes to set. If omitted, reads existing notes.',
          },
        },
        required: ['presentationId', 'slideObjectId'],
      },
    },

    // ── Phase 6: Visual elements ─────────────────────────────────────────────
    {
      name: 'gslides_add_shape',
      description: 'Add a shape (rectangle, ellipse, text box, etc.) to a slide',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          pageObjectId: { type: 'string', description: 'Object ID of the slide to add the shape to' },
          shapeType: {
            type: 'string',
            description: 'Shape type (e.g., RECTANGLE, ELLIPSE, TEXT_BOX, ROUND_RECTANGLE, CLOUD, ARROW_EAST)',
          },
          x: { type: 'number', description: 'Left edge position in points' },
          y: { type: 'number', description: 'Top edge position in points' },
          width: { type: 'number', description: 'Width in points' },
          height: { type: 'number', description: 'Height in points' },
        },
        required: ['presentationId', 'pageObjectId', 'shapeType', 'x', 'y', 'width', 'height'],
      },
    },
    {
      name: 'gslides_add_image',
      description: 'Insert an image into a slide from a URL',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          pageObjectId: { type: 'string', description: 'Object ID of the slide' },
          imageUrl: {
            type: 'string',
            description: 'Publicly accessible image URL',
          },
          x: { type: 'number', description: 'Left edge position in points' },
          y: { type: 'number', description: 'Top edge position in points' },
          width: { type: 'number', description: 'Width in points' },
          height: { type: 'number', description: 'Height in points' },
        },
        required: ['presentationId', 'pageObjectId', 'imageUrl', 'x', 'y', 'width', 'height'],
      },
    },
    {
      name: 'gslides_add_table',
      description: 'Create a table on a slide',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          pageObjectId: { type: 'string', description: 'Object ID of the slide' },
          rows: { type: 'number', description: 'Number of rows' },
          columns: { type: 'number', description: 'Number of columns' },
          x: { type: 'number', description: 'Left edge position in points (default: 100)' },
          y: { type: 'number', description: 'Top edge position in points (default: 100)' },
          width: { type: 'number', description: 'Width in points (default: 600)' },
          height: { type: 'number', description: 'Height in points (default: 300)' },
        },
        required: ['presentationId', 'pageObjectId', 'rows', 'columns'],
      },
    },
    {
      name: 'gslides_format_text',
      description: 'Apply formatting to text in a shape or table cell',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          objectId: { type: 'string', description: 'Object ID of the shape or table' },
          startIndex: { type: 'number', description: 'Start character index' },
          endIndex: { type: 'number', description: 'End character index' },
          rowIndex: { type: 'number', description: 'Table row index (optional, for table cells)' },
          columnIndex: { type: 'number', description: 'Table column index (optional, for table cells)' },
          format: {
            type: 'object',
            description: 'Formatting options',
            properties: {
              bold: { type: 'boolean' },
              italic: { type: 'boolean' },
              underline: { type: 'boolean' },
              fontSize: { type: 'number', description: 'Font size in points' },
              fontFamily: { type: 'string' },
              foregroundColor: {
                type: 'object',
                description: '{ red, green, blue } values 0-1',
                properties: {
                  red: { type: 'number' },
                  green: { type: 'number' },
                  blue: { type: 'number' },
                },
              },
            },
          },
        },
        required: ['presentationId', 'objectId', 'format'],
      },
    },

    // ── Phase 10: Batch operations ──────────────────────────────────────
    {
      name: 'gslides_batch_format_text',
      description:
        'Apply formatting to multiple text targets (shapes and/or table cells) in a single API call. ' +
        'Much faster than calling gslides_format_text repeatedly.',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          operations: {
            type: 'array',
            description: 'Array of formatting operations',
            items: {
              type: 'object',
              properties: {
                objectId: { type: 'string', description: 'Object ID of the shape or table' },
                startIndex: { type: 'number', description: 'Start character index (optional, omit for ALL)' },
                endIndex: { type: 'number', description: 'End character index (optional, omit for ALL)' },
                rowIndex: { type: 'number', description: 'Table row index (optional, for table cells)' },
                columnIndex: { type: 'number', description: 'Table column index (optional, for table cells)' },
                format: {
                  type: 'object',
                  description: 'Formatting options',
                  properties: {
                    bold: { type: 'boolean' },
                    italic: { type: 'boolean' },
                    underline: { type: 'boolean' },
                    fontSize: { type: 'number', description: 'Font size in points' },
                    fontFamily: { type: 'string' },
                    foregroundColor: {
                      type: 'object',
                      description: '{ red, green, blue } values 0-1',
                      properties: {
                        red: { type: 'number' },
                        green: { type: 'number' },
                        blue: { type: 'number' },
                      },
                    },
                  },
                },
              },
              required: ['objectId', 'format'],
            },
          },
        },
        required: ['presentationId', 'operations'],
      },
    },

    {
      name: 'gslides_format_table',
      description:
        'Format a table: set column widths, row heights, cell background colors, and content alignment. ' +
        'All properties are optional — include only what you want to change.',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          tableObjectId: { type: 'string', description: 'Object ID of the table' },
          columnWidths: {
            type: 'array',
            description: 'Array of column widths in points (e.g. [150, 450] for a 2-column table). Must match column count.',
            items: { type: 'number' },
          },
          rowHeights: {
            type: 'array',
            description: 'Array of minimum row heights in points. Use null/0 entries to skip rows.',
            items: { type: 'number' },
          },
          contentAlignment: {
            type: 'string',
            description: 'Vertical content alignment for all cells: TOP, MIDDLE, or BOTTOM',
            enum: ['TOP', 'MIDDLE', 'BOTTOM'],
          },
          headerRowBackgroundColor: {
            type: 'object',
            description: 'Background color for the first row (header) as { red, green, blue } values 0-1',
            properties: {
              red: { type: 'number' },
              green: { type: 'number' },
              blue: { type: 'number' },
            },
          },
          bodyRowBackgroundColor: {
            type: 'object',
            description: 'Background color for non-header rows as { red, green, blue } values 0-1',
            properties: {
              red: { type: 'number' },
              green: { type: 'number' },
              blue: { type: 'number' },
            },
          },
          totalRows: {
            type: 'number',
            description: 'Total number of rows in the table (needed for background color operations)',
          },
          totalColumns: {
            type: 'number',
            description: 'Total number of columns in the table (needed for background color operations)',
          },
        },
        required: ['presentationId', 'tableObjectId'],
      },
    },

    // ── Phase 7: Utilities ───────────────────────────────────────────────────
    {
      name: 'gslides_get_thumbnail',
      description: 'Get a PNG thumbnail URL for a slide',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          slideObjectId: { type: 'string', description: 'Object ID of the slide' },
          thumbnailSize: {
            type: 'string',
            description: 'Thumbnail size: SMALL, MEDIUM (default), or LARGE',
            enum: ['SMALL', 'MEDIUM', 'LARGE'],
          },
        },
        required: ['presentationId', 'slideObjectId'],
      },
    },
    {
      name: 'gslides_delete_slide',
      description: 'Delete a slide from a presentation',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          slideObjectId: { type: 'string', description: 'Object ID of the slide to delete' },
        },
        required: ['presentationId', 'slideObjectId'],
      },
    },

    // ── Phase 8: Tables & bullets ───────────────────────────────────────────
    {
      name: 'gslides_insert_table_text',
      description: 'Insert text into a specific table cell by row and column index',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          tableObjectId: { type: 'string', description: 'Object ID of the table (from gslides_get_presentation)' },
          rowIndex: { type: 'number', description: 'Row index (0-based)' },
          columnIndex: { type: 'number', description: 'Column index (0-based)' },
          text: { type: 'string', description: 'Text to insert' },
          insertionIndex: {
            type: 'number',
            description: 'Character index within the cell to insert at (default: 0)',
          },
        },
        required: ['presentationId', 'tableObjectId', 'rowIndex', 'columnIndex', 'text'],
      },
    },
    {
      name: 'gslides_create_bullets',
      description: 'Convert paragraphs in a shape into a bulleted or numbered list',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          objectId: { type: 'string', description: 'Object ID of the shape containing text' },
          bulletPreset: {
            type: 'string',
            description: 'Bullet style preset',
            enum: [
              'BULLET_DISC_CIRCLE_SQUARE',
              'BULLET_DIAMONDX_ARROW3D_SQUARE',
              'BULLET_CHECKBOX',
              'BULLET_ARROW_DIAMOND_DISC',
              'BULLET_STAR_CIRCLE_SQUARE',
              'BULLET_ARROW3D_CIRCLE_SQUARE',
              'BULLET_LEFTTRIANGLE_DIAMOND_DISC',
              'BULLET_DIAMONDX_HOLLOWDIAMOND_SQUARE',
              'BULLET_DIAMOND_CIRCLE_SQUARE',
              'NUMBERED_DIGIT_ALPHA_ROMAN',
              'NUMBERED_DIGIT_ALPHA_ROMAN_PARENS',
              'NUMBERED_DIGIT_NESTED',
              'NUMBERED_UPPERALPHA_ALPHA_ROMAN',
              'NUMBERED_UPPERROMAN_UPPERALPHA_DIGIT',
              'NUMBERED_ZERODIGIT_ALPHA_ROMAN',
            ],
          },
          startIndex: {
            type: 'number',
            description: 'Start character index of the text range to bullet (optional, defaults to all text)',
          },
          endIndex: {
            type: 'number',
            description: 'End character index of the text range to bullet (optional, defaults to all text)',
          },
        },
        required: ['presentationId', 'objectId'],
      },
    },
    // ── Phase 9: Page properties ──────────────────────────────────────────
    {
      name: 'gslides_update_page_properties',
      description: 'Update slide page properties (background color, etc.)',
      inputSchema: {
        type: 'object',
        properties: {
          presentationId: { type: 'string', description: 'The ID of the presentation' },
          slideObjectId: { type: 'string', description: 'Object ID of the slide' },
          backgroundColor: {
            type: 'object',
            description: 'Background color as { red, green, blue } values 0-1',
            properties: {
              red: { type: 'number' },
              green: { type: 'number' },
              blue: { type: 'number' },
            },
          },
        },
        required: ['presentationId', 'slideObjectId', 'backgroundColor'],
      },
    },
  ];
}

// ── Style analysis helpers ────────────────────────────────────────────────────

interface SlideStyleSample {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  foregroundColor?: { red: number; green: number; blue: number };
  charCount: number;
}

function extractSlidesStyles(slides: slides_v1.Schema$Page[]): {
  dominantStyle: Record<string, any> | null;
} {
  const samples: SlideStyleSample[] = [];

  for (const slide of slides) {
    for (const el of slide.pageElements || []) {
      const textElements = el.shape?.text?.textElements || [];
      for (const te of textElements) {
        const run = te.textRun;
        if (!run?.content || !run.style) continue;

        const text = run.content.replace(/\n/g, '');
        if (text.length === 0) continue;

        const ts = run.style;
        const sample: SlideStyleSample = {
          fontFamily: ts.fontFamily || ts.weightedFontFamily?.fontFamily || undefined,
          fontSize: ts.fontSize?.magnitude || undefined,
          bold: ts.bold || undefined,
          charCount: text.length,
        };

        const fg = ts.foregroundColor?.opaqueColor?.rgbColor;
        if (fg) {
          sample.foregroundColor = {
            red: Math.round((fg.red || 0) * 1000) / 1000,
            green: Math.round((fg.green || 0) * 1000) / 1000,
            blue: Math.round((fg.blue || 0) * 1000) / 1000,
          };
        }

        samples.push(sample);
      }
    }
  }

  return { dominantStyle: pickDominantSlideStyle(samples) };
}

function pickDominantSlideStyle(samples: SlideStyleSample[]): Record<string, any> | null {
  if (samples.length === 0) return null;

  const fontMap = new Map<string, number>();
  const sizeMap = new Map<number, number>();
  const colorMap = new Map<string, { color: { red: number; green: number; blue: number }; count: number }>();
  let boldChars = 0;
  let totalChars = 0;

  for (const s of samples) {
    totalChars += s.charCount;
    if (s.fontFamily) fontMap.set(s.fontFamily, (fontMap.get(s.fontFamily) || 0) + s.charCount);
    if (s.fontSize) sizeMap.set(s.fontSize, (sizeMap.get(s.fontSize) || 0) + s.charCount);
    if (s.bold) boldChars += s.charCount;
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

  let maxFont = 0;
  for (const [font, count] of fontMap) {
    if (count > maxFont) { maxFont = count; result.fontFamily = font; }
  }

  let maxSize = 0;
  for (const [size, count] of sizeMap) {
    if (count > maxSize) { maxSize = count; result.fontSize = size; }
  }

  if (boldChars > totalChars * 0.5) result.bold = true;

  let maxColor = 0;
  for (const [, entry] of colorMap) {
    if (entry.count > maxColor) { maxColor = entry.count; result.foregroundColor = entry.color; }
  }

  return Object.keys(result).length > 0 ? result : null;
}

// ── Handler class ──────────────────────────────────────────────────────────────

/** Convert points to EMU (English Metric Units). 1 pt = 12700 EMU. */
function ptToEmu(pt: number): number {
  return Math.round(pt * 12700);
}

export class SlidesHandler {
  constructor(
    private slides: slides_v1.Slides,
    private drive: drive_v3.Drive,
  ) {}

  async handleTool(name: string, args: any): Promise<ToolResponse | null> {
    switch (name) {
      // Phase 4
      case 'gslides_create_presentation':
        return this.handleCreatePresentation(args);
      case 'gslides_get_presentation':
        return this.handleGetPresentation(args);
      // Phase 5
      case 'gslides_add_slide':
        return this.handleAddSlide(args);
      case 'gslides_insert_text':
        return this.handleInsertText(args);
      case 'gslides_replace_text':
        return this.handleReplaceText(args);
      case 'gslides_speaker_notes':
        return this.handleSpeakerNotes(args);
      // Phase 6
      case 'gslides_add_shape':
        return this.handleAddShape(args);
      case 'gslides_add_image':
        return this.handleAddImage(args);
      case 'gslides_add_table':
        return this.handleAddTable(args);
      case 'gslides_format_text':
        return this.handleFormatText(args);
      // Phase 7
      case 'gslides_get_thumbnail':
        return this.handleGetThumbnail(args);
      case 'gslides_delete_slide':
        return this.handleDeleteSlide(args);
      // Phase 8
      case 'gslides_insert_table_text':
        return this.handleInsertTableText(args);
      case 'gslides_create_bullets':
        return this.handleCreateBullets(args);
      // Phase 9
      case 'gslides_update_page_properties':
        return this.handleUpdatePageProperties(args);
      // Phase 10
      case 'gslides_batch_format_text':
        return this.handleBatchFormatText(args);
      case 'gslides_format_table':
        return this.handleFormatTable(args);
      default:
        return null;
    }
  }

  // ── Shared validation helpers ──────────────────────────────────────────────

  private requireString(args: any, field: string): string {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    const value = args[field];
    if (!value || typeof value !== 'string') {
      throw new Error(`Invalid ${field}: expected non-empty string`);
    }
    return value;
  }

  private requireNumber(args: any, field: string): number {
    const value = args[field];
    if (typeof value !== 'number') {
      throw new Error(`Invalid ${field}: expected number`);
    }
    return value;
  }

  // ── Shared: extract text from page elements ────────────────────────────────

  private extractTextFromElements(elements: slides_v1.Schema$PageElement[] | undefined): string {
    if (!elements) return '';
    const parts: string[] = [];
    for (const el of elements) {
      if (el.shape?.text?.textElements) {
        for (const te of el.shape.text.textElements) {
          if (te.textRun?.content) {
            parts.push(te.textRun.content);
          }
        }
      }
    }
    return parts.join('');
  }

  private extractNotesText(slide: slides_v1.Schema$Page): string {
    const notesPage = slide.slideProperties?.notesPage;
    if (!notesPage?.pageElements) return '';
    for (const el of notesPage.pageElements) {
      if (el.shape?.placeholder?.type === 'BODY') {
        const parts: string[] = [];
        for (const te of el.shape.text?.textElements || []) {
          if (te.textRun?.content) parts.push(te.textRun.content);
        }
        return parts.join('');
      }
    }
    return '';
  }

  private findNotesBodyObjectId(slide: slides_v1.Schema$Page): string | null {
    const notesPage = slide.slideProperties?.notesPage;
    if (!notesPage?.pageElements) return null;
    for (const el of notesPage.pageElements) {
      if (el.shape?.placeholder?.type === 'BODY' && el.objectId) {
        return el.objectId;
      }
    }
    return null;
  }

  private getNotesTextLength(slide: slides_v1.Schema$Page): number {
    const notesPage = slide.slideProperties?.notesPage;
    if (!notesPage?.pageElements) return 0;
    for (const el of notesPage.pageElements) {
      if (el.shape?.placeholder?.type === 'BODY') {
        let len = 0;
        for (const te of el.shape.text?.textElements || []) {
          if (te.textRun?.content) len += te.textRun.content.length;
        }
        return len;
      }
    }
    return 0;
  }

  // ── Phase 4 handlers ───────────────────────────────────────────────────────

  private async handleCreatePresentation(args: any): Promise<ToolResponse> {
    try {
      const title = this.requireString(args, 'title');
      const parentFolderId = args.parentFolderId as string | undefined;

      const response = await this.slides.presentations.create({
        requestBody: { title },
      });

      const presentationId = response.data.presentationId!;

      if (parentFolderId) {
        const file = await this.drive.files.get({ fileId: presentationId, fields: 'parents' });
        const previousParents = file.data.parents?.join(',');
        await this.drive.files.update({
          fileId: presentationId,
          addParents: parentFolderId,
          removeParents: previousParents,
          fields: 'id, parents',
        });
      }

      return successResponse({
        success: true,
        presentationId,
        url: `https://docs.google.com/presentation/d/${presentationId}/edit`,
        title,
        slideCount: response.data.slides?.length || 0,
        parentFolderId: parentFolderId || 'root',
      });
    } catch (error) {
      return errorResponse('creating presentation', error);
    }
  }

  private async handleGetPresentation(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const includeText = args.includeText !== false;

      const response = await this.slides.presentations.get({ presentationId });
      const pres = response.data;
      const slides = pres.slides || [];

      const slidesSummary = slides.map((slide, index) => {
        const info: Record<string, any> = {
          index,
          objectId: slide.objectId,
          layout: slide.slideProperties?.layoutObjectId,
        };

        if (includeText) {
          // List all page elements with their IDs and types
          info.elements = (slide.pageElements || []).map((el) => {
            const elInfo: Record<string, any> = {
              objectId: el.objectId,
            };
            if (el.shape) {
              elInfo.type = 'shape';
              elInfo.shapeType = el.shape.shapeType;
              if (el.shape.placeholder) {
                elInfo.placeholder = el.shape.placeholder.type;
              }
              const text = this.extractTextFromElements([el]);
              if (text.trim()) elInfo.text = text.trim();
            } else if (el.table) {
              elInfo.type = 'table';
              elInfo.rows = el.table.rows;
              elInfo.columns = el.table.columns;
              // Extract text from each cell
              const cellData: string[][] = [];
              if (el.table.tableRows) {
                for (const row of el.table.tableRows) {
                  const rowTexts: string[] = [];
                  if (row.tableCells) {
                    for (const cell of row.tableCells) {
                      let cellText = '';
                      if (cell.text?.textElements) {
                        for (const te of cell.text.textElements) {
                          if (te.textRun?.content) {
                            cellText += te.textRun.content;
                          }
                        }
                      }
                      rowTexts.push(cellText.replace(/\n$/, ''));
                    }
                  }
                  cellData.push(rowTexts);
                }
              }
              elInfo.cellData = cellData;
            } else if (el.image) {
              elInfo.type = 'image';
              elInfo.sourceUrl = el.image.sourceUrl;
            }
            return elInfo;
          });

          const notes = this.extractNotesText(slide);
          if (notes.trim()) info.speakerNotes = notes.trim();
        }

        return info;
      });

      const result: Record<string, any> = {
        success: true,
        presentationId: pres.presentationId,
        title: pres.title,
        slideCount: slides.length,
        pageSize: pres.pageSize
          ? {
              width: pres.pageSize.width,
              height: pres.pageSize.height,
            }
          : undefined,
        slides: slidesSummary,
      };

      // Detect dominant text style across all slides
      if (includeText) {
        const { dominantStyle } = extractSlidesStyles(slides);
        if (dominantStyle) {
          result.detectedStyle = dominantStyle;
        }
      }

      return successResponse(result);
    } catch (error) {
      return errorResponse('getting presentation', error);
    }
  }

  // ── Phase 5 handlers ───────────────────────────────────────────────────────

  private async handleAddSlide(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const layout = args.layout || 'BLANK';
      const insertionIndex = args.insertionIndex as number | undefined;

      const createSlideRequest: slides_v1.Schema$CreateSlideRequest = {
        slideLayoutReference: { predefinedLayout: layout },
      };
      if (insertionIndex !== undefined) {
        createSlideRequest.insertionIndex = insertionIndex;
      }

      const response = await this.slides.presentations.batchUpdate({
        presentationId,
        requestBody: {
          requests: [{ createSlide: createSlideRequest }],
        },
      });

      const slideId = response.data.replies?.[0]?.createSlide?.objectId;

      return successResponse({
        success: true,
        slideObjectId: slideId,
        layout,
        insertionIndex,
      });
    } catch (error) {
      return errorResponse('adding slide', error);
    }
  }

  private async handleInsertText(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const objectId = this.requireString(args, 'objectId');
      const text = normaliseText(this.requireString(args, 'text'));
      const insertionIndex = (args.insertionIndex as number) || 0;

      await this.slides.presentations.batchUpdate({
        presentationId,
        requestBody: {
          requests: [
            {
              insertText: { objectId, text, insertionIndex },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        objectId,
        textLength: text.length,
        insertionIndex,
      });
    } catch (error) {
      return errorResponse('inserting text', error);
    }
  }

  private async handleReplaceText(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const find = normaliseText(this.requireString(args, 'find'));
      const replace = typeof args.replace === 'string' ? normaliseText(args.replace) : '';
      const matchCase = args.matchCase !== false;

      const replaceRequest: slides_v1.Schema$ReplaceAllTextRequest = {
        containsText: { text: find, matchCase },
        replaceText: replace,
      };

      if (Array.isArray(args.pageObjectIds) && args.pageObjectIds.length > 0) {
        replaceRequest.pageObjectIds = args.pageObjectIds;
      }

      const response = await this.slides.presentations.batchUpdate({
        presentationId,
        requestBody: {
          requests: [{ replaceAllText: replaceRequest }],
        },
      });

      const occurrences = response.data.replies?.[0]?.replaceAllText?.occurrencesChanged || 0;

      return successResponse({
        success: true,
        occurrencesChanged: occurrences,
        find,
        replace,
      });
    } catch (error) {
      return errorResponse('replacing text', error);
    }
  }

  private async handleSpeakerNotes(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const slideObjectId = this.requireString(args, 'slideObjectId');
      const notesText = typeof args.notes === 'string' ? normaliseText(args.notes) : undefined;

      // Get the presentation to find the notes body shape
      const pres = await this.slides.presentations.get({ presentationId });
      const slide = pres.data.slides?.find((s) => s.objectId === slideObjectId);
      if (!slide) throw new Error(`Slide not found: ${slideObjectId}`);

      const notesBodyId = this.findNotesBodyObjectId(slide);
      if (!notesBodyId) throw new Error('Speaker notes placeholder not found on this slide');

      // Read mode
      if (notesText === undefined) {
        const existingNotes = this.extractNotesText(slide);
        return successResponse({
          success: true,
          slideObjectId,
          notes: existingNotes,
        });
      }

      // Write mode: clear existing text, then insert new
      const requests: slides_v1.Schema$Request[] = [];

      const existingLength = this.getNotesTextLength(slide);
      if (existingLength > 0) {
        requests.push({
          deleteText: {
            objectId: notesBodyId,
            textRange: { type: 'ALL' },
          },
        });
      }

      if (notesText) {
        requests.push({
          insertText: {
            objectId: notesBodyId,
            text: notesText,
            insertionIndex: 0,
          },
        });
      }

      if (requests.length > 0) {
        await this.slides.presentations.batchUpdate({
          presentationId,
          requestBody: { requests },
        });
      }

      return successResponse({
        success: true,
        slideObjectId,
        notes: notesText,
        action: 'updated',
      });
    } catch (error) {
      return errorResponse('managing speaker notes', error);
    }
  }

  // ── Phase 6 handlers ───────────────────────────────────────────────────────

  private async handleAddShape(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const pageObjectId = this.requireString(args, 'pageObjectId');
      const shapeType = this.requireString(args, 'shapeType');
      const x = this.requireNumber(args, 'x');
      const y = this.requireNumber(args, 'y');
      const width = this.requireNumber(args, 'width');
      const height = this.requireNumber(args, 'height');

      const response = await this.slides.presentations.batchUpdate({
        presentationId,
        requestBody: {
          requests: [
            {
              createShape: {
                shapeType,
                elementProperties: {
                  pageObjectId,
                  size: {
                    width: { magnitude: ptToEmu(width), unit: 'EMU' },
                    height: { magnitude: ptToEmu(height), unit: 'EMU' },
                  },
                  transform: {
                    scaleX: 1,
                    scaleY: 1,
                    translateX: ptToEmu(x),
                    translateY: ptToEmu(y),
                    unit: 'EMU',
                  },
                },
              },
            },
          ],
        },
      });

      const objectId = response.data.replies?.[0]?.createShape?.objectId;

      return successResponse({
        success: true,
        objectId,
        shapeType,
        position: { x, y },
        size: { width, height },
      });
    } catch (error) {
      return errorResponse('adding shape', error);
    }
  }

  private async handleAddImage(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const pageObjectId = this.requireString(args, 'pageObjectId');
      const imageUrl = this.requireString(args, 'imageUrl');
      const x = this.requireNumber(args, 'x');
      const y = this.requireNumber(args, 'y');
      const width = this.requireNumber(args, 'width');
      const height = this.requireNumber(args, 'height');

      const response = await this.slides.presentations.batchUpdate({
        presentationId,
        requestBody: {
          requests: [
            {
              createImage: {
                url: imageUrl,
                elementProperties: {
                  pageObjectId,
                  size: {
                    width: { magnitude: ptToEmu(width), unit: 'EMU' },
                    height: { magnitude: ptToEmu(height), unit: 'EMU' },
                  },
                  transform: {
                    scaleX: 1,
                    scaleY: 1,
                    translateX: ptToEmu(x),
                    translateY: ptToEmu(y),
                    unit: 'EMU',
                  },
                },
              },
            },
          ],
        },
      });

      const objectId = response.data.replies?.[0]?.createImage?.objectId;

      return successResponse({
        success: true,
        objectId,
        imageUrl,
        position: { x, y },
        size: { width, height },
      });
    } catch (error) {
      return errorResponse('adding image', error);
    }
  }

  private async handleAddTable(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const pageObjectId = this.requireString(args, 'pageObjectId');
      const rows = this.requireNumber(args, 'rows');
      const columns = this.requireNumber(args, 'columns');
      const x = (args.x as number) ?? 100;
      const y = (args.y as number) ?? 100;
      const width = (args.width as number) ?? 600;
      const height = (args.height as number) ?? 300;

      const response = await this.slides.presentations.batchUpdate({
        presentationId,
        requestBody: {
          requests: [
            {
              createTable: {
                rows,
                columns,
                elementProperties: {
                  pageObjectId,
                  size: {
                    width: { magnitude: ptToEmu(width), unit: 'EMU' },
                    height: { magnitude: ptToEmu(height), unit: 'EMU' },
                  },
                  transform: {
                    scaleX: 1,
                    scaleY: 1,
                    translateX: ptToEmu(x),
                    translateY: ptToEmu(y),
                    unit: 'EMU',
                  },
                },
              },
            },
          ],
        },
      });

      const objectId = response.data.replies?.[0]?.createTable?.objectId;

      return successResponse({
        success: true,
        tableObjectId: objectId,
        rows,
        columns,
        position: { x, y },
        size: { width, height },
      });
    } catch (error) {
      return errorResponse('adding table', error);
    }
  }

  private async handleFormatText(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      if (!args.format || typeof args.format !== 'object') {
        throw new Error('Invalid format: expected object');
      }

      const request = this.buildTextStyleRequest(args);

      await this.slides.presentations.batchUpdate({
        presentationId,
        requestBody: { requests: [request] },
      });

      const textRange = (request.updateTextStyle as any).textRange;
      return successResponse({
        success: true,
        objectId: args.objectId,
        formattedRange: textRange.type === 'ALL' ? 'ALL' : `${args.startIndex}-${args.endIndex}`,
        appliedFormat: args.format,
      });
    } catch (error) {
      return errorResponse('formatting text', error);
    }
  }

  // ── Phase 7 handlers ───────────────────────────────────────────────────────

  private async handleGetThumbnail(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const slideObjectId = this.requireString(args, 'slideObjectId');
      const thumbnailSize = args.thumbnailSize || 'MEDIUM';

      const response = await this.slides.presentations.pages.getThumbnail({
        presentationId,
        pageObjectId: slideObjectId,
        'thumbnailProperties.thumbnailSize': thumbnailSize,
      });

      return successResponse({
        success: true,
        slideObjectId,
        contentUrl: response.data.contentUrl,
        width: response.data.width,
        height: response.data.height,
      });
    } catch (error) {
      return errorResponse('getting thumbnail', error);
    }
  }

  private async handleDeleteSlide(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const slideObjectId = this.requireString(args, 'slideObjectId');

      await this.slides.presentations.batchUpdate({
        presentationId,
        requestBody: {
          requests: [{ deleteObject: { objectId: slideObjectId } }],
        },
      });

      return successResponse({
        success: true,
        deletedSlideId: slideObjectId,
      });
    } catch (error) {
      return errorResponse('deleting slide', error);
    }
  }

  // ── Phase 8 handlers: Tables & bullets ──────────────────────────────────

  private async handleInsertTableText(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const tableObjectId = this.requireString(args, 'tableObjectId');
      const rowIndex = this.requireNumber(args, 'rowIndex');
      const columnIndex = this.requireNumber(args, 'columnIndex');
      const text = normaliseText(this.requireString(args, 'text'));
      const insertionIndex = (args.insertionIndex as number) || 0;

      await this.slides.presentations.batchUpdate({
        presentationId,
        requestBody: {
          requests: [
            {
              insertText: {
                objectId: tableObjectId,
                cellLocation: { rowIndex, columnIndex },
                text,
                insertionIndex,
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        tableObjectId,
        rowIndex,
        columnIndex,
        textLength: text.length,
      });
    } catch (error) {
      return errorResponse('inserting table text', error);
    }
  }

  private async handleCreateBullets(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const objectId = this.requireString(args, 'objectId');
      const bulletPreset = args.bulletPreset || 'BULLET_DISC_CIRCLE_SQUARE';

      const textRange: slides_v1.Schema$Range =
        typeof args.startIndex === 'number' && typeof args.endIndex === 'number'
          ? { type: 'FIXED_RANGE', startIndex: args.startIndex, endIndex: args.endIndex }
          : { type: 'ALL' };

      await this.slides.presentations.batchUpdate({
        presentationId,
        requestBody: {
          requests: [
            {
              createParagraphBullets: {
                objectId,
                textRange,
                bulletPreset,
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        objectId,
        bulletPreset,
        textRange: textRange.type === 'ALL' ? 'ALL' : `${args.startIndex}-${args.endIndex}`,
      });
    } catch (error) {
      return errorResponse('creating bullets', error);
    }
  }

  // ── Phase 9 handlers: Page properties ─────────────────────────────────

  private async handleUpdatePageProperties(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const slideObjectId = this.requireString(args, 'slideObjectId');
      const backgroundColor = args.backgroundColor;

      if (!backgroundColor || typeof backgroundColor !== 'object') {
        throw new Error('Invalid backgroundColor: expected { red, green, blue } with values 0-1');
      }

      await this.slides.presentations.batchUpdate({
        presentationId,
        requestBody: {
          requests: [
            {
              updatePageProperties: {
                objectId: slideObjectId,
                pageProperties: {
                  pageBackgroundFill: {
                    solidFill: {
                      color: {
                        rgbColor: {
                          red: backgroundColor.red,
                          green: backgroundColor.green,
                          blue: backgroundColor.blue,
                        },
                      },
                    },
                  },
                },
                fields: 'pageBackgroundFill.solidFill.color',
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        slideObjectId,
        backgroundColor,
      });
    } catch (error) {
      return errorResponse('updating page properties', error);
    }
  }

  // ── Phase 10 handlers: Batch operations ──────────────────────────────

  private buildTextStyleRequest(op: any): slides_v1.Schema$Request {
    const format = op.format;
    const style: slides_v1.Schema$TextStyle = {};
    const fieldParts: string[] = [];

    if (format.bold !== undefined) {
      style.bold = format.bold;
      fieldParts.push('bold');
    }
    if (format.italic !== undefined) {
      style.italic = format.italic;
      fieldParts.push('italic');
    }
    if (format.underline !== undefined) {
      style.underline = format.underline;
      fieldParts.push('underline');
    }
    if (format.fontSize !== undefined) {
      style.fontSize = { magnitude: format.fontSize, unit: 'PT' };
      fieldParts.push('fontSize');
    }
    if (format.fontFamily !== undefined) {
      style.fontFamily = format.fontFamily;
      fieldParts.push('fontFamily');
    }
    if (format.foregroundColor) {
      style.foregroundColor = {
        opaqueColor: { rgbColor: format.foregroundColor },
      };
      fieldParts.push('foregroundColor');
    }

    if (fieldParts.length === 0) {
      throw new Error('No formatting properties provided in operation');
    }

    const textRange: any =
      typeof op.startIndex === 'number' && typeof op.endIndex === 'number'
        ? { type: 'FIXED_RANGE', startIndex: op.startIndex, endIndex: op.endIndex }
        : { type: 'ALL' };

    const updateTextStyle: any = {
      objectId: op.objectId,
      textRange,
      style,
      fields: fieldParts.join(','),
    };

    if (typeof op.rowIndex === 'number' && typeof op.columnIndex === 'number') {
      updateTextStyle.cellLocation = {
        rowIndex: op.rowIndex,
        columnIndex: op.columnIndex,
      };
    }

    return { updateTextStyle };
  }

  private async handleFormatTable(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const objectId = this.requireString(args, 'tableObjectId');
      const requests: slides_v1.Schema$Request[] = [];

      // Column widths
      if (Array.isArray(args.columnWidths)) {
        for (let i = 0; i < args.columnWidths.length; i++) {
          const widthPt = args.columnWidths[i];
          if (typeof widthPt === 'number' && widthPt > 0) {
            requests.push({
              updateTableColumnProperties: {
                objectId,
                columnIndices: [i],
                tableColumnProperties: {
                  columnWidth: { magnitude: ptToEmu(widthPt), unit: 'EMU' },
                },
                fields: 'columnWidth',
              },
            });
          }
        }
      }

      // Row heights
      if (Array.isArray(args.rowHeights)) {
        for (let i = 0; i < args.rowHeights.length; i++) {
          const heightPt = args.rowHeights[i];
          if (typeof heightPt === 'number' && heightPt > 0) {
            requests.push({
              updateTableRowProperties: {
                objectId,
                rowIndices: [i],
                tableRowProperties: {
                  minRowHeight: { magnitude: ptToEmu(heightPt), unit: 'EMU' },
                },
                fields: 'minRowHeight',
              },
            });
          }
        }
      }

      // Content alignment (applies to all cells)
      if (args.contentAlignment) {
        requests.push({
          updateTableCellProperties: {
            objectId,
            tableCellProperties: {
              contentAlignment: args.contentAlignment,
            },
            fields: 'contentAlignment',
          },
        });
      }

      // Header row background color
      if (args.headerRowBackgroundColor && args.totalColumns) {
        requests.push({
          updateTableCellProperties: {
            objectId,
            tableRange: {
              location: { rowIndex: 0, columnIndex: 0 },
              rowSpan: 1,
              columnSpan: args.totalColumns,
            },
            tableCellProperties: {
              tableCellBackgroundFill: {
                solidFill: {
                  color: { rgbColor: args.headerRowBackgroundColor },
                },
              },
            },
            fields: 'tableCellBackgroundFill.solidFill.color',
          },
        });
      }

      // Body rows background color
      if (args.bodyRowBackgroundColor && args.totalRows && args.totalColumns && args.totalRows > 1) {
        requests.push({
          updateTableCellProperties: {
            objectId,
            tableRange: {
              location: { rowIndex: 1, columnIndex: 0 },
              rowSpan: args.totalRows - 1,
              columnSpan: args.totalColumns,
            },
            tableCellProperties: {
              tableCellBackgroundFill: {
                solidFill: {
                  color: { rgbColor: args.bodyRowBackgroundColor },
                },
              },
            },
            fields: 'tableCellBackgroundFill.solidFill.color',
          },
        });
      }

      if (requests.length === 0) {
        return successResponse({ success: true, message: 'No properties to update' });
      }

      await this.slides.presentations.batchUpdate({
        presentationId,
        requestBody: { requests },
      });

      return successResponse({
        success: true,
        tableObjectId: objectId,
        requestCount: requests.length,
      });
    } catch (error) {
      return errorResponse('formatting table', error);
    }
  }

  private async handleBatchFormatText(args: any): Promise<ToolResponse> {
    try {
      const presentationId = this.requireString(args, 'presentationId');
      const operations = args.operations;
      if (!Array.isArray(operations) || operations.length === 0) {
        throw new Error('Invalid operations: expected non-empty array');
      }

      const requests: slides_v1.Schema$Request[] = operations.map((op: any) =>
        this.buildTextStyleRequest(op),
      );

      await this.slides.presentations.batchUpdate({
        presentationId,
        requestBody: { requests },
      });

      return successResponse({
        success: true,
        operationCount: requests.length,
      });
    } catch (error) {
      return errorResponse('batch formatting text', error);
    }
  }
}
