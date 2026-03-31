import { docs_v1, drive_v3 } from 'googleapis';
import type {
  CreateDocumentRequest,
  InsertTextRequest,
  ReplaceTextRequest,
  FormatTextRequest,
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
      description: 'Get the content of a Google Document',
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
      description: 'Apply formatting to text in a Google Document',
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
              fontSize: { type: 'number' },
            },
          },
        },
        required: ['documentId', 'startIndex', 'endIndex', 'format'],
      },
    },
    {
      name: 'gdocs_set_heading',
      description: 'Convert text to a heading in a Google Document',
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
      case 'gdocs_set_heading': {
        const formatArgs = this.validateFormatTextArgs(args);
        const headingLevel = (args as any)?.headingLevel;
        if (!headingLevel || headingLevel < 1 || headingLevel > 6) {
          throw new Error('Invalid headingLevel: expected number between 1 and 6');
        }
        return this.handleSetHeading({ ...formatArgs, headingLevel });
      }
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

  private validateFormatTextArgs(args: any): FormatTextRequest {
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
    if (!args.format || typeof args.format !== 'object') {
      throw new Error('Invalid format: expected object');
    }
    return {
      documentId: args.documentId,
      startIndex: args.startIndex,
      endIndex: args.endIndex,
      format: args.format,
    };
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
        const file = await this.drive.files.get({ fileId: documentId, fields: 'parents' });
        const previousParents = file.data.parents?.join(',');
        await this.drive.files.update({
          fileId: documentId,
          addParents: args.parentFolderId,
          removeParents: previousParents,
          fields: 'id, parents',
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
            for (const [id, change] of Object.entries(run.suggestedTextStyleChanges)) {
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
      const processTable = (table: docs_v1.Schema$Table): string => {
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
            text += processTable(element.table);
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
        for (const [id, group] of grouped) {
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

      return textResponse(`Document: ${doc.title}\n\nContent:\n${content}${suggestionsInfo}${styleInfo}`);
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
      const textStyle: any = {};
      if (args.format.bold !== undefined) textStyle.bold = args.format.bold;
      if (args.format.italic !== undefined) textStyle.italic = args.format.italic;
      if (args.format.underline !== undefined) textStyle.underline = args.format.underline;
      if (args.format.fontSize !== undefined) {
        textStyle.fontSize = { magnitude: args.format.fontSize, unit: 'PT' };
      }

      await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              updateTextStyle: {
                range: { startIndex: args.startIndex, endIndex: args.endIndex },
                textStyle,
                fields: Object.keys(textStyle).join(','),
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

  private async handleSetHeading(
    args: FormatTextRequest & { headingLevel: number },
  ): Promise<ToolResponse> {
    try {
      const namedStyleType = `HEADING_${args.headingLevel}` as any;

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
}
