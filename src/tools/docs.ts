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
      const response = await this.docs.documents.get({ documentId: args.documentId });
      const doc = response.data;

      let content = '';
      if (doc.body?.content) {
        for (const element of doc.body.content) {
          if (element.paragraph?.elements) {
            for (const elem of element.paragraph.elements) {
              if (elem.textRun?.content) {
                content += elem.textRun.content;
              }
            }
          }
        }
      }

      return textResponse(`Document: ${doc.title}\n\nContent:\n${content}`);
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
