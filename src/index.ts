import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { google, sheets_v4, drive_v3, docs_v1 } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';

interface BatchUpdateRequest {
  spreadsheetId: string;
  updates: {
    range: string;
    values: any[][];
  }[];
}

interface CreateSheetRequest {
  title: string;
  sheetTitle?: string;
  parentFolderId?: string;
}

interface CreateDocumentRequest {
  title: string;
  content?: string;
  parentFolderId?: string;
}

interface InsertTextRequest {
  documentId: string;
  text: string;
  index?: number;
}

interface ReplaceTextRequest {
  documentId: string;
  find: string;
  replace: string;
}

interface FormatTextRequest {
  documentId: string;
  startIndex: number;
  endIndex: number;
  format: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontSize?: number;
  };
}

interface MoveFileRequest {
  fileId: string;
  targetFolderId: string;
}

interface BatchMoveRequest {
  fileIds: string[];
  targetFolderId: string;
}

interface CreateFolderRequest {
  name: string;
  parentFolderId?: string;
}

interface CopyFolderRequest {
  sourceFolderId: string;
  targetParentFolderId?: string;
  newName?: string;
}

interface SearchContentRequest {
  query: string;
  folderId?: string;
  maxResults?: number;
}

interface GetRevisionsRequest {
  fileId: string;
  maxResults?: number;
}

interface ExportFileRequest {
  fileId: string;
  mimeType: string;
}

interface BatchExportRequest {
  fileIds: string[];
  format: 'pdf' | 'docx' | 'xlsx' | 'pptx';
}

interface ListFolderTreeRequest {
  folderId: string;
  recursive?: boolean;
  includeMetadata?: boolean;
}

interface ListPermissionsRequest {
  fileId: string;
}

class GoogleWorkspaceMCP {
  private server: Server;
  private auth: OAuth2Client;
  private sheets: sheets_v4.Sheets;
  private drive: drive_v3.Drive;
  private docs: docs_v1.Docs;

  constructor() {
    // Add debugging
    console.error('Starting GoogleWorkspaceMCP...');
    console.error('Environment check:');
    console.error('GOOGLE_CLIENT_ID:', process.env.GOOGLE_CLIENT_ID ? 'Set' : 'Missing');
    console.error('GOOGLE_CLIENT_SECRET:', process.env.GOOGLE_CLIENT_SECRET ? 'Set' : 'Missing');
    console.error('GOOGLE_REFRESH_TOKEN:', process.env.GOOGLE_REFRESH_TOKEN ? 'Set' : 'Missing');

    this.server = new Server(
      {
        name: 'google-workspace',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    try {
      // Initialize OAuth2 Client with error handling
      const clientId = process.env.GOOGLE_CLIENT_ID;
      const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
      
      if (!clientId || !clientSecret) {
        console.error('Warning: Google OAuth credentials not set. Some features may not work.');
        // Create a minimal auth client that will fail gracefully
        this.auth = new OAuth2Client();
      } else {
        this.auth = new OAuth2Client({
          clientId,
          clientSecret,
          redirectUri: 'http://localhost:3000/oauth/callback', // Modern OAuth flow
        });

        // Set the refresh token if available
        if (process.env.GOOGLE_REFRESH_TOKEN) {
          this.auth.setCredentials({
            refresh_token: process.env.GOOGLE_REFRESH_TOKEN,
          });
        }
      }

      this.sheets = google.sheets({ version: 'v4', auth: this.auth });
      this.drive = google.drive({ version: 'v3', auth: this.auth });
      this.docs = google.docs({ version: 'v1', auth: this.auth });
      console.error('Google Workspace clients initialized successfully');

      this.setupToolHandlers();
      console.error('Tool handlers set up successfully');
      
    } catch (error) {
      console.error('Error during initialization:', error);
      throw error;
    }
  }

  // Add a method to get the authorization URL for initial setup
  private getAuthUrl(): string {
    const scopes = [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive.file',  // Changed from drive.readonly to drive.file for write access
      'https://www.googleapis.com/auth/documents'
    ];

    return this.auth.generateAuthUrl({
      access_type: 'offline',
      scope: scopes,
      prompt: 'consent', // Forces refresh token generation
    });
  }

  // Add a method to exchange authorization code for tokens
  private async getTokensFromCode(code: string) {
    const { tokens } = await this.auth.getToken(code);
    this.auth.setCredentials(tokens);
    return tokens;
  }

  private validateBatchUpdateArgs(args: any): BatchUpdateRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.spreadsheetId || typeof args.spreadsheetId !== 'string') {
      throw new Error('Invalid spreadsheetId: expected non-empty string');
    }
    if (!Array.isArray(args.updates)) {
      throw new Error('Invalid updates: expected array');
    }
    
    const updates = args.updates.map((update: any, index: number) => {
      if (!update || typeof update !== 'object') {
        throw new Error(`Invalid update at index ${index}: expected object`);
      }
      if (!update.range || typeof update.range !== 'string') {
        throw new Error(`Invalid range at index ${index}: expected non-empty string`);
      }
      if (!Array.isArray(update.values)) {
        throw new Error(`Invalid values at index ${index}: expected 2D array`);
      }
      return {
        range: update.range,
        values: update.values
      };
    });

    return {
      spreadsheetId: args.spreadsheetId,
      updates
    };
  }

  private validateCreateArgs(args: any) {
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
      sheetTitle: args.sheetTitle || 'Sheet1',
      data: args.data || [],
      parentFolderId: args.parentFolderId
    };
  }

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
      content: args.content || '',
      parentFolderId: args.parentFolderId
    };
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
      text: args.text,
      index: args.index || undefined
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
      find: args.find,
      replace: args.replace
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
      format: args.format
    };
  }

  private validateAppendArgs(args: any) {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.spreadsheetId || typeof args.spreadsheetId !== 'string') {
      throw new Error('Invalid spreadsheetId: expected non-empty string');
    }
    if (!args.range || typeof args.range !== 'string') {
      throw new Error('Invalid range: expected non-empty string');
    }
    if (!Array.isArray(args.values)) {
      throw new Error('Invalid values: expected 2D array');
    }
    return {
      spreadsheetId: args.spreadsheetId,
      range: args.range,
      values: args.values
    };
  }

  private validateFormatArgs(args: any) {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.spreadsheetId || typeof args.spreadsheetId !== 'string') {
      throw new Error('Invalid spreadsheetId: expected non-empty string');
    }
    if (!Array.isArray(args.requests)) {
      throw new Error('Invalid requests: expected array');
    }
    return {
      spreadsheetId: args.spreadsheetId,
      requests: args.requests
    };
  }

  private validateDriveSearchArgs(args: any) {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.query || typeof args.query !== 'string') {
      throw new Error('Invalid query: expected non-empty string');
    }
    return {
      query: args.query,
      pageSize: args.pageSize || 10,
      pageToken: args.pageToken || undefined
    };
  }

  private validateDriveReadArgs(args: any) {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.fileId || typeof args.fileId !== 'string') {
      throw new Error('Invalid fileId: expected non-empty string');
    }
    return {
      fileId: args.fileId
    };
  }

  private validateDocumentIdArgs(args: any) {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.documentId || typeof args.documentId !== 'string') {
      throw new Error('Invalid documentId: expected non-empty string');
    }
    return {
      documentId: args.documentId
    };
  }

  private validateMoveFileArgs(args: any): MoveFileRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.fileId || typeof args.fileId !== 'string') {
      throw new Error('Invalid fileId: expected non-empty string');
    }
    if (!args.targetFolderId || typeof args.targetFolderId !== 'string') {
      throw new Error('Invalid targetFolderId: expected non-empty string');
    }
    return {
      fileId: args.fileId,
      targetFolderId: args.targetFolderId
    };
  }

  private validateBatchMoveArgs(args: any): BatchMoveRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!Array.isArray(args.fileIds) || args.fileIds.length === 0) {
      throw new Error('Invalid fileIds: expected non-empty array');
    }
    if (!args.targetFolderId || typeof args.targetFolderId !== 'string') {
      throw new Error('Invalid targetFolderId: expected non-empty string');
    }
    return {
      fileIds: args.fileIds,
      targetFolderId: args.targetFolderId
    };
  }

  private validateCreateFolderArgs(args: any): CreateFolderRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.name || typeof args.name !== 'string') {
      throw new Error('Invalid name: expected non-empty string');
    }
    if (args.parentFolderId && typeof args.parentFolderId !== 'string') {
      throw new Error('Invalid parentFolderId: expected string');
    }
    return {
      name: args.name,
      parentFolderId: args.parentFolderId
    };
  }

  private validateCopyFolderArgs(args: any): CopyFolderRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.sourceFolderId || typeof args.sourceFolderId !== 'string') {
      throw new Error('Invalid sourceFolderId: expected non-empty string');
    }
    if (args.targetParentFolderId && typeof args.targetParentFolderId !== 'string') {
      throw new Error('Invalid targetParentFolderId: expected string');
    }
    if (args.newName && typeof args.newName !== 'string') {
      throw new Error('Invalid newName: expected string');
    }
    return {
      sourceFolderId: args.sourceFolderId,
      targetParentFolderId: args.targetParentFolderId,
      newName: args.newName
    };
  }

  private validateSearchContentArgs(args: any): SearchContentRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.query || typeof args.query !== 'string') {
      throw new Error('Invalid query: expected non-empty string');
    }
    if (args.folderId && typeof args.folderId !== 'string') {
      throw new Error('Invalid folderId: expected string');
    }
    if (args.maxResults && typeof args.maxResults !== 'number') {
      throw new Error('Invalid maxResults: expected number');
    }
    return {
      query: args.query,
      folderId: args.folderId,
      maxResults: args.maxResults || 100
    };
  }

  private validateGetRevisionsArgs(args: any): GetRevisionsRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.fileId || typeof args.fileId !== 'string') {
      throw new Error('Invalid fileId: expected non-empty string');
    }
    if (args.maxResults && typeof args.maxResults !== 'number') {
      throw new Error('Invalid maxResults: expected number');
    }
    return {
      fileId: args.fileId,
      maxResults: args.maxResults || 20
    };
  }

  private validateExportFileArgs(args: any): ExportFileRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.fileId || typeof args.fileId !== 'string') {
      throw new Error('Invalid fileId: expected non-empty string');
    }
    if (!args.mimeType || typeof args.mimeType !== 'string') {
      throw new Error('Invalid mimeType: expected non-empty string');
    }
    return {
      fileId: args.fileId,
      mimeType: args.mimeType
    };
  }

  private validateBatchExportArgs(args: any): BatchExportRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!Array.isArray(args.fileIds) || args.fileIds.length === 0) {
      throw new Error('Invalid fileIds: expected non-empty array');
    }
    if (!args.format || !['pdf', 'docx', 'xlsx', 'pptx'].includes(args.format)) {
      throw new Error('Invalid format: expected one of: pdf, docx, xlsx, pptx');
    }
    return {
      fileIds: args.fileIds,
      format: args.format
    };
  }

  private validateListFolderTreeArgs(args: any): ListFolderTreeRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.folderId || typeof args.folderId !== 'string') {
      throw new Error('Invalid folderId: expected non-empty string');
    }
    return {
      folderId: args.folderId,
      recursive: args.recursive !== false, // Default to true
      includeMetadata: args.includeMetadata === true
    };
  }

  private validateListPermissionsArgs(args: any): ListPermissionsRequest {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
    if (!args.fileId || typeof args.fileId !== 'string') {
      throw new Error('Invalid fileId: expected non-empty string');
    }
    return {
      fileId: args.fileId
    };
  }

  private setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        // Existing Google Sheets tools
        {
          name: 'gsheets_batch_update',
          description: 'Update multiple ranges in a Google Sheet in a single API call',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: {
                type: 'string',
                description: 'The ID of the spreadsheet to update',
              },
              updates: {
                type: 'array',
                description: 'Array of range updates to perform',
                items: {
                  type: 'object',
                  properties: {
                    range: {
                      type: 'string',
                      description: 'A1 notation range (e.g., "Sheet1!A1:C3")',
                    },
                    values: {
                      type: 'array',
                      description: '2D array of values to insert',
                      items: {
                        type: 'array',
                        items: { type: 'string' }
                      }
                    },
                  },
                  required: ['range', 'values'],
                },
              },
            },
            required: ['spreadsheetId', 'updates'],
          },
        },
        {
          name: 'gsheets_create_and_populate',
          description: 'Create a new Google Sheet and populate it with data',
          inputSchema: {
            type: 'object',
            properties: {
              title: {
                type: 'string',
                description: 'Title for the new spreadsheet',
              },
              sheetTitle: {
                type: 'string',
                description: 'Title for the first sheet tab (optional)',
              },
              data: {
                type: 'array',
                description: '2D array of data to populate the sheet',
                items: {
                  type: 'array',
                  items: { type: 'string' }
                }
              },
              parentFolderId: {
                type: 'string',
                description: 'ID of the parent folder to create the spreadsheet in (optional, defaults to root)',
              },
            },
            required: ['title', 'data'],
          },
        },
        {
          name: 'gsheets_append_rows',
          description: 'Append rows to the end of a sheet',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: {
                type: 'string',
                description: 'The ID of the spreadsheet',
              },
              range: {
                type: 'string',
                description: 'Range to append to (e.g., "Sheet1!A:Z")',
              },
              values: {
                type: 'array',
                description: '2D array of values to append',
                items: {
                  type: 'array',
                  items: { type: 'string' }
                }
              },
            },
            required: ['spreadsheetId', 'range', 'values'],
          },
        },
        {
          name: 'gsheets_format_cells',
          description: 'Apply formatting to cell ranges',
          inputSchema: {
            type: 'object',
            properties: {
              spreadsheetId: {
                type: 'string',
                description: 'The ID of the spreadsheet',
              },
              requests: {
                type: 'array',
                description: 'Array of formatting requests',
                items: {
                  type: 'object',
                  properties: {
                    range: {
                      type: 'string',
                      description: 'A1 notation range to format',
                    },
                    format: {
                      type: 'object',
                      description: 'Formatting options (bold, backgroundColor, etc.)',
                    },
                  },
                },
              },
            },
            required: ['spreadsheetId', 'requests'],
          },
        },
        // New Google Docs tools
        {
          name: 'gdocs_create_document',
          description: 'Create a new Google Document',
          inputSchema: {
            type: 'object',
            properties: {
              title: {
                type: 'string',
                description: 'Title for the new document',
              },
              content: {
                type: 'string',
                description: 'Initial content for the document (optional)',
              },
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
              documentId: {
                type: 'string',
                description: 'The ID of the document to retrieve',
              },
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
              documentId: {
                type: 'string',
                description: 'The ID of the document',
              },
              text: {
                type: 'string',
                description: 'Text to insert',
              },
              index: {
                type: 'number',
                description: 'Position to insert text (optional, defaults to end)',
              },
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
              documentId: {
                type: 'string',
                description: 'The ID of the document',
              },
              text: {
                type: 'string',
                description: 'Text to append',
              },
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
              documentId: {
                type: 'string',
                description: 'The ID of the document',
              },
              find: {
                type: 'string',
                description: 'Text to find',
              },
              replace: {
                type: 'string',
                description: 'Text to replace with',
              },
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
              documentId: {
                type: 'string',
                description: 'The ID of the document',
              },
              startIndex: {
                type: 'number',
                description: 'Start index of text to format',
              },
              endIndex: {
                type: 'number',
                description: 'End index of text to format',
              },
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
              documentId: {
                type: 'string',
                description: 'The ID of the document',
              },
              startIndex: {
                type: 'number',
                description: 'Start index of text to make a heading',
              },
              endIndex: {
                type: 'number',
                description: 'End index of text to make a heading',
              },
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
        // Existing auth and drive tools
        {
          name: 'gsheets_get_auth_url',
          description: 'Get Google OAuth authorization URL for initial setup',
          inputSchema: {
            type: 'object',
            properties: {},
          },
        },
        {
          name: 'gsheets_set_auth_code',
          description: 'Exchange authorization code for access tokens',
          inputSchema: {
            type: 'object',
            properties: {
              code: {
                type: 'string',
                description: 'Authorization code from Google OAuth flow',
              },
            },
            required: ['code'],
          },
        },
        {
          name: 'gdrive_search',
          description: 'Search for files in Google Drive',
          inputSchema: {
            type: 'object',
            properties: {
              query: {
                type: 'string',
                description: 'Search query (e.g., "name contains \'budget\'" or "type = \'spreadsheet\'")',
              },
              pageSize: {
                type: 'number',
                description: 'Number of results per page (max 100)',
                default: 10,
              },
              pageToken: {
                type: 'string',
                description: 'Token for the next page of results',
              },
            },
            required: ['query'],
          },
        },
        {
          name: 'gdrive_read_file',
          description: 'Read contents of a file from Google Drive',
          inputSchema: {
            type: 'object',
            properties: {
              fileId: {
                type: 'string',
                description: 'ID of the file to read',
              },
            },
            required: ['fileId'],
          },
        },
        {
          name: 'gdrive_get_file_info',
          description: 'Get metadata information about a Google Drive file',
          inputSchema: {
            type: 'object',
            properties: {
              fileId: {
                type: 'string',
                description: 'ID of the file to get info for',
              },
            },
            required: ['fileId'],
          },
        },
        {
          name: 'gdrive_move_file',
          description: 'Move a file to a different folder in Google Drive',
          inputSchema: {
            type: 'object',
            properties: {
              fileId: {
                type: 'string',
                description: 'ID of the file to move',
              },
              targetFolderId: {
                type: 'string',
                description: 'ID of the target folder to move the file to',
              },
            },
            required: ['fileId', 'targetFolderId'],
          },
        },
        {
          name: 'gdrive_batch_move',
          description: 'Move multiple files to a folder at once (bulk operation)',
          inputSchema: {
            type: 'object',
            properties: {
              fileIds: {
                type: 'array',
                description: 'Array of file IDs to move',
                items: { type: 'string' }
              },
              targetFolderId: {
                type: 'string',
                description: 'ID of the target folder',
              },
            },
            required: ['fileIds', 'targetFolderId'],
          },
        },
        {
          name: 'gdrive_create_folder',
          description: 'Create a new folder in Google Drive',
          inputSchema: {
            type: 'object',
            properties: {
              name: {
                type: 'string',
                description: 'Name of the folder to create',
              },
              parentFolderId: {
                type: 'string',
                description: 'ID of the parent folder (optional, defaults to root)',
              },
            },
            required: ['name'],
          },
        },
        {
          name: 'gdrive_copy_folder',
          description: 'Recursively copy a folder and all its contents',
          inputSchema: {
            type: 'object',
            properties: {
              sourceFolderId: {
                type: 'string',
                description: 'ID of the folder to copy',
              },
              targetParentFolderId: {
                type: 'string',
                description: 'ID of the parent folder for the copy (optional, defaults to root)',
              },
              newName: {
                type: 'string',
                description: 'Name for the copied folder (optional, defaults to "Copy of [original name]")',
              },
            },
            required: ['sourceFolderId'],
          },
        },
        {
          name: 'gdrive_search_content',
          description: 'Search for text within file contents (fullText search)',
          inputSchema: {
            type: 'object',
            properties: {
              query: {
                type: 'string',
                description: 'Text to search for within file contents',
              },
              folderId: {
                type: 'string',
                description: 'Limit search to specific folder (optional)',
              },
              maxResults: {
                type: 'number',
                description: 'Maximum number of results to return (default: 100)',
              },
            },
            required: ['query'],
          },
        },
        {
          name: 'gdrive_get_revisions',
          description: 'Get version history/revisions for a file',
          inputSchema: {
            type: 'object',
            properties: {
              fileId: {
                type: 'string',
                description: 'ID of the file',
              },
              maxResults: {
                type: 'number',
                description: 'Maximum number of revisions to return (default: 20)',
              },
            },
            required: ['fileId'],
          },
        },
        {
          name: 'gdrive_export_file',
          description: 'Export a Google Workspace file to a specific format (PDF, Word, Excel, etc.)',
          inputSchema: {
            type: 'object',
            properties: {
              fileId: {
                type: 'string',
                description: 'ID of the file to export',
              },
              mimeType: {
                type: 'string',
                description: 'Target MIME type (e.g., "application/pdf", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")',
              },
            },
            required: ['fileId', 'mimeType'],
          },
        },
        {
          name: 'gdrive_batch_export',
          description: 'Export multiple files to a specific format at once',
          inputSchema: {
            type: 'object',
            properties: {
              fileIds: {
                type: 'array',
                description: 'Array of file IDs to export',
                items: { type: 'string' }
              },
              format: {
                type: 'string',
                description: 'Export format: pdf, docx, xlsx, or pptx',
                enum: ['pdf', 'docx', 'xlsx', 'pptx']
              },
            },
            required: ['fileIds', 'format'],
          },
        },
        {
          name: 'gdrive_list_folder_tree',
          description: 'List all files in a folder, optionally recursive with metadata',
          inputSchema: {
            type: 'object',
            properties: {
              folderId: {
                type: 'string',
                description: 'ID of the folder to list',
              },
              recursive: {
                type: 'boolean',
                description: 'Recursively list subfolders (default: true)',
              },
              includeMetadata: {
                type: 'boolean',
                description: 'Include detailed metadata for each file (default: false)',
              },
            },
            required: ['folderId'],
          },
        },
        {
          name: 'gdrive_list_permissions',
          description: 'List all permissions (who has access) for a file or folder',
          inputSchema: {
            type: 'object',
            properties: {
              fileId: {
                type: 'string',
                description: 'ID of the file or folder',
              },
            },
            required: ['fileId'],
          },
        },
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        switch (request.params.name) {
          // Existing Google Sheets handlers
          case 'gsheets_batch_update':
            const batchArgs = this.validateBatchUpdateArgs(request.params.arguments);
            return this.handleBatchUpdate(batchArgs);
          
          case 'gsheets_create_and_populate':
            const createArgs = this.validateCreateArgs(request.params.arguments);
            return this.handleCreateAndPopulate(createArgs);
          
          case 'gsheets_append_rows':
            const appendArgs = this.validateAppendArgs(request.params.arguments);
            return this.handleAppendRows(appendArgs);
          
          case 'gsheets_format_cells':
            const formatArgs = this.validateFormatArgs(request.params.arguments);
            return this.handleFormatCells(formatArgs);
          
          // New Google Docs handlers
          case 'gdocs_create_document':
            const createDocArgs = this.validateCreateDocArgs(request.params.arguments);
            return this.handleCreateDocument(createDocArgs);
          
          case 'gdocs_get_document':
            const getDocArgs = this.validateDocumentIdArgs(request.params.arguments);
            return this.handleGetDocument(getDocArgs);
          
          case 'gdocs_insert_text':
            const insertArgs = this.validateInsertTextArgs(request.params.arguments);
            return this.handleInsertText(insertArgs);
          
          case 'gdocs_append_text':
            const appendTextArgs = this.validateInsertTextArgs(request.params.arguments);
            return this.handleAppendText(appendTextArgs);
          
          case 'gdocs_replace_text':
            const replaceArgs = this.validateReplaceTextArgs(request.params.arguments);
            return this.handleReplaceText(replaceArgs);
          
          case 'gdocs_format_text':
            const formatTextArgs = this.validateFormatTextArgs(request.params.arguments);
            return this.handleFormatText(formatTextArgs);
          
          case 'gdocs_set_heading':
            const headingArgs = {
              ...this.validateFormatTextArgs(request.params.arguments),
              headingLevel: (request.params.arguments as any)?.headingLevel
            };
            if (!headingArgs.headingLevel || headingArgs.headingLevel < 1 || headingArgs.headingLevel > 6) {
              throw new Error('Invalid headingLevel: expected number between 1 and 6');
            }
            return this.handleSetHeading(headingArgs);
          
          // Existing auth and drive handlers
          case 'gsheets_get_auth_url':
            return this.handleGetAuthUrl();
          
          case 'gsheets_set_auth_code':
            const code = (request.params.arguments as any)?.code;
            if (!code) {
              throw new Error('Authorization code is required');
            }
            return this.handleSetAuthCode(code);
          
          case 'gdrive_search':
            const searchArgs = this.validateDriveSearchArgs(request.params.arguments);
            return this.handleDriveSearch(searchArgs);
          
          case 'gdrive_read_file':
            const readArgs = this.validateDriveReadArgs(request.params.arguments);
            return this.handleDriveReadFile(readArgs);
          
          case 'gdrive_get_file_info':
            const infoArgs = this.validateDriveReadArgs(request.params.arguments);
            return this.handleDriveGetFileInfo(infoArgs);

          case 'gdrive_move_file':
            const moveArgs = this.validateMoveFileArgs(request.params.arguments);
            return this.handleDriveMoveFile(moveArgs);

          case 'gdrive_batch_move':
            const batchMoveArgs = this.validateBatchMoveArgs(request.params.arguments);
            return this.handleDriveBatchMove(batchMoveArgs);

          case 'gdrive_create_folder':
            const createFolderArgs = this.validateCreateFolderArgs(request.params.arguments);
            return this.handleDriveCreateFolder(createFolderArgs);

          case 'gdrive_copy_folder':
            const copyFolderArgs = this.validateCopyFolderArgs(request.params.arguments);
            return this.handleDriveCopyFolder(copyFolderArgs);

          case 'gdrive_search_content':
            const searchContentArgs = this.validateSearchContentArgs(request.params.arguments);
            return this.handleDriveSearchContent(searchContentArgs);

          case 'gdrive_get_revisions':
            const revisionsArgs = this.validateGetRevisionsArgs(request.params.arguments);
            return this.handleDriveGetRevisions(revisionsArgs);

          case 'gdrive_export_file':
            const exportArgs = this.validateExportFileArgs(request.params.arguments);
            return this.handleDriveExportFile(exportArgs);

          case 'gdrive_batch_export':
            const batchExportArgs = this.validateBatchExportArgs(request.params.arguments);
            return this.handleDriveBatchExport(batchExportArgs);

          case 'gdrive_list_folder_tree':
            const listTreeArgs = this.validateListFolderTreeArgs(request.params.arguments);
            return this.handleDriveListFolderTree(listTreeArgs);

          case 'gdrive_list_permissions':
            const permissionsArgs = this.validateListPermissionsArgs(request.params.arguments);
            return this.handleDriveListPermissions(permissionsArgs);

          default:
            throw new Error(`Unknown tool: ${request.params.name}`);
        }
      } catch (error) {
        return {
          content: [
            {
              type: 'text',
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    });
  }

  // Existing Google Sheets handlers (unchanged)
  private async handleGetAuthUrl() {
    try {
      const url = this.getAuthUrl();
      return {
        content: [
          {
            type: 'text',
            text: `Please visit this URL to authorize the application:\n\n${url}\n\nAfter authorizing, copy the authorization code and use the gsheets_set_auth_code tool.`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error generating auth URL: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleSetAuthCode(code: string) {
    try {
      const tokens = await this.getTokensFromCode(code);
      return {
        content: [
          {
            type: 'text',
            text: `Authorization successful! Please save this refresh token in your environment:\n\nGOOGLE_REFRESH_TOKEN=${tokens.refresh_token}\n\nRestart the MCP server with this environment variable set.`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error exchanging code for tokens: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDriveSearch(args: {
    query: string;
    pageSize: number;
    pageToken?: string;
  }) {
    try {
      const response = await this.drive.files.list({
        q: args.query,
        pageSize: args.pageSize,
        pageToken: args.pageToken,
        fields: 'nextPageToken, files(id, name, mimeType, size, modifiedTime, createdTime)',
      });

      const files = response.data.files || [];
      const fileList = files.map(file => 
        `${file.id} ${file.name} (${file.mimeType})`
      ).join('\n');

      let result = `Found ${files.length} files:\n${fileList}`;
      
      if (response.data.nextPageToken) {
        result += `\n\nMore results available. Use pageToken: ${response.data.nextPageToken}`;
      }

      return {
        content: [
          {
            type: 'text',
            text: result,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error searching Drive: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDriveReadFile(args: { fileId: string }) {
    try {
      // First get file metadata to determine type
      const fileInfo = await this.drive.files.get({
        fileId: args.fileId,
        fields: 'name, mimeType, size',
      });

      const mimeType = fileInfo.data.mimeType;
      const fileName = fileInfo.data.name;

      let content: string;

      // Handle different file types
      if (mimeType === 'application/vnd.google-apps.document') {
        // Google Docs - export as plain text
        const response = await this.drive.files.export({
          fileId: args.fileId,
          mimeType: 'text/plain',
        });
        content = response.data as string;
      } else if (mimeType === 'application/vnd.google-apps.spreadsheet') {
        // Google Sheets - export as CSV
        const response = await this.drive.files.export({
          fileId: args.fileId,
          mimeType: 'text/csv',
        });
        content = response.data as string;
      } else if (mimeType?.startsWith('text/') || mimeType === 'application/json') {
        // Text files - read directly
        const response = await this.drive.files.get({
          fileId: args.fileId,
          alt: 'media',
        });
        content = response.data as string;
      } else {
        throw new Error(`Unsupported file type: ${mimeType}. Can only read text files, Google Docs, and Google Sheets.`);
      }

      return {
        content: [
          {
            type: 'text',
            text: `Contents of ${fileName}:\n\n${content}`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error reading file: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDriveGetFileInfo(args: { fileId: string }) {
    try {
      const response = await this.drive.files.get({
        fileId: args.fileId,
        fields: 'id, name, mimeType, size, createdTime, modifiedTime, owners, permissions',
      });

      const file = response.data;
      const info = {
        id: file.id,
        name: file.name,
        mimeType: file.mimeType,
        size: file.size,
        createdTime: file.createdTime,
        modifiedTime: file.modifiedTime,
        owners: file.owners?.map(owner => owner.emailAddress),
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(info, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error getting file info: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDriveMoveFile(args: MoveFileRequest) {
    try {
      // Get current parents
      const file = await this.drive.files.get({
        fileId: args.fileId,
        fields: 'id, name, parents',
      });

      const previousParents = file.data.parents?.join(',');

      // Move to new folder
      const response = await this.drive.files.update({
        fileId: args.fileId,
        addParents: args.targetFolderId,
        removeParents: previousParents,
        fields: 'id, name, parents',
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              fileId: args.fileId,
              fileName: response.data.name,
              previousParents: previousParents?.split(',') || [],
              newParent: args.targetFolderId,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error moving file: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDriveBatchMove(args: BatchMoveRequest) {
    try {
      const results = {
        success: [] as string[],
        failed: [] as { fileId: string; error: string }[]
      };

      for (const fileId of args.fileIds) {
        try {
          // Get current parents
          const file = await this.drive.files.get({
            fileId,
            fields: 'id, name, parents',
          });

          const previousParents = file.data.parents?.join(',');

          // Move to new folder
          await this.drive.files.update({
            fileId,
            addParents: args.targetFolderId,
            removeParents: previousParents,
            fields: 'id, parents',
          });

          results.success.push(fileId);
        } catch (error) {
          results.failed.push({
            fileId,
            error: error instanceof Error ? error.message : String(error)
          });
        }
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              movedCount: results.success.length,
              failedCount: results.failed.length,
              targetFolderId: args.targetFolderId,
              details: results
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error in batch move: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDriveCreateFolder(args: CreateFolderRequest) {
    try {
      const fileMetadata: any = {
        name: args.name,
        mimeType: 'application/vnd.google-apps.folder'
      };

      if (args.parentFolderId) {
        fileMetadata.parents = [args.parentFolderId];
      }

      const response = await this.drive.files.create({
        requestBody: fileMetadata,
        fields: 'id, name, parents, webViewLink',
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              folderId: response.data.id,
              name: response.data.name,
              parentFolderId: args.parentFolderId || 'root',
              url: response.data.webViewLink,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error creating folder: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDriveCopyFolder(args: CopyFolderRequest) {
    try {
      // Get source folder info
      const sourceFolder = await this.drive.files.get({
        fileId: args.sourceFolderId,
        fields: 'id, name',
      });

      const newFolderName = args.newName || `Copy of ${sourceFolder.data.name}`;

      // Create new folder
      const newFolderMetadata: any = {
        name: newFolderName,
        mimeType: 'application/vnd.google-apps.folder'
      };

      if (args.targetParentFolderId) {
        newFolderMetadata.parents = [args.targetParentFolderId];
      }

      const newFolder = await this.drive.files.create({
        requestBody: newFolderMetadata,
        fields: 'id, name, webViewLink',
      });

      const newFolderId = newFolder.data.id!;

      // Get all files in source folder
      const files = await this.drive.files.list({
        q: `'${args.sourceFolderId}' in parents and trashed=false`,
        fields: 'files(id, name, mimeType)',
      });

      const copiedFiles = {
        folders: 0,
        files: 0,
        errors: [] as string[]
      };

      // Copy each file/folder
      for (const file of files.data.files || []) {
        try {
          if (file.mimeType === 'application/vnd.google-apps.folder') {
            // Recursively copy subfolder
            await this.handleDriveCopyFolder({
              sourceFolderId: file.id!,
              targetParentFolderId: newFolderId,
            });
            copiedFiles.folders++;
          } else {
            // Copy file - works for all file types including Google Workspace files
            const copiedFile = await this.drive.files.copy({
              fileId: file.id!,
              requestBody: {
                name: file.name,
                parents: [newFolderId]
              },
              fields: 'id, name',
            });

            if (copiedFile.data.id) {
              copiedFiles.files++;
            } else {
              copiedFiles.errors.push(`Failed to copy ${file.name}: No file ID returned`);
            }
          }
        } catch (error) {
          const errorMsg = error instanceof Error ? error.message : String(error);
          copiedFiles.errors.push(`Failed to copy ${file.name}: ${errorMsg}`);
        }
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              newFolderId: newFolderId,
              newFolderName: newFolderName,
              url: newFolder.data.webViewLink,
              sourceFilesFound: files.data.files?.length || 0,
              copiedFiles: copiedFiles.files,
              copiedFolders: copiedFiles.folders,
              errors: copiedFiles.errors.length > 0 ? copiedFiles.errors : undefined,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error copying folder: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDriveSearchContent(args: SearchContentRequest) {
    try {
      let query = `fullText contains '${args.query}' and trashed=false`;

      if (args.folderId) {
        query += ` and '${args.folderId}' in parents`;
      }

      const response = await this.drive.files.list({
        q: query,
        pageSize: args.maxResults,
        fields: 'files(id, name, mimeType, modifiedTime, webViewLink, owners)',
        orderBy: 'modifiedTime desc',
      });

      const files = response.data.files || [];

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              query: args.query,
              resultCount: files.length,
              results: files.map(f => ({
                id: f.id,
                name: f.name,
                type: f.mimeType,
                modified: f.modifiedTime,
                url: f.webViewLink,
                owner: f.owners?.[0]?.emailAddress
              }))
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error searching content: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDriveGetRevisions(args: GetRevisionsRequest) {
    try {
      const response = await this.drive.revisions.list({
        fileId: args.fileId,
        pageSize: args.maxResults,
        fields: 'revisions(id, modifiedTime, lastModifyingUser, size, originalFilename, keepForever)',
      });

      const revisions = response.data.revisions || [];

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              fileId: args.fileId,
              revisionCount: revisions.length,
              revisions: revisions.map(r => ({
                id: r.id,
                modifiedTime: r.modifiedTime,
                modifiedBy: r.lastModifyingUser?.displayName || r.lastModifyingUser?.emailAddress,
                size: r.size,
                filename: r.originalFilename,
                keepForever: r.keepForever
              }))
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error getting revisions: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDriveExportFile(args: ExportFileRequest) {
    try {
      const response = await this.drive.files.export({
        fileId: args.fileId,
        mimeType: args.mimeType,
      }, {
        responseType: 'arraybuffer'
      });

      // Convert to base64 for transmission
      const buffer = Buffer.from(response.data as ArrayBuffer);
      const base64 = buffer.toString('base64');
      const sizeInBytes = buffer.length;

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              fileId: args.fileId,
              mimeType: args.mimeType,
              sizeInBytes,
              sizeFormatted: `${(sizeInBytes / 1024).toFixed(2)} KB`,
              base64Data: base64.substring(0, 100) + '... (truncated for display)',
              note: 'Full base64 data available but truncated in output. Use this for downloads.'
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error exporting file: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDriveBatchExport(args: BatchExportRequest) {
    try {
      const mimeTypeMap = {
        pdf: 'application/pdf',
        docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
      };

      const mimeType = mimeTypeMap[args.format];
      const results = {
        success: [] as { fileId: string; size: number }[],
        failed: [] as { fileId: string; error: string }[]
      };

      for (const fileId of args.fileIds) {
        try {
          const response = await this.drive.files.export({
            fileId,
            mimeType,
          }, {
            responseType: 'arraybuffer'
          });

          const buffer = Buffer.from(response.data as ArrayBuffer);
          results.success.push({
            fileId,
            size: buffer.length
          });
        } catch (error) {
          results.failed.push({
            fileId,
            error: error instanceof Error ? error.message : String(error)
          });
        }
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              format: args.format,
              mimeType,
              exportedCount: results.success.length,
              failedCount: results.failed.length,
              totalSize: results.success.reduce((sum, r) => sum + r.size, 0),
              details: results
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error in batch export: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDriveListFolderTree(args: ListFolderTreeRequest) {
    try {
      const listFilesRecursive = async (folderId: string, path: string = ''): Promise<any[]> => {
        const query = `'${folderId}' in parents and trashed=false`;
        const fields = args.includeMetadata
          ? 'files(id, name, mimeType, size, createdTime, modifiedTime, owners, webViewLink)'
          : 'files(id, name, mimeType)';

        const response = await this.drive.files.list({
          q: query,
          fields,
          pageSize: 1000,
        });

        const files = response.data.files || [];
        const results: any[] = [];

        for (const file of files) {
          if (!file.id || !file.name) continue; // Skip files without ID or name

          const filePath = path ? `${path}/${file.name}` : file.name;
          const fileInfo: any = {
            id: file.id,
            name: file.name,
            path: filePath,
            type: file.mimeType,
          };

          if (args.includeMetadata) {
            fileInfo.size = file.size;
            fileInfo.created = file.createdTime;
            fileInfo.modified = file.modifiedTime;
            fileInfo.owner = file.owners?.[0]?.emailAddress;
            fileInfo.url = file.webViewLink;
          }

          results.push(fileInfo);

          // Recursively list subfolders
          if (args.recursive && file.mimeType === 'application/vnd.google-apps.folder') {
            const subFiles = await listFilesRecursive(file.id, filePath);
            results.push(...subFiles);
          }
        }

        return results;
      };

      const allFiles = await listFilesRecursive(args.folderId);

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              folderId: args.folderId,
              totalFiles: allFiles.length,
              recursive: args.recursive,
              includeMetadata: args.includeMetadata,
              files: allFiles
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error listing folder tree: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleDriveListPermissions(args: ListPermissionsRequest) {
    try {
      const response = await this.drive.permissions.list({
        fileId: args.fileId,
        fields: 'permissions(id, type, role, emailAddress, domain, displayName, expirationTime, deleted)',
      });

      const permissions = response.data.permissions || [];

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              fileId: args.fileId,
              permissionCount: permissions.length,
              permissions: permissions.map(p => ({
                id: p.id,
                type: p.type, // user, group, domain, anyone
                role: p.role, // owner, organizer, fileOrganizer, writer, commenter, reader
                email: p.emailAddress,
                domain: p.domain,
                displayName: p.displayName,
                expirationTime: p.expirationTime,
                deleted: p.deleted
              }))
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error listing permissions: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleBatchUpdate(args: BatchUpdateRequest) {
    try {
      const batchUpdateRequest: sheets_v4.Schema$BatchUpdateValuesRequest = {
        valueInputOption: 'USER_ENTERED',
        data: args.updates.map(update => ({
          range: update.range,
          values: update.values,
        })),
      };

      const response = await this.sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: args.spreadsheetId,
        requestBody: batchUpdateRequest,
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              updatedCells: response.data.totalUpdatedCells,
              updatedRows: response.data.totalUpdatedRows,
              updatedColumns: response.data.totalUpdatedColumns,
              updatedSheets: response.data.totalUpdatedSheets,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error updating spreadsheet: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleCreateAndPopulate(args: {
    title: string;
    sheetTitle: string;
    data: any[][];
    parentFolderId?: string;
  }) {
    try {
      // Create new spreadsheet
      const createResponse = await this.sheets.spreadsheets.create({
        requestBody: {
          properties: {
            title: args.title,
          },
          sheets: [
            {
              properties: {
                title: args.sheetTitle || 'Sheet1',
              },
            },
          ],
        },
      });

      const spreadsheetId = createResponse.data.spreadsheetId!;
      const sheetName = args.sheetTitle || 'Sheet1';

      // Populate with data
      if (args.data && args.data.length > 0) {
        await this.sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheetName}!A1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: {
            values: args.data,
          },
        });
      }

      // Move to parent folder if specified
      if (args.parentFolderId) {
        // Get current parents
        const file = await this.drive.files.get({
          fileId: spreadsheetId,
          fields: 'parents',
        });

        const previousParents = file.data.parents?.join(',');

        // Move to new folder
        await this.drive.files.update({
          fileId: spreadsheetId,
          addParents: args.parentFolderId,
          removeParents: previousParents,
          fields: 'id, parents',
        });
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              spreadsheetId,
              url: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
              rowsAdded: args.data?.length || 0,
              parentFolderId: args.parentFolderId || 'root',
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error creating spreadsheet: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleAppendRows(args: {
    spreadsheetId: string;
    range: string;
    values: any[][];
  }) {
    try {
      const response = await this.sheets.spreadsheets.values.append({
        spreadsheetId: args.spreadsheetId,
        range: args.range,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: args.values,
        },
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              updatedRange: response.data.updates?.updatedRange,
              updatedRows: response.data.updates?.updatedRows,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error appending rows: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleFormatCells(args: {
    spreadsheetId: string;
    requests: any[];
  }) {
    try {
      const requests = args.requests.map((req: any) => {
        // Convert A1 notation to GridRange
        const range = this.a1ToGridRange(req.range);
        
        return {
          repeatCell: {
            range,
            cell: {
              userEnteredFormat: req.format,
            },
            fields: Object.keys(req.format).map(key => `userEnteredFormat.${key}`).join(','),
          },
        };
      });

      const response = await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId: args.spreadsheetId,
        requestBody: {
          requests,
        },
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              appliedFormats: requests.length,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error formatting cells: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  // New Google Docs handlers
  private async handleCreateDocument(args: CreateDocumentRequest) {
    try {
      const response = await this.docs.documents.create({
        requestBody: {
          title: args.title,
        },
      });

      const documentId = response.data.documentId!;

      // Add initial content if provided
      if (args.content && args.content.trim()) {
        await this.docs.documents.batchUpdate({
          documentId,
          requestBody: {
            requests: [
              {
                insertText: {
                  location: {
                    index: 1,
                  },
                  text: args.content,
                },
              },
            ],
          },
        });
      }

      // Move to parent folder if specified
      if (args.parentFolderId) {
        // Get current parents
        const file = await this.drive.files.get({
          fileId: documentId,
          fields: 'parents',
        });

        const previousParents = file.data.parents?.join(',');

        // Move to new folder
        await this.drive.files.update({
          fileId: documentId,
          addParents: args.parentFolderId,
          removeParents: previousParents,
          fields: 'id, parents',
        });
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              documentId,
              url: `https://docs.google.com/document/d/${documentId}/edit`,
              title: args.title,
              parentFolderId: args.parentFolderId || 'root',
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error creating document: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleGetDocument(args: { documentId: string }) {
    try {
      const response = await this.docs.documents.get({
        documentId: args.documentId,
      });

      const doc = response.data;
      
      // Extract text content from the document
      let content = '';
      if (doc.body && doc.body.content) {
        for (const element of doc.body.content) {
          if (element.paragraph && element.paragraph.elements) {
            for (const elem of element.paragraph.elements) {
              if (elem.textRun && elem.textRun.content) {
                content += elem.textRun.content;
              }
            }
          }
        }
      }

      return {
        content: [
          {
            type: 'text',
            text: `Document: ${doc.title}\n\nContent:\n${content}`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error getting document: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleInsertText(args: InsertTextRequest) {
    try {
      // If no index specified, get document length to append at end
      let insertIndex = args.index;
      
      if (insertIndex === undefined) {
        const doc = await this.docs.documents.get({
          documentId: args.documentId,
        });
        
        // Find the end index of the document
        insertIndex = 1; // Start at beginning if we can't determine length
        if (doc.data.body && doc.data.body.content) {
          for (const element of doc.data.body.content) {
            if (element.endIndex) {
              insertIndex = Math.max(insertIndex, element.endIndex - 1);
            }
          }
        }
      }

      const response = await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              insertText: {
                location: {
                  index: insertIndex,
                },
                text: args.text,
              },
            },
          ],
        },
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              insertedAt: insertIndex,
              textLength: args.text.length,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error inserting text: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleAppendText(args: InsertTextRequest) {
    try {
      // Get document to find the end
      const doc = await this.docs.documents.get({
        documentId: args.documentId,
      });
      
      // Find the end index of the document
      let endIndex = 1;
      if (doc.data.body && doc.data.body.content) {
        for (const element of doc.data.body.content) {
          if (element.endIndex) {
            endIndex = Math.max(endIndex, element.endIndex - 1);
          }
        }
      }

      const response = await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              insertText: {
                location: {
                  index: endIndex,
                },
                text: args.text,
              },
            },
          ],
        },
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              appendedAt: endIndex,
              textLength: args.text.length,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error appending text: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleReplaceText(args: ReplaceTextRequest) {
    try {
      const response = await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              replaceAllText: {
                containsText: {
                  text: args.find,
                  matchCase: true,
                },
                replaceText: args.replace,
              },
            },
          ],
        },
      });

      const replaceCount = response.data.replies?.[0]?.replaceAllText?.occurrencesChanged || 0;

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              replacements: replaceCount,
              find: args.find,
              replace: args.replace,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error replacing text: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleFormatText(args: FormatTextRequest) {
    try {
      const textStyle: any = {};
      
      if (args.format.bold !== undefined) {
        textStyle.bold = args.format.bold;
      }
      if (args.format.italic !== undefined) {
        textStyle.italic = args.format.italic;
      }
      if (args.format.underline !== undefined) {
        textStyle.underline = args.format.underline;
      }
      if (args.format.fontSize !== undefined) {
        textStyle.fontSize = {
          magnitude: args.format.fontSize,
          unit: 'PT',
        };
      }

      const response = await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              updateTextStyle: {
                range: {
                  startIndex: args.startIndex,
                  endIndex: args.endIndex,
                },
                textStyle,
                fields: Object.keys(textStyle).join(','),
              },
            },
          ],
        },
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              formattedRange: `${args.startIndex}-${args.endIndex}`,
              appliedFormat: args.format,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error formatting text: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private async handleSetHeading(args: FormatTextRequest & { headingLevel: number }) {
    try {
      const namedStyleType = `HEADING_${args.headingLevel}` as any;

      const response = await this.docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: {
          requests: [
            {
              updateParagraphStyle: {
                range: {
                  startIndex: args.startIndex,
                  endIndex: args.endIndex,
                },
                paragraphStyle: {
                  namedStyleType,
                },
                fields: 'namedStyleType',
              },
            },
          ],
        },
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              headingLevel: args.headingLevel,
              formattedRange: `${args.startIndex}-${args.endIndex}`,
            }, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: 'text',
            text: `Error setting heading: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }

  private a1ToGridRange(a1Range: string): sheets_v4.Schema$GridRange {
    // Simple A1 to GridRange conversion
    // This is a simplified version - you'd want a more robust parser
    const match = a1Range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) {
      throw new Error(`Invalid A1 range: ${a1Range}`);
    }

    const [, startCol, startRow, endCol, endRow] = match;
    
    return {
      startRowIndex: parseInt(startRow) - 1,
      endRowIndex: parseInt(endRow),
      startColumnIndex: this.columnToIndex(startCol),
      endColumnIndex: this.columnToIndex(endCol) + 1,
    };
  }

  private columnToIndex(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result - 1;
  }

  async run() {
    try {
      console.error('Starting server transport...');
      const transport = new StdioServerTransport();
      await this.server.connect(transport);
      console.error('Server connected successfully');
    } catch (error) {
      console.error('Error starting server:', error);
      process.exit(1);
    }
  }
}

// Start the server with error handling
// ES module way to check if this is the main module
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

if (process.argv[1] === __filename) {
  console.error('Starting MCP server...');
  try {
    const server = new GoogleWorkspaceMCP();
    server.run().catch((error) => {
      console.error('Server run error:', error);
      process.exit(1);
    });
  } catch (error) {
    console.error('Server initialization error:', error);
    process.exit(1);
  }
}

export { GoogleWorkspaceMCP };