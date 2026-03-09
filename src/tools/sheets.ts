import { sheets_v4, drive_v3 } from 'googleapis';
import type {
  BatchUpdateRequest,
  AppendRowsRequest,
  FormatCellsRequest,
  ToolDefinition,
  ToolResponse,
} from '../types.js';
import { a1ToGridRange, successResponse, errorResponse } from '../utils.js';

// ── Tool definitions ───────────────────────────────────────────────────────────

export function getSheetsToolDefinitions(): ToolDefinition[] {
  return [
    // ── Existing tools ─────────────────────────────────────────────────────
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
                    items: { type: 'string' },
                  },
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
              items: { type: 'string' },
            },
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
              items: { type: 'string' },
            },
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

    // ── Phase 1: Read operations ───────────────────────────────────────────
    {
      name: 'gsheets_read_data',
      description: 'Read cell data from a Google Spreadsheet',
      inputSchema: {
        type: 'object',
        properties: {
          spreadsheetId: {
            type: 'string',
            description: 'The ID of the spreadsheet',
          },
          range: {
            type: 'string',
            description: 'A1 notation range (e.g., "Sheet1!A1:D10" or "Sheet1"). Quote sheet names with spaces.',
          },
          valueRenderOption: {
            type: 'string',
            description: 'How to render values: FORMATTED_VALUE (default), UNFORMATTED_VALUE, or FORMULA',
            enum: ['FORMATTED_VALUE', 'UNFORMATTED_VALUE', 'FORMULA'],
          },
        },
        required: ['spreadsheetId', 'range'],
      },
    },
    {
      name: 'gsheets_get_spreadsheet_info',
      description: 'Get spreadsheet metadata including all sheet/tab names, IDs, and properties',
      inputSchema: {
        type: 'object',
        properties: {
          spreadsheetId: {
            type: 'string',
            description: 'The ID of the spreadsheet',
          },
        },
        required: ['spreadsheetId'],
      },
    },

    // ── Phase 2: Tab management ────────────────────────────────────────────
    {
      name: 'gsheets_add_sheet',
      description: 'Add a new tab/sheet to an existing spreadsheet',
      inputSchema: {
        type: 'object',
        properties: {
          spreadsheetId: {
            type: 'string',
            description: 'The ID of the spreadsheet',
          },
          title: {
            type: 'string',
            description: 'Name for the new tab',
          },
          index: {
            type: 'number',
            description: 'Position to insert the tab (0-based, optional)',
          },
          tabColor: {
            type: 'object',
            description: 'Tab color as { red, green, blue } with values 0-1 (optional)',
            properties: {
              red: { type: 'number' },
              green: { type: 'number' },
              blue: { type: 'number' },
            },
          },
          rowCount: {
            type: 'number',
            description: 'Initial row count (default: 1000)',
          },
          columnCount: {
            type: 'number',
            description: 'Initial column count (default: 26)',
          },
        },
        required: ['spreadsheetId', 'title'],
      },
    },
    {
      name: 'gsheets_delete_sheet',
      description: 'Delete a tab/sheet from a spreadsheet (permanent, cannot delete the last tab)',
      inputSchema: {
        type: 'object',
        properties: {
          spreadsheetId: {
            type: 'string',
            description: 'The ID of the spreadsheet',
          },
          sheetId: {
            type: 'number',
            description: 'Numeric sheet ID (from gsheets_get_spreadsheet_info), not the tab name',
          },
        },
        required: ['spreadsheetId', 'sheetId'],
      },
    },
    {
      name: 'gsheets_rename_sheet',
      description: 'Rename a tab or update its properties (color, frozen rows/cols, visibility)',
      inputSchema: {
        type: 'object',
        properties: {
          spreadsheetId: {
            type: 'string',
            description: 'The ID of the spreadsheet',
          },
          sheetId: {
            type: 'number',
            description: 'Numeric sheet ID (from gsheets_get_spreadsheet_info)',
          },
          title: {
            type: 'string',
            description: 'New name for the tab (optional)',
          },
          tabColor: {
            type: 'object',
            description: 'New tab color as { red, green, blue } with values 0-1 (optional)',
            properties: {
              red: { type: 'number' },
              green: { type: 'number' },
              blue: { type: 'number' },
            },
          },
          frozenRowCount: {
            type: 'number',
            description: 'Number of rows to freeze at the top (optional)',
          },
          frozenColumnCount: {
            type: 'number',
            description: 'Number of columns to freeze on the left (optional)',
          },
          hidden: {
            type: 'boolean',
            description: 'Whether to hide the tab (optional)',
          },
        },
        required: ['spreadsheetId', 'sheetId'],
      },
    },
    {
      name: 'gsheets_duplicate_sheet',
      description: 'Copy a tab within the same spreadsheet or to another spreadsheet',
      inputSchema: {
        type: 'object',
        properties: {
          spreadsheetId: {
            type: 'string',
            description: 'Source spreadsheet ID',
          },
          sheetId: {
            type: 'number',
            description: 'Numeric sheet ID of the tab to copy',
          },
          destinationSpreadsheetId: {
            type: 'string',
            description: 'Target spreadsheet ID (optional, defaults to same spreadsheet)',
          },
          newSheetName: {
            type: 'string',
            description: 'Name for the copied tab (optional)',
          },
          insertSheetIndex: {
            type: 'number',
            description: 'Position for the copy (0-based, optional)',
          },
        },
        required: ['spreadsheetId', 'sheetId'],
      },
    },

    // ── Phase 3: Dimensions & sort ─────────────────────────────────────────
    {
      name: 'gsheets_insert_delete_dimensions',
      description: 'Insert or delete rows or columns in a sheet',
      inputSchema: {
        type: 'object',
        properties: {
          spreadsheetId: {
            type: 'string',
            description: 'The ID of the spreadsheet',
          },
          sheetId: {
            type: 'number',
            description: 'Numeric sheet ID',
          },
          dimension: {
            type: 'string',
            description: 'ROWS or COLUMNS',
            enum: ['ROWS', 'COLUMNS'],
          },
          operation: {
            type: 'string',
            description: 'INSERT or DELETE',
            enum: ['INSERT', 'DELETE'],
          },
          startIndex: {
            type: 'number',
            description: 'Start index (0-based, inclusive)',
          },
          endIndex: {
            type: 'number',
            description: 'End index (0-based, exclusive)',
          },
          inheritFromBefore: {
            type: 'boolean',
            description: 'For inserts: inherit formatting from the row/column before (default: true)',
          },
        },
        required: ['spreadsheetId', 'sheetId', 'dimension', 'operation', 'startIndex', 'endIndex'],
      },
    },
    {
      name: 'gsheets_sort_range',
      description: 'Sort a range of cells by one or more columns',
      inputSchema: {
        type: 'object',
        properties: {
          spreadsheetId: {
            type: 'string',
            description: 'The ID of the spreadsheet',
          },
          sheetId: {
            type: 'number',
            description: 'Numeric sheet ID',
          },
          startRowIndex: {
            type: 'number',
            description: 'Start row (0-based, inclusive)',
          },
          endRowIndex: {
            type: 'number',
            description: 'End row (0-based, exclusive)',
          },
          startColumnIndex: {
            type: 'number',
            description: 'Start column (0-based, inclusive)',
          },
          endColumnIndex: {
            type: 'number',
            description: 'End column (0-based, exclusive)',
          },
          sortSpecs: {
            type: 'array',
            description: 'Sort specifications. dimensionIndex is relative to the range start column.',
            items: {
              type: 'object',
              properties: {
                dimensionIndex: {
                  type: 'number',
                  description: 'Column index to sort by (relative to startColumnIndex)',
                },
                sortOrder: {
                  type: 'string',
                  description: 'ASCENDING or DESCENDING',
                  enum: ['ASCENDING', 'DESCENDING'],
                },
              },
              required: ['dimensionIndex', 'sortOrder'],
            },
          },
        },
        required: ['spreadsheetId', 'sheetId', 'startRowIndex', 'endRowIndex', 'startColumnIndex', 'endColumnIndex', 'sortSpecs'],
      },
    },
  ];
}

// ── Handler class ──────────────────────────────────────────────────────────────

export class SheetsHandler {
  constructor(
    private sheets: sheets_v4.Sheets,
    private drive: drive_v3.Drive,
  ) {}

  /** Route a tool call to the appropriate handler. Returns null if not handled. */
  async handleTool(name: string, args: any): Promise<ToolResponse | null> {
    switch (name) {
      // Existing
      case 'gsheets_batch_update':
        return this.handleBatchUpdate(this.validateBatchUpdateArgs(args));
      case 'gsheets_create_and_populate':
        return this.handleCreateAndPopulate(this.validateCreateArgs(args));
      case 'gsheets_append_rows':
        return this.handleAppendRows(this.validateAppendArgs(args));
      case 'gsheets_format_cells':
        return this.handleFormatCells(this.validateFormatArgs(args));
      // Phase 1
      case 'gsheets_read_data':
        return this.handleReadData(this.validateReadDataArgs(args));
      case 'gsheets_get_spreadsheet_info':
        return this.handleGetSpreadsheetInfo(this.validateSpreadsheetIdArg(args));
      // Phase 2
      case 'gsheets_add_sheet':
        return this.handleAddSheet(this.validateAddSheetArgs(args));
      case 'gsheets_delete_sheet':
        return this.handleDeleteSheet(this.validateSheetIdArgs(args));
      case 'gsheets_rename_sheet':
        return this.handleRenameSheet(this.validateRenameSheetArgs(args));
      case 'gsheets_duplicate_sheet':
        return this.handleDuplicateSheet(this.validateDuplicateSheetArgs(args));
      // Phase 3
      case 'gsheets_insert_delete_dimensions':
        return this.handleInsertDeleteDimensions(this.validateDimensionArgs(args));
      case 'gsheets_sort_range':
        return this.handleSortRange(this.validateSortRangeArgs(args));
      default:
        return null;
    }
  }

  // ── Shared validators ──────────────────────────────────────────────────────

  private requireObject(args: any): asserts args is Record<string, any> {
    if (!args || typeof args !== 'object') {
      throw new Error('Invalid arguments: expected object');
    }
  }

  private requireString(args: any, field: string, label?: string): string {
    const value = args[field];
    if (!value || typeof value !== 'string') {
      throw new Error(`Invalid ${label || field}: expected non-empty string`);
    }
    return value;
  }

  private requireNumber(args: any, field: string, label?: string): number {
    const value = args[field];
    if (typeof value !== 'number') {
      throw new Error(`Invalid ${label || field}: expected number`);
    }
    return value;
  }

  private validateSpreadsheetIdArg(args: any): { spreadsheetId: string } {
    this.requireObject(args);
    return { spreadsheetId: this.requireString(args, 'spreadsheetId') };
  }

  private validateSheetIdArgs(args: any): { spreadsheetId: string; sheetId: number } {
    this.requireObject(args);
    return {
      spreadsheetId: this.requireString(args, 'spreadsheetId'),
      sheetId: this.requireNumber(args, 'sheetId'),
    };
  }

  // ── Existing validators ────────────────────────────────────────────────────

  private validateBatchUpdateArgs(args: any): BatchUpdateRequest {
    this.requireObject(args);
    const spreadsheetId = this.requireString(args, 'spreadsheetId');
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
      return { range: update.range, values: update.values };
    });

    return { spreadsheetId, updates };
  }

  private validateCreateArgs(args: any) {
    this.requireObject(args);
    const title = this.requireString(args, 'title');
    if (args.parentFolderId && typeof args.parentFolderId !== 'string') {
      throw new Error('Invalid parentFolderId: expected string');
    }
    return {
      title,
      sheetTitle: args.sheetTitle || 'Sheet1',
      data: args.data || [],
      parentFolderId: args.parentFolderId,
    };
  }

  private validateAppendArgs(args: any): AppendRowsRequest {
    this.requireObject(args);
    if (!Array.isArray(args.values)) {
      throw new Error('Invalid values: expected 2D array');
    }
    return {
      spreadsheetId: this.requireString(args, 'spreadsheetId'),
      range: this.requireString(args, 'range'),
      values: args.values,
    };
  }

  private validateFormatArgs(args: any): FormatCellsRequest {
    this.requireObject(args);
    if (!Array.isArray(args.requests)) {
      throw new Error('Invalid requests: expected array');
    }
    return {
      spreadsheetId: this.requireString(args, 'spreadsheetId'),
      requests: args.requests,
    };
  }

  // ── Phase 1 validators ─────────────────────────────────────────────────────

  private validateReadDataArgs(args: any) {
    this.requireObject(args);
    const validRenderOptions = ['FORMATTED_VALUE', 'UNFORMATTED_VALUE', 'FORMULA'];
    if (args.valueRenderOption && !validRenderOptions.includes(args.valueRenderOption)) {
      throw new Error(`Invalid valueRenderOption: expected one of ${validRenderOptions.join(', ')}`);
    }
    return {
      spreadsheetId: this.requireString(args, 'spreadsheetId'),
      range: this.requireString(args, 'range'),
      valueRenderOption: args.valueRenderOption || 'FORMATTED_VALUE',
    };
  }

  // ── Phase 2 validators ─────────────────────────────────────────────────────

  private validateAddSheetArgs(args: any) {
    this.requireObject(args);
    return {
      spreadsheetId: this.requireString(args, 'spreadsheetId'),
      title: this.requireString(args, 'title'),
      index: args.index as number | undefined,
      tabColor: args.tabColor as { red: number; green: number; blue: number } | undefined,
      rowCount: args.rowCount as number | undefined,
      columnCount: args.columnCount as number | undefined,
    };
  }

  private validateRenameSheetArgs(args: any) {
    this.requireObject(args);
    const base = this.validateSheetIdArgs(args);
    const hasUpdate =
      args.title !== undefined ||
      args.tabColor !== undefined ||
      args.frozenRowCount !== undefined ||
      args.frozenColumnCount !== undefined ||
      args.hidden !== undefined;
    if (!hasUpdate) {
      throw new Error('At least one property to update must be provided (title, tabColor, frozenRowCount, frozenColumnCount, hidden)');
    }
    return {
      ...base,
      title: args.title as string | undefined,
      tabColor: args.tabColor as { red: number; green: number; blue: number } | undefined,
      frozenRowCount: args.frozenRowCount as number | undefined,
      frozenColumnCount: args.frozenColumnCount as number | undefined,
      hidden: args.hidden as boolean | undefined,
    };
  }

  private validateDuplicateSheetArgs(args: any) {
    this.requireObject(args);
    return {
      ...this.validateSheetIdArgs(args),
      destinationSpreadsheetId: args.destinationSpreadsheetId as string | undefined,
      newSheetName: args.newSheetName as string | undefined,
      insertSheetIndex: args.insertSheetIndex as number | undefined,
    };
  }

  // ── Phase 3 validators ─────────────────────────────────────────────────────

  private validateDimensionArgs(args: any) {
    this.requireObject(args);
    const base = this.validateSheetIdArgs(args);
    const dimension = this.requireString(args, 'dimension');
    const operation = this.requireString(args, 'operation');
    if (!['ROWS', 'COLUMNS'].includes(dimension)) {
      throw new Error('Invalid dimension: expected ROWS or COLUMNS');
    }
    if (!['INSERT', 'DELETE'].includes(operation)) {
      throw new Error('Invalid operation: expected INSERT or DELETE');
    }
    const startIndex = this.requireNumber(args, 'startIndex');
    const endIndex = this.requireNumber(args, 'endIndex');
    if (endIndex <= startIndex) {
      throw new Error('Invalid range: endIndex must be greater than startIndex');
    }
    return {
      ...base,
      dimension: dimension as 'ROWS' | 'COLUMNS',
      operation: operation as 'INSERT' | 'DELETE',
      startIndex,
      endIndex,
      inheritFromBefore: args.inheritFromBefore !== false,
    };
  }

  private validateSortRangeArgs(args: any) {
    this.requireObject(args);
    const base = this.validateSheetIdArgs(args);
    if (!Array.isArray(args.sortSpecs) || args.sortSpecs.length === 0) {
      throw new Error('Invalid sortSpecs: expected non-empty array');
    }
    for (const spec of args.sortSpecs) {
      if (typeof spec.dimensionIndex !== 'number') {
        throw new Error('Invalid sortSpec: dimensionIndex must be a number');
      }
      if (!['ASCENDING', 'DESCENDING'].includes(spec.sortOrder)) {
        throw new Error('Invalid sortSpec: sortOrder must be ASCENDING or DESCENDING');
      }
    }
    return {
      ...base,
      startRowIndex: this.requireNumber(args, 'startRowIndex'),
      endRowIndex: this.requireNumber(args, 'endRowIndex'),
      startColumnIndex: this.requireNumber(args, 'startColumnIndex'),
      endColumnIndex: this.requireNumber(args, 'endColumnIndex'),
      sortSpecs: args.sortSpecs as { dimensionIndex: number; sortOrder: 'ASCENDING' | 'DESCENDING' }[],
    };
  }

  // ── Existing handlers ──────────────────────────────────────────────────────

  private async handleBatchUpdate(args: BatchUpdateRequest): Promise<ToolResponse> {
    try {
      const batchUpdateRequest: sheets_v4.Schema$BatchUpdateValuesRequest = {
        valueInputOption: 'USER_ENTERED',
        data: args.updates.map((update) => ({
          range: update.range,
          values: update.values,
        })),
      };

      const response = await this.sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: args.spreadsheetId,
        requestBody: batchUpdateRequest,
      });

      return successResponse({
        success: true,
        updatedCells: response.data.totalUpdatedCells,
        updatedRows: response.data.totalUpdatedRows,
        updatedColumns: response.data.totalUpdatedColumns,
        updatedSheets: response.data.totalUpdatedSheets,
      });
    } catch (error) {
      return errorResponse('updating spreadsheet', error);
    }
  }

  private async handleCreateAndPopulate(args: {
    title: string;
    sheetTitle: string;
    data: any[][];
    parentFolderId?: string;
  }): Promise<ToolResponse> {
    try {
      const createResponse = await this.sheets.spreadsheets.create({
        requestBody: {
          properties: { title: args.title },
          sheets: [{ properties: { title: args.sheetTitle || 'Sheet1' } }],
        },
      });

      const spreadsheetId = createResponse.data.spreadsheetId!;
      const sheetName = args.sheetTitle || 'Sheet1';

      if (args.data && args.data.length > 0) {
        await this.sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${sheetName}!A1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: args.data },
        });
      }

      if (args.parentFolderId) {
        const file = await this.drive.files.get({
          fileId: spreadsheetId,
          fields: 'parents',
        });
        const previousParents = file.data.parents?.join(',');
        await this.drive.files.update({
          fileId: spreadsheetId,
          addParents: args.parentFolderId,
          removeParents: previousParents,
          fields: 'id, parents',
        });
      }

      return successResponse({
        success: true,
        spreadsheetId,
        url: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
        rowsAdded: args.data?.length || 0,
        parentFolderId: args.parentFolderId || 'root',
      });
    } catch (error) {
      return errorResponse('creating spreadsheet', error);
    }
  }

  private async handleAppendRows(args: AppendRowsRequest): Promise<ToolResponse> {
    try {
      const response = await this.sheets.spreadsheets.values.append({
        spreadsheetId: args.spreadsheetId,
        range: args.range,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: args.values },
      });

      return successResponse({
        success: true,
        updatedRange: response.data.updates?.updatedRange,
        updatedRows: response.data.updates?.updatedRows,
      });
    } catch (error) {
      return errorResponse('appending rows', error);
    }
  }

  private async handleFormatCells(args: FormatCellsRequest): Promise<ToolResponse> {
    try {
      const requests = args.requests.map((req: any) => {
        const range = a1ToGridRange(req.range);
        return {
          repeatCell: {
            range,
            cell: { userEnteredFormat: req.format },
            fields: Object.keys(req.format)
              .map((key) => `userEnteredFormat.${key}`)
              .join(','),
          },
        };
      });

      await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId: args.spreadsheetId,
        requestBody: { requests },
      });

      return successResponse({
        success: true,
        appliedFormats: requests.length,
      });
    } catch (error) {
      return errorResponse('formatting cells', error);
    }
  }

  // ── Phase 1 handlers: Read operations ──────────────────────────────────────

  private async handleReadData(args: {
    spreadsheetId: string;
    range: string;
    valueRenderOption: string;
  }): Promise<ToolResponse> {
    try {
      const response = await this.sheets.spreadsheets.values.get({
        spreadsheetId: args.spreadsheetId,
        range: args.range,
        valueRenderOption: args.valueRenderOption,
      });

      const values = response.data.values || [];

      return successResponse({
        success: true,
        range: response.data.range,
        values,
        rowCount: values.length,
        columnCount: values.length > 0 ? Math.max(...values.map((r: any[]) => r.length)) : 0,
      });
    } catch (error) {
      return errorResponse('reading spreadsheet data', error);
    }
  }

  private async handleGetSpreadsheetInfo(args: { spreadsheetId: string }): Promise<ToolResponse> {
    try {
      const response = await this.sheets.spreadsheets.get({
        spreadsheetId: args.spreadsheetId,
        fields: 'spreadsheetId,properties.title,sheets.properties',
      });

      const sheets = response.data.sheets || [];

      return successResponse({
        success: true,
        spreadsheetId: response.data.spreadsheetId,
        title: response.data.properties?.title,
        sheetCount: sheets.length,
        sheets: sheets.map((s) => ({
          sheetId: s.properties?.sheetId,
          title: s.properties?.title,
          index: s.properties?.index,
          rowCount: s.properties?.gridProperties?.rowCount,
          columnCount: s.properties?.gridProperties?.columnCount,
          frozenRowCount: s.properties?.gridProperties?.frozenRowCount || 0,
          frozenColumnCount: s.properties?.gridProperties?.frozenColumnCount || 0,
          hidden: s.properties?.hidden || false,
          tabColor: s.properties?.tabColorStyle?.rgbColor || null,
        })),
      });
    } catch (error) {
      return errorResponse('getting spreadsheet info', error);
    }
  }

  // ── Phase 2 handlers: Tab management ───────────────────────────────────────

  private async handleAddSheet(args: {
    spreadsheetId: string;
    title: string;
    index?: number;
    tabColor?: { red: number; green: number; blue: number };
    rowCount?: number;
    columnCount?: number;
  }): Promise<ToolResponse> {
    try {
      const properties: sheets_v4.Schema$SheetProperties = {
        title: args.title,
      };

      if (args.index !== undefined) properties.index = args.index;
      if (args.tabColor) {
        properties.tabColorStyle = { rgbColor: args.tabColor };
      }

      const gridProperties: sheets_v4.Schema$GridProperties = {};
      if (args.rowCount) gridProperties.rowCount = args.rowCount;
      if (args.columnCount) gridProperties.columnCount = args.columnCount;
      if (Object.keys(gridProperties).length > 0) {
        properties.gridProperties = gridProperties;
      }

      const response = await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId: args.spreadsheetId,
        requestBody: {
          requests: [{ addSheet: { properties } }],
        },
      });

      const addedSheet = response.data.replies?.[0]?.addSheet?.properties;

      return successResponse({
        success: true,
        sheetId: addedSheet?.sheetId,
        title: addedSheet?.title,
        index: addedSheet?.index,
      });
    } catch (error) {
      return errorResponse('adding sheet', error);
    }
  }

  private async handleDeleteSheet(args: {
    spreadsheetId: string;
    sheetId: number;
  }): Promise<ToolResponse> {
    try {
      await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId: args.spreadsheetId,
        requestBody: {
          requests: [{ deleteSheet: { sheetId: args.sheetId } }],
        },
      });

      return successResponse({
        success: true,
        deletedSheetId: args.sheetId,
      });
    } catch (error) {
      return errorResponse('deleting sheet', error);
    }
  }

  private async handleRenameSheet(args: {
    spreadsheetId: string;
    sheetId: number;
    title?: string;
    tabColor?: { red: number; green: number; blue: number };
    frozenRowCount?: number;
    frozenColumnCount?: number;
    hidden?: boolean;
  }): Promise<ToolResponse> {
    try {
      const properties: sheets_v4.Schema$SheetProperties = {
        sheetId: args.sheetId,
      };
      const fieldParts: string[] = [];

      if (args.title !== undefined) {
        properties.title = args.title;
        fieldParts.push('title');
      }
      if (args.tabColor !== undefined) {
        properties.tabColorStyle = { rgbColor: args.tabColor };
        fieldParts.push('tabColorStyle.rgbColor');
      }
      if (args.hidden !== undefined) {
        properties.hidden = args.hidden;
        fieldParts.push('hidden');
      }

      const gridProperties: sheets_v4.Schema$GridProperties = {};
      if (args.frozenRowCount !== undefined) {
        gridProperties.frozenRowCount = args.frozenRowCount;
        fieldParts.push('gridProperties.frozenRowCount');
      }
      if (args.frozenColumnCount !== undefined) {
        gridProperties.frozenColumnCount = args.frozenColumnCount;
        fieldParts.push('gridProperties.frozenColumnCount');
      }
      if (Object.keys(gridProperties).length > 0) {
        properties.gridProperties = gridProperties;
      }

      await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId: args.spreadsheetId,
        requestBody: {
          requests: [
            {
              updateSheetProperties: {
                properties,
                fields: fieldParts.join(','),
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        sheetId: args.sheetId,
        updatedProperties: {
          ...(args.title !== undefined && { title: args.title }),
          ...(args.tabColor !== undefined && { tabColor: args.tabColor }),
          ...(args.frozenRowCount !== undefined && { frozenRowCount: args.frozenRowCount }),
          ...(args.frozenColumnCount !== undefined && { frozenColumnCount: args.frozenColumnCount }),
          ...(args.hidden !== undefined && { hidden: args.hidden }),
        },
      });
    } catch (error) {
      return errorResponse('updating sheet properties', error);
    }
  }

  private async handleDuplicateSheet(args: {
    spreadsheetId: string;
    sheetId: number;
    destinationSpreadsheetId?: string;
    newSheetName?: string;
    insertSheetIndex?: number;
  }): Promise<ToolResponse> {
    try {
      const destId = args.destinationSpreadsheetId || args.spreadsheetId;

      // Copy the sheet
      const response = await this.sheets.spreadsheets.sheets.copyTo({
        spreadsheetId: args.spreadsheetId,
        sheetId: args.sheetId,
        requestBody: {
          destinationSpreadsheetId: destId,
        },
      });

      const newSheetId = response.data.sheetId!;
      let finalTitle = response.data.title;

      // Rename and/or reposition if requested
      if (args.newSheetName || args.insertSheetIndex !== undefined) {
        const properties: sheets_v4.Schema$SheetProperties = { sheetId: newSheetId };
        const fieldParts: string[] = [];

        if (args.newSheetName) {
          properties.title = args.newSheetName;
          fieldParts.push('title');
          finalTitle = args.newSheetName;
        }
        if (args.insertSheetIndex !== undefined) {
          properties.index = args.insertSheetIndex;
          fieldParts.push('index');
        }

        await this.sheets.spreadsheets.batchUpdate({
          spreadsheetId: destId,
          requestBody: {
            requests: [
              {
                updateSheetProperties: {
                  properties,
                  fields: fieldParts.join(','),
                },
              },
            ],
          },
        });
      }

      return successResponse({
        success: true,
        newSheetId,
        title: finalTitle,
        destinationSpreadsheetId: destId,
      });
    } catch (error) {
      return errorResponse('duplicating sheet', error);
    }
  }

  // ── Phase 3 handlers: Dimensions & sort ────────────────────────────────────

  private async handleInsertDeleteDimensions(args: {
    spreadsheetId: string;
    sheetId: number;
    dimension: 'ROWS' | 'COLUMNS';
    operation: 'INSERT' | 'DELETE';
    startIndex: number;
    endIndex: number;
    inheritFromBefore: boolean;
  }): Promise<ToolResponse> {
    try {
      const dimensionRange = {
        sheetId: args.sheetId,
        dimension: args.dimension,
        startIndex: args.startIndex,
        endIndex: args.endIndex,
      };

      const request: sheets_v4.Schema$Request =
        args.operation === 'INSERT'
          ? { insertDimension: { range: dimensionRange, inheritFromBefore: args.inheritFromBefore } }
          : { deleteDimension: { range: dimensionRange } };

      await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId: args.spreadsheetId,
        requestBody: { requests: [request] },
      });

      return successResponse({
        success: true,
        operation: args.operation,
        dimension: args.dimension,
        startIndex: args.startIndex,
        endIndex: args.endIndex,
        count: args.endIndex - args.startIndex,
      });
    } catch (error) {
      return errorResponse(`${args.operation.toLowerCase()}ing ${args.dimension.toLowerCase()}`, error);
    }
  }

  private async handleSortRange(args: {
    spreadsheetId: string;
    sheetId: number;
    startRowIndex: number;
    endRowIndex: number;
    startColumnIndex: number;
    endColumnIndex: number;
    sortSpecs: { dimensionIndex: number; sortOrder: 'ASCENDING' | 'DESCENDING' }[];
  }): Promise<ToolResponse> {
    try {
      await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId: args.spreadsheetId,
        requestBody: {
          requests: [
            {
              sortRange: {
                range: {
                  sheetId: args.sheetId,
                  startRowIndex: args.startRowIndex,
                  endRowIndex: args.endRowIndex,
                  startColumnIndex: args.startColumnIndex,
                  endColumnIndex: args.endColumnIndex,
                },
                sortSpecs: args.sortSpecs,
              },
            },
          ],
        },
      });

      return successResponse({
        success: true,
        sortedRange: {
          sheetId: args.sheetId,
          startRowIndex: args.startRowIndex,
          endRowIndex: args.endRowIndex,
          startColumnIndex: args.startColumnIndex,
          endColumnIndex: args.endColumnIndex,
        },
        sortSpecs: args.sortSpecs,
      });
    } catch (error) {
      return errorResponse('sorting range', error);
    }
  }
}
