// Shared types for Google Workspace MCP tools

// ── Google Sheets ──────────────────────────────────────────────────────────────

export interface BatchUpdateRequest {
  spreadsheetId: string;
  updates: {
    range: string;
    values: any[][];
  }[];
}

export interface CreateSheetRequest {
  title: string;
  sheetTitle?: string;
  parentFolderId?: string;
}

export interface AppendRowsRequest {
  spreadsheetId: string;
  range: string;
  values: any[][];
}

export interface FormatCellsRequest {
  spreadsheetId: string;
  requests: any[];
}

// ── Google Docs ────────────────────────────────────────────────────────────────

export interface CreateDocumentRequest {
  title: string;
  content?: string;
  parentFolderId?: string;
}

export interface InsertTextRequest {
  documentId: string;
  text: string;
  index?: number;
}

export interface ReplaceTextRequest {
  documentId: string;
  find: string;
  replace: string;
}

export interface FormatTextRequest {
  documentId: string;
  startIndex: number;
  endIndex: number;
  format: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strikethrough?: boolean;
    fontSize?: number;
    fontFamily?: string;
    foregroundColor?: { red: number; green: number; blue: number };
    backgroundColor?: { red: number; green: number; blue: number };
    link?: { url: string };
  };
}

export interface FormatParagraphRequest {
  documentId: string;
  startIndex: number;
  endIndex: number;
  style: {
    namedStyleType?:
      | 'NORMAL_TEXT'
      | 'TITLE'
      | 'SUBTITLE'
      | 'HEADING_1'
      | 'HEADING_2'
      | 'HEADING_3'
      | 'HEADING_4'
      | 'HEADING_5'
      | 'HEADING_6';
    alignment?: 'START' | 'CENTER' | 'END' | 'JUSTIFIED';
    lineSpacing?: number; // e.g. 115 = 1.15x
    spaceAbove?: number; // points
    spaceBelow?: number; // points
    indentFirstLine?: number; // points
    indentStart?: number; // points
    indentEnd?: number; // points
    keepWithNext?: boolean;
  };
}

export interface SetHeadingRequest {
  documentId: string;
  startIndex: number;
  endIndex: number;
  headingLevel: number;
}

export interface CreateBulletsRequest {
  documentId: string;
  startIndex: number;
  endIndex: number;
  preset?: string; // e.g. BULLET_DISC_CIRCLE_SQUARE, NUMBERED_DECIMAL_ALPHA_ROMAN
}

export interface InsertTableRequest {
  documentId: string;
  index: number;
  rows: number;
  columns: number;
  cellContent?: string[][];
}

export interface UpdateTableRequest {
  documentId: string;
  tableStartIndex: number;
  rows: number;
  columns: number;
  columnWidths?: number[];
  headerRowBackgroundColor?: { red: number; green: number; blue: number };
  bodyRowBackgroundColor?: { red: number; green: number; blue: number };
  cellPadding?: number;
  borders?: 'NONE' | 'ALL';
  contentAlignment?: 'TOP' | 'MIDDLE' | 'BOTTOM';
}

export interface SetDocumentDefaultsRequest {
  documentId: string;
  fontFamily?: string;
  fontSize?: number;
  foregroundColor?: { red: number; green: number; blue: number };
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
}

export interface CreateFromTemplateRequest {
  templateId: string;
  title: string;
  replacements?: Record<string, string>;
  parentFolderId?: string;
}

// ── Google Drive ───────────────────────────────────────────────────────────────

export interface MoveFileRequest {
  fileId: string;
  targetFolderId: string;
}

export interface BatchMoveRequest {
  fileIds: string[];
  targetFolderId: string;
}

export interface CreateFolderRequest {
  name: string;
  parentFolderId?: string;
}

export interface CopyFolderRequest {
  sourceFolderId: string;
  targetParentFolderId?: string;
  newName?: string;
}

export interface SearchContentRequest {
  query: string;
  folderId?: string;
  maxResults?: number;
}

export interface GetRevisionsRequest {
  fileId: string;
  maxResults?: number;
}

export interface ExportFileRequest {
  fileId: string;
  mimeType: string;
}

export interface BatchExportRequest {
  fileIds: string[];
  format: 'pdf' | 'docx' | 'xlsx' | 'pptx';
}

export interface ListFolderTreeRequest {
  folderId: string;
  recursive?: boolean;
  includeMetadata?: boolean;
}

export interface CopyFileRequest {
  fileId: string;
  name?: string;
  parentFolderId?: string;
}

export interface ListPermissionsRequest {
  fileId: string;
}

export interface ListCommentsRequest {
  fileId: string;
  includeDeleted?: boolean;
  includeResolved?: boolean;
  pageSize?: number;
  pageToken?: string;
}

export interface CreateCommentRequest {
  fileId: string;
  content: string;
}

export interface ReplyToCommentRequest {
  fileId: string;
  commentId: string;
  content: string;
  action?: 'resolve' | 'reopen';
}

export interface DeleteCommentRequest {
  fileId: string;
  commentId: string;
}

export interface SuggestionActivityRequest {
  fileId: string;
}

// ── Shared ─────────────────────────────────────────────────────────────────────

export interface ToolResponse {
  [key: string]: unknown;
  content: { type: string; text: string }[];
  isError?: boolean;
}

export interface ToolDefinition {
  name: string;
  description: string;
  inputSchema: Record<string, any>;
}
