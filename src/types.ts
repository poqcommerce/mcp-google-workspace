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
    fontSize?: number;
  };
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
