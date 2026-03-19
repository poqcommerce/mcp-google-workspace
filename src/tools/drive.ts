import { drive_v3 } from 'googleapis';
import type {
  MoveFileRequest,
  BatchMoveRequest,
  CreateFolderRequest,
  CopyFolderRequest,
  CopyFileRequest,
  SearchContentRequest,
  GetRevisionsRequest,
  ExportFileRequest,
  BatchExportRequest,
  ListFolderTreeRequest,
  ListPermissionsRequest,
  ListCommentsRequest,
  CreateCommentRequest,
  ReplyToCommentRequest,
  DeleteCommentRequest,
  ToolDefinition,
  ToolResponse,
} from '../types.js';
import { successResponse, textResponse, errorResponse } from '../utils.js';

// ── Tool definitions ───────────────────────────────────────────────────────────

export function getDriveToolDefinitions(): ToolDefinition[] {
  return [
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
          pageSize: { type: 'number', description: 'Number of results per page (max 100)', default: 10 },
          pageToken: { type: 'string', description: 'Token for the next page of results' },
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
          fileId: { type: 'string', description: 'ID of the file to read' },
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
          fileId: { type: 'string', description: 'ID of the file to get info for' },
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
          fileId: { type: 'string', description: 'ID of the file to move' },
          targetFolderId: { type: 'string', description: 'ID of the target folder to move the file to' },
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
          fileIds: { type: 'array', description: 'Array of file IDs to move', items: { type: 'string' } },
          targetFolderId: { type: 'string', description: 'ID of the target folder' },
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
          name: { type: 'string', description: 'Name of the folder to create' },
          parentFolderId: { type: 'string', description: 'ID of the parent folder (optional, defaults to root)' },
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
          sourceFolderId: { type: 'string', description: 'ID of the folder to copy' },
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
      name: 'gdrive_copy_file',
      description: 'Copy a file in Google Drive (preserves themes, formatting, etc.)',
      inputSchema: {
        type: 'object',
        properties: {
          fileId: { type: 'string', description: 'ID of the file to copy' },
          name: { type: 'string', description: 'Name for the copy (optional, defaults to "Copy of [original]")' },
          parentFolderId: {
            type: 'string',
            description: 'ID of the parent folder for the copy (optional, defaults to same location)',
          },
        },
        required: ['fileId'],
      },
    },
    {
      name: 'gdrive_search_content',
      description: 'Search for text within file contents (fullText search)',
      inputSchema: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Text to search for within file contents' },
          folderId: { type: 'string', description: 'Limit search to specific folder (optional)' },
          maxResults: { type: 'number', description: 'Maximum number of results to return (default: 100)' },
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
          fileId: { type: 'string', description: 'ID of the file' },
          maxResults: { type: 'number', description: 'Maximum number of revisions to return (default: 20)' },
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
          fileId: { type: 'string', description: 'ID of the file to export' },
          mimeType: {
            type: 'string',
            description:
              'Target MIME type (e.g., "application/pdf", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")',
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
          fileIds: { type: 'array', description: 'Array of file IDs to export', items: { type: 'string' } },
          format: {
            type: 'string',
            description: 'Export format: pdf, docx, xlsx, or pptx',
            enum: ['pdf', 'docx', 'xlsx', 'pptx'],
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
          folderId: { type: 'string', description: 'ID of the folder to list' },
          recursive: { type: 'boolean', description: 'Recursively list subfolders (default: true)' },
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
          fileId: { type: 'string', description: 'ID of the file or folder' },
        },
        required: ['fileId'],
      },
    },
    {
      name: 'gdrive_list_comments',
      description: 'List all comments on a Google Drive file (Docs, Sheets, or Slides)',
      inputSchema: {
        type: 'object',
        properties: {
          fileId: { type: 'string', description: 'ID of the file' },
          includeDeleted: { type: 'boolean', description: 'Include deleted comments (default: false)' },
          includeResolved: { type: 'boolean', description: 'Include resolved comments (default: true)' },
          pageSize: { type: 'number', description: 'Number of comments per page (max 100, default: 100)' },
          pageToken: { type: 'string', description: 'Token for the next page of results' },
        },
        required: ['fileId'],
      },
    },
    {
      name: 'gdrive_create_comment',
      description: 'Add a comment to a Google Drive file (Docs, Sheets, or Slides)',
      inputSchema: {
        type: 'object',
        properties: {
          fileId: { type: 'string', description: 'ID of the file to comment on' },
          content: { type: 'string', description: 'Comment text' },
        },
        required: ['fileId', 'content'],
      },
    },
    {
      name: 'gdrive_reply_to_comment',
      description: 'Reply to a comment, or resolve/reopen it',
      inputSchema: {
        type: 'object',
        properties: {
          fileId: { type: 'string', description: 'ID of the file' },
          commentId: { type: 'string', description: 'ID of the comment to reply to' },
          content: { type: 'string', description: 'Reply text' },
          action: {
            type: 'string',
            description: 'Optional action: "resolve" to resolve the comment, "reopen" to reopen it',
            enum: ['resolve', 'reopen'],
          },
        },
        required: ['fileId', 'commentId', 'content'],
      },
    },
    {
      name: 'gdrive_delete_comment',
      description: 'Delete a comment (only works if you are the comment author)',
      inputSchema: {
        type: 'object',
        properties: {
          fileId: { type: 'string', description: 'ID of the file' },
          commentId: { type: 'string', description: 'ID of the comment to delete' },
        },
        required: ['fileId', 'commentId'],
      },
    },
  ];
}

// ── Handler class ──────────────────────────────────────────────────────────────

export class DriveHandler {
  constructor(private drive: drive_v3.Drive) {}

  /** Route a tool call to the appropriate handler. Returns null if not handled. */
  async handleTool(name: string, args: any): Promise<ToolResponse | null> {
    switch (name) {
      case 'gdrive_search':
        return this.handleSearch(this.validateSearchArgs(args));
      case 'gdrive_read_file':
        return this.handleReadFile(this.validateFileIdArgs(args));
      case 'gdrive_get_file_info':
        return this.handleGetFileInfo(this.validateFileIdArgs(args));
      case 'gdrive_move_file':
        return this.handleMoveFile(this.validateMoveFileArgs(args));
      case 'gdrive_batch_move':
        return this.handleBatchMove(this.validateBatchMoveArgs(args));
      case 'gdrive_create_folder':
        return this.handleCreateFolder(this.validateCreateFolderArgs(args));
      case 'gdrive_copy_folder':
        return this.handleCopyFolder(this.validateCopyFolderArgs(args));
      case 'gdrive_copy_file':
        return this.handleCopyFile(this.validateCopyFileArgs(args));
      case 'gdrive_search_content':
        return this.handleSearchContent(this.validateSearchContentArgs(args));
      case 'gdrive_get_revisions':
        return this.handleGetRevisions(this.validateGetRevisionsArgs(args));
      case 'gdrive_export_file':
        return this.handleExportFile(this.validateExportFileArgs(args));
      case 'gdrive_batch_export':
        return this.handleBatchExport(this.validateBatchExportArgs(args));
      case 'gdrive_list_folder_tree':
        return this.handleListFolderTree(this.validateListFolderTreeArgs(args));
      case 'gdrive_list_permissions':
        return this.handleListPermissions(this.validateListPermissionsArgs(args));
      case 'gdrive_list_comments':
        return this.handleListComments(this.validateListCommentsArgs(args));
      case 'gdrive_create_comment':
        return this.handleCreateComment(this.validateCreateCommentArgs(args));
      case 'gdrive_reply_to_comment':
        return this.handleReplyToComment(this.validateReplyToCommentArgs(args));
      case 'gdrive_delete_comment':
        return this.handleDeleteComment(this.validateDeleteCommentArgs(args));
      default:
        return null;
    }
  }

  // ── Validators ─────────────────────────────────────────────────────────────

  private validateSearchArgs(args: any) {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.query || typeof args.query !== 'string') throw new Error('Invalid query: expected non-empty string');
    return { query: args.query, pageSize: args.pageSize || 10, pageToken: args.pageToken || undefined };
  }

  private validateFileIdArgs(args: any) {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.fileId || typeof args.fileId !== 'string') throw new Error('Invalid fileId: expected non-empty string');
    return { fileId: args.fileId };
  }

  private validateMoveFileArgs(args: any): MoveFileRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.fileId || typeof args.fileId !== 'string') throw new Error('Invalid fileId: expected non-empty string');
    if (!args.targetFolderId || typeof args.targetFolderId !== 'string')
      throw new Error('Invalid targetFolderId: expected non-empty string');
    return { fileId: args.fileId, targetFolderId: args.targetFolderId };
  }

  private validateBatchMoveArgs(args: any): BatchMoveRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!Array.isArray(args.fileIds) || args.fileIds.length === 0)
      throw new Error('Invalid fileIds: expected non-empty array');
    if (!args.targetFolderId || typeof args.targetFolderId !== 'string')
      throw new Error('Invalid targetFolderId: expected non-empty string');
    return { fileIds: args.fileIds, targetFolderId: args.targetFolderId };
  }

  private validateCreateFolderArgs(args: any): CreateFolderRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.name || typeof args.name !== 'string') throw new Error('Invalid name: expected non-empty string');
    if (args.parentFolderId && typeof args.parentFolderId !== 'string')
      throw new Error('Invalid parentFolderId: expected string');
    return { name: args.name, parentFolderId: args.parentFolderId };
  }

  private validateCopyFolderArgs(args: any): CopyFolderRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.sourceFolderId || typeof args.sourceFolderId !== 'string')
      throw new Error('Invalid sourceFolderId: expected non-empty string');
    if (args.targetParentFolderId && typeof args.targetParentFolderId !== 'string')
      throw new Error('Invalid targetParentFolderId: expected string');
    if (args.newName && typeof args.newName !== 'string') throw new Error('Invalid newName: expected string');
    return {
      sourceFolderId: args.sourceFolderId,
      targetParentFolderId: args.targetParentFolderId,
      newName: args.newName,
    };
  }

  private validateCopyFileArgs(args: any): CopyFileRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.fileId || typeof args.fileId !== 'string') throw new Error('Invalid fileId: expected non-empty string');
    if (args.name && typeof args.name !== 'string') throw new Error('Invalid name: expected string');
    if (args.parentFolderId && typeof args.parentFolderId !== 'string')
      throw new Error('Invalid parentFolderId: expected string');
    return { fileId: args.fileId, name: args.name, parentFolderId: args.parentFolderId };
  }

  private validateSearchContentArgs(args: any): SearchContentRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.query || typeof args.query !== 'string') throw new Error('Invalid query: expected non-empty string');
    if (args.folderId && typeof args.folderId !== 'string') throw new Error('Invalid folderId: expected string');
    if (args.maxResults && typeof args.maxResults !== 'number') throw new Error('Invalid maxResults: expected number');
    return { query: args.query, folderId: args.folderId, maxResults: args.maxResults || 100 };
  }

  private validateGetRevisionsArgs(args: any): GetRevisionsRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.fileId || typeof args.fileId !== 'string') throw new Error('Invalid fileId: expected non-empty string');
    if (args.maxResults && typeof args.maxResults !== 'number') throw new Error('Invalid maxResults: expected number');
    return { fileId: args.fileId, maxResults: args.maxResults || 20 };
  }

  private validateExportFileArgs(args: any): ExportFileRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.fileId || typeof args.fileId !== 'string') throw new Error('Invalid fileId: expected non-empty string');
    if (!args.mimeType || typeof args.mimeType !== 'string')
      throw new Error('Invalid mimeType: expected non-empty string');
    return { fileId: args.fileId, mimeType: args.mimeType };
  }

  private validateBatchExportArgs(args: any): BatchExportRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!Array.isArray(args.fileIds) || args.fileIds.length === 0)
      throw new Error('Invalid fileIds: expected non-empty array');
    if (!args.format || !['pdf', 'docx', 'xlsx', 'pptx'].includes(args.format))
      throw new Error('Invalid format: expected one of: pdf, docx, xlsx, pptx');
    return { fileIds: args.fileIds, format: args.format };
  }

  private validateListFolderTreeArgs(args: any): ListFolderTreeRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.folderId || typeof args.folderId !== 'string')
      throw new Error('Invalid folderId: expected non-empty string');
    return {
      folderId: args.folderId,
      recursive: args.recursive !== false,
      includeMetadata: args.includeMetadata === true,
    };
  }

  private validateListPermissionsArgs(args: any): ListPermissionsRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.fileId || typeof args.fileId !== 'string') throw new Error('Invalid fileId: expected non-empty string');
    return { fileId: args.fileId };
  }

  private validateListCommentsArgs(args: any): ListCommentsRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.fileId || typeof args.fileId !== 'string') throw new Error('Invalid fileId: expected non-empty string');
    return {
      fileId: args.fileId,
      includeDeleted: args.includeDeleted === true,
      includeResolved: args.includeResolved !== false,
      pageSize: args.pageSize || 100,
      pageToken: args.pageToken,
    };
  }

  private validateCreateCommentArgs(args: any): CreateCommentRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.fileId || typeof args.fileId !== 'string') throw new Error('Invalid fileId: expected non-empty string');
    if (!args.content || typeof args.content !== 'string') throw new Error('Invalid content: expected non-empty string');
    return { fileId: args.fileId, content: args.content };
  }

  private validateReplyToCommentArgs(args: any): ReplyToCommentRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.fileId || typeof args.fileId !== 'string') throw new Error('Invalid fileId: expected non-empty string');
    if (!args.commentId || typeof args.commentId !== 'string') throw new Error('Invalid commentId: expected non-empty string');
    if (!args.content || typeof args.content !== 'string') throw new Error('Invalid content: expected non-empty string');
    if (args.action && !['resolve', 'reopen'].includes(args.action)) throw new Error('Invalid action: expected "resolve" or "reopen"');
    return { fileId: args.fileId, commentId: args.commentId, content: args.content, action: args.action };
  }

  private validateDeleteCommentArgs(args: any): DeleteCommentRequest {
    if (!args || typeof args !== 'object') throw new Error('Invalid arguments: expected object');
    if (!args.fileId || typeof args.fileId !== 'string') throw new Error('Invalid fileId: expected non-empty string');
    if (!args.commentId || typeof args.commentId !== 'string') throw new Error('Invalid commentId: expected non-empty string');
    return { fileId: args.fileId, commentId: args.commentId };
  }

  // ── Handlers ───────────────────────────────────────────────────────────────

  private async handleSearch(args: {
    query: string;
    pageSize: number;
    pageToken?: string;
  }): Promise<ToolResponse> {
    try {
      const response = await this.drive.files.list({
        q: args.query,
        pageSize: args.pageSize,
        pageToken: args.pageToken,
        fields: 'nextPageToken, files(id, name, mimeType, size, modifiedTime, createdTime)',
      });

      const files = response.data.files || [];
      const fileList = files.map((file) => `${file.id} ${file.name} (${file.mimeType})`).join('\n');

      let result = `Found ${files.length} files:\n${fileList}`;
      if (response.data.nextPageToken) {
        result += `\n\nMore results available. Use pageToken: ${response.data.nextPageToken}`;
      }

      return textResponse(result);
    } catch (error) {
      return errorResponse('searching Drive', error);
    }
  }

  private async handleReadFile(args: { fileId: string }): Promise<ToolResponse> {
    try {
      const fileInfo = await this.drive.files.get({
        fileId: args.fileId,
        fields: 'name, mimeType, size',
      });

      const mimeType = fileInfo.data.mimeType;
      const fileName = fileInfo.data.name;
      let content: string;

      if (mimeType === 'application/vnd.google-apps.document') {
        const response = await this.drive.files.export({ fileId: args.fileId, mimeType: 'text/plain' });
        content = response.data as string;
      } else if (mimeType === 'application/vnd.google-apps.spreadsheet') {
        const response = await this.drive.files.export({ fileId: args.fileId, mimeType: 'text/csv' });
        content = response.data as string;
      } else if (mimeType?.startsWith('text/') || mimeType === 'application/json') {
        const response = await this.drive.files.get({ fileId: args.fileId, alt: 'media' });
        content = response.data as string;
      } else {
        throw new Error(
          `Unsupported file type: ${mimeType}. Can only read text files, Google Docs, and Google Sheets.`,
        );
      }

      return textResponse(`Contents of ${fileName}:\n\n${content}`);
    } catch (error) {
      return errorResponse('reading file', error);
    }
  }

  private async handleGetFileInfo(args: { fileId: string }): Promise<ToolResponse> {
    try {
      const response = await this.drive.files.get({
        fileId: args.fileId,
        fields: 'id, name, mimeType, size, createdTime, modifiedTime, owners, permissions',
      });

      const file = response.data;
      return successResponse({
        id: file.id,
        name: file.name,
        mimeType: file.mimeType,
        size: file.size,
        createdTime: file.createdTime,
        modifiedTime: file.modifiedTime,
        owners: file.owners?.map((owner) => owner.emailAddress),
      });
    } catch (error) {
      return errorResponse('getting file info', error);
    }
  }

  private async handleMoveFile(args: MoveFileRequest): Promise<ToolResponse> {
    try {
      const file = await this.drive.files.get({ fileId: args.fileId, fields: 'id, name, parents' });
      const previousParents = file.data.parents?.join(',');

      const response = await this.drive.files.update({
        fileId: args.fileId,
        addParents: args.targetFolderId,
        removeParents: previousParents,
        fields: 'id, name, parents',
      });

      return successResponse({
        success: true,
        fileId: args.fileId,
        fileName: response.data.name,
        previousParents: previousParents?.split(',') || [],
        newParent: args.targetFolderId,
      });
    } catch (error) {
      return errorResponse('moving file', error);
    }
  }

  private async handleBatchMove(args: BatchMoveRequest): Promise<ToolResponse> {
    try {
      const results = {
        success: [] as string[],
        failed: [] as { fileId: string; error: string }[],
      };

      for (const fileId of args.fileIds) {
        try {
          const file = await this.drive.files.get({ fileId, fields: 'id, name, parents' });
          const previousParents = file.data.parents?.join(',');
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
            error: error instanceof Error ? error.message : String(error),
          });
        }
      }

      return successResponse({
        success: true,
        movedCount: results.success.length,
        failedCount: results.failed.length,
        targetFolderId: args.targetFolderId,
        details: results,
      });
    } catch (error) {
      return errorResponse('in batch move', error);
    }
  }

  private async handleCreateFolder(args: CreateFolderRequest): Promise<ToolResponse> {
    try {
      const fileMetadata: any = {
        name: args.name,
        mimeType: 'application/vnd.google-apps.folder',
      };
      if (args.parentFolderId) {
        fileMetadata.parents = [args.parentFolderId];
      }

      const response = await this.drive.files.create({
        requestBody: fileMetadata,
        fields: 'id, name, parents, webViewLink',
      });

      return successResponse({
        success: true,
        folderId: response.data.id,
        name: response.data.name,
        parentFolderId: args.parentFolderId || 'root',
        url: response.data.webViewLink,
      });
    } catch (error) {
      return errorResponse('creating folder', error);
    }
  }

  private async handleCopyFolder(args: CopyFolderRequest): Promise<ToolResponse> {
    try {
      const sourceFolder = await this.drive.files.get({ fileId: args.sourceFolderId, fields: 'id, name' });
      const newFolderName = args.newName || `Copy of ${sourceFolder.data.name}`;

      const newFolderMetadata: any = {
        name: newFolderName,
        mimeType: 'application/vnd.google-apps.folder',
      };
      if (args.targetParentFolderId) {
        newFolderMetadata.parents = [args.targetParentFolderId];
      }

      const newFolder = await this.drive.files.create({
        requestBody: newFolderMetadata,
        fields: 'id, name, webViewLink',
      });
      const newFolderId = newFolder.data.id!;

      const files = await this.drive.files.list({
        q: `'${args.sourceFolderId}' in parents and trashed=false`,
        fields: 'files(id, name, mimeType)',
      });

      const copiedFiles = { folders: 0, files: 0, errors: [] as string[] };

      for (const file of files.data.files || []) {
        try {
          if (file.mimeType === 'application/vnd.google-apps.folder') {
            await this.handleCopyFolder({
              sourceFolderId: file.id!,
              targetParentFolderId: newFolderId,
            });
            copiedFiles.folders++;
          } else {
            const copiedFile = await this.drive.files.copy({
              fileId: file.id!,
              requestBody: { name: file.name, parents: [newFolderId] },
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

      return successResponse({
        success: true,
        newFolderId,
        newFolderName,
        url: newFolder.data.webViewLink,
        sourceFilesFound: files.data.files?.length || 0,
        copiedFiles: copiedFiles.files,
        copiedFolders: copiedFiles.folders,
        errors: copiedFiles.errors.length > 0 ? copiedFiles.errors : undefined,
      });
    } catch (error) {
      return errorResponse('copying folder', error);
    }
  }

  private async handleCopyFile(args: CopyFileRequest): Promise<ToolResponse> {
    try {
      const requestBody: any = {};
      if (args.name) requestBody.name = args.name;
      if (args.parentFolderId) requestBody.parents = [args.parentFolderId];

      const response = await this.drive.files.copy({
        fileId: args.fileId,
        requestBody,
        fields: 'id, name, mimeType, parents, webViewLink',
      });

      return successResponse({
        success: true,
        sourceFileId: args.fileId,
        copiedFileId: response.data.id,
        name: response.data.name,
        mimeType: response.data.mimeType,
        url: response.data.webViewLink,
      });
    } catch (error) {
      return errorResponse('copying file', error);
    }
  }

  private async handleSearchContent(args: SearchContentRequest): Promise<ToolResponse> {
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
      return successResponse({
        success: true,
        query: args.query,
        resultCount: files.length,
        results: files.map((f) => ({
          id: f.id,
          name: f.name,
          type: f.mimeType,
          modified: f.modifiedTime,
          url: f.webViewLink,
          owner: f.owners?.[0]?.emailAddress,
        })),
      });
    } catch (error) {
      return errorResponse('searching content', error);
    }
  }

  private async handleGetRevisions(args: GetRevisionsRequest): Promise<ToolResponse> {
    try {
      const response = await this.drive.revisions.list({
        fileId: args.fileId,
        pageSize: args.maxResults,
        fields: 'revisions(id, modifiedTime, lastModifyingUser, size, originalFilename, keepForever)',
      });

      const revisions = response.data.revisions || [];
      return successResponse({
        success: true,
        fileId: args.fileId,
        revisionCount: revisions.length,
        revisions: revisions.map((r) => ({
          id: r.id,
          modifiedTime: r.modifiedTime,
          modifiedBy: r.lastModifyingUser?.displayName || r.lastModifyingUser?.emailAddress,
          size: r.size,
          filename: r.originalFilename,
          keepForever: r.keepForever,
        })),
      });
    } catch (error) {
      return errorResponse('getting revisions', error);
    }
  }

  private async handleExportFile(args: ExportFileRequest): Promise<ToolResponse> {
    try {
      const response = await this.drive.files.export(
        { fileId: args.fileId, mimeType: args.mimeType },
        { responseType: 'arraybuffer' },
      );

      const buffer = Buffer.from(response.data as ArrayBuffer);
      const base64 = buffer.toString('base64');
      const sizeInBytes = buffer.length;

      return successResponse({
        success: true,
        fileId: args.fileId,
        mimeType: args.mimeType,
        sizeInBytes,
        sizeFormatted: `${(sizeInBytes / 1024).toFixed(2)} KB`,
        base64Data: base64.substring(0, 100) + '... (truncated for display)',
        note: 'Full base64 data available but truncated in output. Use this for downloads.',
      });
    } catch (error) {
      return errorResponse('exporting file', error);
    }
  }

  private async handleBatchExport(args: BatchExportRequest): Promise<ToolResponse> {
    try {
      const mimeTypeMap = {
        pdf: 'application/pdf',
        docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      };

      const mimeType = mimeTypeMap[args.format];
      const results = {
        success: [] as { fileId: string; size: number }[],
        failed: [] as { fileId: string; error: string }[],
      };

      for (const fileId of args.fileIds) {
        try {
          const response = await this.drive.files.export(
            { fileId, mimeType },
            { responseType: 'arraybuffer' },
          );
          const buffer = Buffer.from(response.data as ArrayBuffer);
          results.success.push({ fileId, size: buffer.length });
        } catch (error) {
          results.failed.push({
            fileId,
            error: error instanceof Error ? error.message : String(error),
          });
        }
      }

      return successResponse({
        success: true,
        format: args.format,
        mimeType,
        exportedCount: results.success.length,
        failedCount: results.failed.length,
        totalSize: results.success.reduce((sum, r) => sum + r.size, 0),
        details: results,
      });
    } catch (error) {
      return errorResponse('in batch export', error);
    }
  }

  private async handleListFolderTree(args: ListFolderTreeRequest): Promise<ToolResponse> {
    try {
      const listFilesRecursive = async (folderId: string, path: string = ''): Promise<any[]> => {
        const query = `'${folderId}' in parents and trashed=false`;
        const fields = args.includeMetadata
          ? 'files(id, name, mimeType, size, createdTime, modifiedTime, owners, webViewLink)'
          : 'files(id, name, mimeType)';

        const response = await this.drive.files.list({ q: query, fields, pageSize: 1000 });
        const files = response.data.files || [];
        const results: any[] = [];

        for (const file of files) {
          if (!file.id || !file.name) continue;
          const filePath = path ? `${path}/${file.name}` : file.name;
          const fileInfo: any = { id: file.id, name: file.name, path: filePath, type: file.mimeType };

          if (args.includeMetadata) {
            fileInfo.size = file.size;
            fileInfo.created = file.createdTime;
            fileInfo.modified = file.modifiedTime;
            fileInfo.owner = file.owners?.[0]?.emailAddress;
            fileInfo.url = file.webViewLink;
          }

          results.push(fileInfo);

          if (args.recursive && file.mimeType === 'application/vnd.google-apps.folder') {
            const subFiles = await listFilesRecursive(file.id, filePath);
            results.push(...subFiles);
          }
        }
        return results;
      };

      const allFiles = await listFilesRecursive(args.folderId);

      return successResponse({
        success: true,
        folderId: args.folderId,
        totalFiles: allFiles.length,
        recursive: args.recursive,
        includeMetadata: args.includeMetadata,
        files: allFiles,
      });
    } catch (error) {
      return errorResponse('listing folder tree', error);
    }
  }

  private async handleListPermissions(args: ListPermissionsRequest): Promise<ToolResponse> {
    try {
      const response = await this.drive.permissions.list({
        fileId: args.fileId,
        fields:
          'permissions(id, type, role, emailAddress, domain, displayName, expirationTime, deleted)',
      });

      const permissions = response.data.permissions || [];
      return successResponse({
        success: true,
        fileId: args.fileId,
        permissionCount: permissions.length,
        permissions: permissions.map((p) => ({
          id: p.id,
          type: p.type,
          role: p.role,
          email: p.emailAddress,
          domain: p.domain,
          displayName: p.displayName,
          expirationTime: p.expirationTime,
          deleted: p.deleted,
        })),
      });
    } catch (error) {
      return errorResponse('listing permissions', error);
    }
  }

  private async handleListComments(args: ListCommentsRequest): Promise<ToolResponse> {
    try {
      const response = await this.drive.comments.list({
        fileId: args.fileId,
        includeDeleted: args.includeDeleted,
        pageSize: args.pageSize,
        pageToken: args.pageToken,
        fields: 'nextPageToken, comments(id, content, author, createdTime, modifiedTime, resolved, deleted, quotedFileContent, replies(id, content, author, createdTime, action))',
      });

      const comments = response.data.comments || [];

      // Filter out resolved if not requested
      const filtered = args.includeResolved
        ? comments
        : comments.filter((c) => !c.resolved);

      const result = filtered.map((c) => ({
        id: c.id,
        content: c.content,
        author: c.author?.displayName || c.author?.emailAddress,
        createdTime: c.createdTime,
        modifiedTime: c.modifiedTime,
        resolved: c.resolved,
        deleted: c.deleted,
        quotedContent: c.quotedFileContent?.value,
        replies: c.replies?.map((r) => ({
          id: r.id,
          content: r.content,
          author: r.author?.displayName || r.author?.emailAddress,
          createdTime: r.createdTime,
          action: r.action,
        })),
      }));

      let output: Record<string, unknown> = {
        success: true,
        fileId: args.fileId,
        commentCount: result.length,
        comments: result,
      };
      if (response.data.nextPageToken) {
        output.nextPageToken = response.data.nextPageToken;
      }
      return successResponse(output);
    } catch (error) {
      return errorResponse('listing comments', error);
    }
  }

  private async handleCreateComment(args: CreateCommentRequest): Promise<ToolResponse> {
    try {
      const response = await this.drive.comments.create({
        fileId: args.fileId,
        fields: 'id, content, author, createdTime',
        requestBody: {
          content: args.content,
        },
      });

      return successResponse({
        success: true,
        fileId: args.fileId,
        commentId: response.data.id,
        content: response.data.content,
        author: response.data.author?.displayName || response.data.author?.emailAddress,
        createdTime: response.data.createdTime,
      });
    } catch (error) {
      return errorResponse('creating comment', error);
    }
  }

  private async handleReplyToComment(args: ReplyToCommentRequest): Promise<ToolResponse> {
    try {
      const requestBody: any = { content: args.content };
      if (args.action) {
        requestBody.action = args.action;
      }

      const response = await this.drive.replies.create({
        fileId: args.fileId,
        commentId: args.commentId,
        fields: 'id, content, author, createdTime, action',
        requestBody,
      });

      return successResponse({
        success: true,
        fileId: args.fileId,
        commentId: args.commentId,
        replyId: response.data.id,
        content: response.data.content,
        author: response.data.author?.displayName || response.data.author?.emailAddress,
        action: response.data.action,
        createdTime: response.data.createdTime,
      });
    } catch (error) {
      return errorResponse('replying to comment', error);
    }
  }

  private async handleDeleteComment(args: DeleteCommentRequest): Promise<ToolResponse> {
    try {
      await this.drive.comments.delete({
        fileId: args.fileId,
        commentId: args.commentId,
      });

      return successResponse({
        success: true,
        fileId: args.fileId,
        commentId: args.commentId,
        message: 'Comment deleted',
      });
    } catch (error) {
      return errorResponse('deleting comment', error);
    }
  }
}
