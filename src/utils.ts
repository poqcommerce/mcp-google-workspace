import { sheets_v4 } from 'googleapis';
import type { ToolResponse } from './types.js';

/**
 * Convert A1 notation (e.g. "A1:C3" or "Sheet1!A1:C3") to a GridRange object.
 * Handles multi-letter columns (AA, AB, etc).
 * Returns the sheet name separately so callers can resolve it to a sheetId if needed.
 */
export function a1ToGridRange(a1Range: string, sheetId: number = 0): sheets_v4.Schema$GridRange & { sheetName?: string } {
  // Strip optional sheet name prefix: "Sheet1!A1:C3" → sheetName="Sheet1", rangePart="A1:C3"
  let sheetName: string | undefined;
  let rangePart = a1Range;
  const prefixMatch = a1Range.match(/^(?:'([^']+)'|([^!]+))!(.+)$/);
  if (prefixMatch) {
    sheetName = prefixMatch[1] || prefixMatch[2];
    rangePart = prefixMatch[3];
  }

  const match = rangePart.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid A1 range: ${a1Range}`);
  }

  const [, startCol, startRow, endCol, endRow] = match;

  return {
    sheetId,
    startRowIndex: parseInt(startRow) - 1,
    endRowIndex: parseInt(endRow),
    startColumnIndex: columnToIndex(startCol),
    endColumnIndex: columnToIndex(endCol) + 1,
    sheetName,
  };
}

/**
 * Convert a column letter (A, B, ..., Z, AA, AB, ...) to a 0-based index.
 */
export function columnToIndex(column: string): number {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return result - 1;
}

/** Build a success response in the standard MCP format. */
export function successResponse(data: Record<string, any>): ToolResponse {
  return {
    content: [{ type: 'text', text: JSON.stringify(data, null, 2) }],
  };
}

/** Build a plain-text success response (non-JSON). */
export function textResponse(text: string): ToolResponse {
  return {
    content: [{ type: 'text', text }],
  };
}

/** Normalise escaped sequences (literal \n, \t, \uXXXX) into real characters. */
export function normaliseText(text: string): string {
  return text
    .replace(/\\n/g, '\n')
    .replace(/\\t/g, '\t')
    .replace(/\\u([0-9a-fA-F]{4})/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
}

/** Build an error response in the standard MCP format. */
export function errorResponse(context: string, error: unknown): ToolResponse {
  const message = error instanceof Error ? error.message : String(error);
  return {
    content: [{ type: 'text', text: `Error ${context}: ${message}` }],
    isError: true,
  };
}
