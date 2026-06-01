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

/**
 * Build an error response in the standard MCP format.
 *
 * Surfaces diagnostic detail from googleapis errors (GaxiosError) where possible:
 * HTTP status code, the API's structured `status` (e.g. FAILED_PRECONDITION,
 * INVALID_ARGUMENT), and the per-field error details. Without this, all
 * Google API errors collapse to short messages like "Internal error encountered."
 * with no actionable diagnostic — particularly painful for batchUpdate failures
 * (e.g. replaceAllText colliding with pending suggested changes).
 */
export function errorResponse(context: string, error: unknown, suffix?: string): ToolResponse {
  let message: string;

  if (error instanceof Error) {
    message = error.message;
    const e = error as any;

    // googleapis throws GaxiosError where .response.data.error holds the structured API error
    const apiError = e.response?.data?.error;
    if (apiError && typeof apiError === 'object') {
      const parts: string[] = [];

      // HTTP code (often the fastest signal: 4xx = your input is wrong; 5xx = Google's issue)
      const httpCode = e.response?.status || apiError.code;
      if (httpCode) parts.push(`[HTTP ${httpCode}]`);

      // Google's textual status (FAILED_PRECONDITION, INVALID_ARGUMENT, etc.)
      if (apiError.status && apiError.status !== String(httpCode)) {
        parts.push(`[${apiError.status}]`);
      }

      if (apiError.message) parts.push(apiError.message);

      // Per-field details — usually the most actionable content.
      // Skip entries whose message exactly matches apiError.message (Google often duplicates
      // the same string across .message and .errors[].message) to avoid noisy redundant output.
      const details: string[] = [];
      const topMessage = apiError.message;
      if (Array.isArray(apiError.errors)) {
        for (const d of apiError.errors) {
          if (!d || typeof d !== 'object') continue;
          const isRedundantMessage = d.message && d.message === topMessage;
          const bits: string[] = [];
          if (d.message && !isRedundantMessage) bits.push(d.message);
          if (d.reason) bits.push(`(${d.reason})`);
          if (d.location) bits.push(`at ${d.location}`);
          if (bits.length > 0) details.push(bits.join(' '));
        }
      }
      if (Array.isArray(apiError.details)) {
        for (const d of apiError.details) {
          if (!d || typeof d !== 'object') continue;
          if (d.detail && d.detail !== topMessage) details.push(d.detail);
          else if (d.message && d.message !== topMessage) details.push(d.message);
          else if (d.fieldViolations && Array.isArray(d.fieldViolations)) {
            for (const fv of d.fieldViolations) {
              if (fv.description) details.push(`${fv.field || ''}: ${fv.description}`.trim());
            }
          }
        }
      }
      if (details.length > 0) {
        parts.push(`— ${details.join('; ')}`);
      }

      message = parts.join(' ');
    } else if (e.response?.status) {
      // No structured body but we have an HTTP code — surface that at least
      message = `[HTTP ${e.response.status}] ${message}`;
    }
  } else {
    message = String(error);
  }

  if (suffix) message += suffix;

  return {
    content: [{ type: 'text', text: `Error ${context}: ${message}` }],
    isError: true,
  };
}
