import { describe, it, expect, vi } from 'vitest';
import { DocsHandler } from './docs.js';
import type { docs_v1, drive_v3 } from 'googleapis';

/** Build a minimal mock Docs + Drive client pair and return the handler */
function makeHandler(docData: docs_v1.Schema$Document) {
  const mockDocs = {
    documents: {
      get: vi.fn().mockResolvedValue({ data: docData }),
      create: vi.fn(),
      batchUpdate: vi.fn(),
    },
  } as unknown as docs_v1.Docs;

  const mockDrive = {} as unknown as drive_v3.Drive;
  return new DocsHandler(mockDocs, mockDrive);
}

/** Shorthand to build a text run structural element */
function textParagraph(text: string, extras?: Partial<docs_v1.Schema$TextRun>): docs_v1.Schema$StructuralElement {
  return {
    paragraph: {
      elements: [{ textRun: { content: text, ...extras } }],
    },
  };
}

describe('gdocs_get_document — table rendering', () => {
  it('renders a simple 2x2 table as pipe-delimited markdown', async () => {
    const handler = makeHandler({
      title: 'Test Doc',
      body: {
        content: [
          textParagraph('Before table\n'),
          {
            table: {
              rows: 2,
              columns: 2,
              tableRows: [
                {
                  tableCells: [
                    { content: [textParagraph('Header A\n')] },
                    { content: [textParagraph('Header B\n')] },
                  ],
                },
                {
                  tableCells: [
                    { content: [textParagraph('Value 1\n')] },
                    { content: [textParagraph('Value 2\n')] },
                  ],
                },
              ],
            },
          },
          textParagraph('After table\n'),
        ],
      },
    });

    const result = await handler.handleTool('gdocs_get_document', { documentId: 'test-id' });
    expect(result).not.toBeNull();

    const text = (result as any).content[0].text;
    expect(text).toContain('| Header A | Header B |');
    expect(text).toContain('| --- | --- |');
    expect(text).toContain('| Value 1 | Value 2 |');
    // Ensure surrounding text is preserved
    expect(text).toContain('Before table');
    expect(text).toContain('After table');
  });

  it('renders a single-row table (header only)', async () => {
    const handler = makeHandler({
      title: 'Single Row',
      body: {
        content: [
          {
            table: {
              rows: 1,
              columns: 2,
              tableRows: [
                {
                  tableCells: [
                    { content: [textParagraph('Col A\n')] },
                    { content: [textParagraph('Col B\n')] },
                  ],
                },
              ],
            },
          },
        ],
      },
    });

    const result = await handler.handleTool('gdocs_get_document', { documentId: 'test-id' });
    const text = (result as any).content[0].text;
    expect(text).toContain('| Col A | Col B |');
    expect(text).toContain('| --- | --- |');
  });

  it('handles multi-paragraph cells (collapses newlines)', async () => {
    const handler = makeHandler({
      title: 'Multi Para',
      body: {
        content: [
          {
            table: {
              rows: 1,
              columns: 1,
              tableRows: [
                {
                  tableCells: [
                    {
                      content: [
                        textParagraph('Line one\n'),
                        textParagraph('Line two\n'),
                      ],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    });

    const result = await handler.handleTool('gdocs_get_document', { documentId: 'test-id' });
    const text = (result as any).content[0].text;
    // Cell text should be collapsed to single line
    expect(text).toContain('| Line one Line two |');
  });

  it('captures suggestions inside table cells', async () => {
    const handler = makeHandler({
      title: 'Suggestions in Table',
      body: {
        content: [
          {
            table: {
              rows: 1,
              columns: 1,
              tableRows: [
                {
                  tableCells: [
                    {
                      content: [
                        {
                          paragraph: {
                            elements: [
                              {
                                textRun: {
                                  content: 'old text',
                                  suggestedDeletionIds: ['sug-1'],
                                },
                              },
                              {
                                textRun: {
                                  content: 'new text',
                                  suggestedInsertionIds: ['sug-1'],
                                },
                              },
                            ],
                          },
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    });

    const result = await handler.handleTool('gdocs_get_document', { documentId: 'test-id' });
    const text = (result as any).content[0].text;
    // Should contain the table cell text
    expect(text).toContain('old text');
    expect(text).toContain('new text');
    // Should report suggestions
    expect(text).toContain('Suggested Changes');
    expect(text).toContain('DELETE: "old text"');
    expect(text).toContain('INSERT: "new text"');
  });

  it('handles nested tables', async () => {
    const handler = makeHandler({
      title: 'Nested Table',
      body: {
        content: [
          {
            table: {
              rows: 1,
              columns: 1,
              tableRows: [
                {
                  tableCells: [
                    {
                      content: [
                        {
                          table: {
                            rows: 1,
                            columns: 1,
                            tableRows: [
                              {
                                tableCells: [
                                  { content: [textParagraph('Nested cell\n')] },
                                ],
                              },
                            ],
                          },
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    });

    const result = await handler.handleTool('gdocs_get_document', { documentId: 'test-id' });
    const text = (result as any).content[0].text;
    expect(text).toContain('Nested cell');
  });

  it('handles empty table gracefully', async () => {
    const handler = makeHandler({
      title: 'Empty Table',
      body: {
        content: [
          {
            table: {
              rows: 0,
              columns: 0,
              tableRows: [],
            },
          },
        ],
      },
    });

    const result = await handler.handleTool('gdocs_get_document', { documentId: 'test-id' });
    expect(result).not.toBeNull();
    // Should not crash — just produce no table output
  });

  it('preserves paragraph-only documents unchanged', async () => {
    const handler = makeHandler({
      title: 'No Tables',
      body: {
        content: [
          textParagraph('Hello world\n'),
          textParagraph('Second paragraph\n'),
        ],
      },
    });

    const result = await handler.handleTool('gdocs_get_document', { documentId: 'test-id' });
    const text = (result as any).content[0].text;
    expect(text).toContain('Hello world');
    expect(text).toContain('Second paragraph');
    expect(text).not.toContain('|');
  });
});

// ── gdocs_fill_table_cell ───────────────────────────────────────────────────

/**
 * Build a handler with a table at a known startIndex. Each cell has a single
 * paragraph whose start/end indices are explicit. Returns the handler plus the
 * batchUpdate mock so tests can inspect what requests were sent.
 *
 * Layout for a 2×2 table at startIndex=100 with cells containing "A\n", "B\n", "C\n", "D\n":
 *   Cell [0,0]: paragraph at [102, 104] (content "A\n", trailing \n at 103)
 *   Cell [0,1]: paragraph at [105, 107]
 *   Cell [1,0]: paragraph at [108, 110]
 *   Cell [1,1]: paragraph at [111, 113]
 * (Exact spacing isn't important for the tests — only that the start/end indices are present.)
 */
function makeHandlerWithTable(opts: {
  tableStartIndex: number;
  cellContent: { startIndex: number; endIndex: number; text: string }[][];
}) {
  const tableRows = opts.cellContent.map((row) => ({
    tableCells: row.map((cell) => ({
      content: [
        {
          startIndex: cell.startIndex,
          endIndex: cell.endIndex,
          paragraph: {
            elements: [
              {
                startIndex: cell.startIndex,
                endIndex: cell.endIndex,
                textRun: { content: cell.text },
              },
            ],
          },
        },
      ],
    })),
  }));

  const docData: docs_v1.Schema$Document = {
    title: 'Table Test',
    body: {
      content: [
        {
          startIndex: opts.tableStartIndex,
          endIndex: opts.tableStartIndex + 50,
          table: {
            rows: opts.cellContent.length,
            columns: opts.cellContent[0]?.length || 0,
            tableRows,
          },
        },
      ],
    },
  };

  const batchUpdate = vi.fn().mockResolvedValue({ data: {} });
  const mockDocs = {
    documents: {
      get: vi.fn().mockResolvedValue({ data: docData }),
      create: vi.fn(),
      batchUpdate,
    },
  } as unknown as docs_v1.Docs;

  const mockDrive = {} as unknown as drive_v3.Drive;
  return { handler: new DocsHandler(mockDocs, mockDrive), batchUpdate };
}

describe('gdocs_fill_table_cell', () => {
  it('replaces existing content: deletes [firstStart, lastEnd-1) then inserts', async () => {
    // Cell at [102, 108] holding "Hello\n": delete should be [102, 107), insert at 102.
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [
        [{ startIndex: 102, endIndex: 108, text: 'Hello\n' }],
      ],
    });

    const result = await handler.handleTool('gdocs_fill_table_cell', {
      documentId: 'test-id',
      tableStartIndex: 100,
      rowIndex: 0,
      columnIndex: 0,
      text: 'Replaced',
    });
    expect(result).not.toBeNull();
    expect((result as any).isError).toBeFalsy();

    expect(batchUpdate).toHaveBeenCalledTimes(1);
    const requests = batchUpdate.mock.calls[0][0].requestBody.requests;
    expect(requests).toHaveLength(2);
    expect(requests[0].deleteContentRange.range).toEqual({ startIndex: 102, endIndex: 107 });
    expect(requests[1].insertText).toEqual({ location: { index: 102 }, text: 'Replaced' });
  });

  it('replace on empty cell (just \\n): skips delete, only inserts', async () => {
    // Cell at [102, 103] holding "\n" only — nothing to delete.
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [
        [{ startIndex: 102, endIndex: 103, text: '\n' }],
      ],
    });

    await handler.handleTool('gdocs_fill_table_cell', {
      documentId: 'test-id',
      tableStartIndex: 100,
      rowIndex: 0,
      columnIndex: 0,
      text: 'FirstValue',
    });

    const requests = batchUpdate.mock.calls[0][0].requestBody.requests;
    expect(requests).toHaveLength(1);
    expect(requests[0].insertText).toEqual({ location: { index: 102 }, text: 'FirstValue' });
  });

  it('append mode: inserts at lastEnd-1 (before trailing newline)', async () => {
    // Cell at [102, 108] holding "Hello\n" — append should insert at 107.
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [
        [{ startIndex: 102, endIndex: 108, text: 'Hello\n' }],
      ],
    });

    await handler.handleTool('gdocs_fill_table_cell', {
      documentId: 'test-id',
      tableStartIndex: 100,
      rowIndex: 0,
      columnIndex: 0,
      text: ' world',
      mode: 'append',
    });

    const requests = batchUpdate.mock.calls[0][0].requestBody.requests;
    expect(requests).toHaveLength(1);
    expect(requests[0].insertText).toEqual({ location: { index: 107 }, text: ' world' });
  });

  it('addresses the correct cell in a multi-row, multi-column table', async () => {
    // Verify rowIndex/columnIndex routing: target cell [1,1] in a 2×2.
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [
        [
          { startIndex: 102, endIndex: 108, text: 'A0,0\n' },
          { startIndex: 110, endIndex: 116, text: 'A0,1\n' },
        ],
        [
          { startIndex: 118, endIndex: 124, text: 'A1,0\n' },
          { startIndex: 126, endIndex: 132, text: 'A1,1\n' },
        ],
      ],
    });

    await handler.handleTool('gdocs_fill_table_cell', {
      documentId: 'test-id',
      tableStartIndex: 100,
      rowIndex: 1,
      columnIndex: 1,
      text: 'NEW',
    });

    const requests = batchUpdate.mock.calls[0][0].requestBody.requests;
    expect(requests[0].deleteContentRange.range).toEqual({ startIndex: 126, endIndex: 131 });
    expect(requests[1].insertText).toEqual({ location: { index: 126 }, text: 'NEW' });
  });

  it('errors when rowIndex is out of bounds', async () => {
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [[{ startIndex: 102, endIndex: 108, text: 'A\n' }]],
    });

    const result = await handler.handleTool('gdocs_fill_table_cell', {
      documentId: 'test-id',
      tableStartIndex: 100,
      rowIndex: 5,
      columnIndex: 0,
      text: 'x',
    });
    expect((result as any).isError).toBe(true);
    expect((result as any).content[0].text).toContain('rowIndex 5 out of bounds');
    expect(batchUpdate).not.toHaveBeenCalled();
  });

  it('errors when no table exists at tableStartIndex', async () => {
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [[{ startIndex: 102, endIndex: 108, text: 'A\n' }]],
    });

    const result = await handler.handleTool('gdocs_fill_table_cell', {
      documentId: 'test-id',
      tableStartIndex: 999, // wrong index
      rowIndex: 0,
      columnIndex: 0,
      text: 'x',
    });
    expect((result as any).isError).toBe(true);
    expect((result as any).content[0].text).toContain('No table found');
    expect(batchUpdate).not.toHaveBeenCalled();
  });

  it('empty text is a no-op (returns success but no API call)', async () => {
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [[{ startIndex: 102, endIndex: 103, text: '\n' }]],
    });

    const result = await handler.handleTool('gdocs_fill_table_cell', {
      documentId: 'test-id',
      tableStartIndex: 100,
      rowIndex: 0,
      columnIndex: 0,
      text: '',
      mode: 'append',
    });
    expect((result as any).isError).toBeFalsy();
    expect(batchUpdate).not.toHaveBeenCalled();
  });
});

// ── gdocs_insert_table_row / gdocs_delete_table_row ─────────────────────────

describe('gdocs_insert_table_row', () => {
  it('sends insertTableRow with insertBelow=true by default', async () => {
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [
        [{ startIndex: 102, endIndex: 108, text: 'A\n' }],
        [{ startIndex: 110, endIndex: 116, text: 'B\n' }],
      ],
    });

    await handler.handleTool('gdocs_insert_table_row', {
      documentId: 'test-id',
      tableStartIndex: 100,
      rowIndex: 1,
    });

    expect(batchUpdate).toHaveBeenCalledTimes(1);
    const req = batchUpdate.mock.calls[0][0].requestBody.requests[0];
    expect(req.insertTableRow.tableCellLocation).toEqual({
      tableStartLocation: { index: 100 },
      rowIndex: 1,
      columnIndex: 0,
    });
    expect(req.insertTableRow.insertBelow).toBe(true);
  });

  it('honours insertBelow=false', async () => {
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [[{ startIndex: 102, endIndex: 108, text: 'A\n' }]],
    });

    await handler.handleTool('gdocs_insert_table_row', {
      documentId: 'test-id',
      tableStartIndex: 100,
      rowIndex: 0,
      insertBelow: false,
    });

    const req = batchUpdate.mock.calls[0][0].requestBody.requests[0];
    expect(req.insertTableRow.insertBelow).toBe(false);
  });

  it('errors when rowIndex is out of bounds', async () => {
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [[{ startIndex: 102, endIndex: 108, text: 'A\n' }]],
    });

    const result = await handler.handleTool('gdocs_insert_table_row', {
      documentId: 'test-id',
      tableStartIndex: 100,
      rowIndex: 5,
    });
    expect((result as any).isError).toBe(true);
    expect(batchUpdate).not.toHaveBeenCalled();
  });
});

describe('gdocs_delete_table_row', () => {
  it('sends deleteTableRow with the right tableCellLocation', async () => {
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [
        [{ startIndex: 102, endIndex: 108, text: 'A\n' }],
        [{ startIndex: 110, endIndex: 116, text: 'B\n' }],
      ],
    });

    await handler.handleTool('gdocs_delete_table_row', {
      documentId: 'test-id',
      tableStartIndex: 100,
      rowIndex: 0,
    });

    const req = batchUpdate.mock.calls[0][0].requestBody.requests[0];
    expect(req.deleteTableRow.tableCellLocation).toEqual({
      tableStartLocation: { index: 100 },
      rowIndex: 0,
      columnIndex: 0,
    });
  });

  it('refuses to delete the last remaining row', async () => {
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [[{ startIndex: 102, endIndex: 108, text: 'Only\n' }]],
    });

    const result = await handler.handleTool('gdocs_delete_table_row', {
      documentId: 'test-id',
      tableStartIndex: 100,
      rowIndex: 0,
    });
    expect((result as any).isError).toBe(true);
    expect((result as any).content[0].text).toContain('last row');
    expect(batchUpdate).not.toHaveBeenCalled();
  });
});

describe('gdocs_insert_table_column', () => {
  it('sends insertTableColumn with insertRight=true by default', async () => {
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [
        [
          { startIndex: 102, endIndex: 108, text: 'A\n' },
          { startIndex: 110, endIndex: 116, text: 'B\n' },
        ],
      ],
    });

    await handler.handleTool('gdocs_insert_table_column', {
      documentId: 'test-id',
      tableStartIndex: 100,
      columnIndex: 1,
    });

    const req = batchUpdate.mock.calls[0][0].requestBody.requests[0];
    expect(req.insertTableColumn.tableCellLocation).toEqual({
      tableStartLocation: { index: 100 },
      rowIndex: 0,
      columnIndex: 1,
    });
    expect(req.insertTableColumn.insertRight).toBe(true);
  });

  it('honours insertRight=false', async () => {
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [
        [
          { startIndex: 102, endIndex: 108, text: 'A\n' },
          { startIndex: 110, endIndex: 116, text: 'B\n' },
        ],
      ],
    });

    await handler.handleTool('gdocs_insert_table_column', {
      documentId: 'test-id',
      tableStartIndex: 100,
      columnIndex: 0,
      insertRight: false,
    });

    const req = batchUpdate.mock.calls[0][0].requestBody.requests[0];
    expect(req.insertTableColumn.insertRight).toBe(false);
  });
});

describe('gdocs_delete_table_column', () => {
  it('sends deleteTableColumn with the right tableCellLocation', async () => {
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [
        [
          { startIndex: 102, endIndex: 108, text: 'A\n' },
          { startIndex: 110, endIndex: 116, text: 'B\n' },
        ],
      ],
    });

    await handler.handleTool('gdocs_delete_table_column', {
      documentId: 'test-id',
      tableStartIndex: 100,
      columnIndex: 1,
    });

    const req = batchUpdate.mock.calls[0][0].requestBody.requests[0];
    expect(req.deleteTableColumn.tableCellLocation).toEqual({
      tableStartLocation: { index: 100 },
      rowIndex: 0,
      columnIndex: 1,
    });
  });

  it('refuses to delete the last remaining column', async () => {
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [[{ startIndex: 102, endIndex: 108, text: 'Only\n' }]],
    });

    const result = await handler.handleTool('gdocs_delete_table_column', {
      documentId: 'test-id',
      tableStartIndex: 100,
      columnIndex: 0,
    });
    expect((result as any).isError).toBe(true);
    expect((result as any).content[0].text).toContain('last column');
    expect(batchUpdate).not.toHaveBeenCalled();
  });
});

describe('gdocs_delete_table', () => {
  it('deletes the table\'s full structural element range', async () => {
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [[{ startIndex: 102, endIndex: 108, text: 'A\n' }]],
    });

    await handler.handleTool('gdocs_delete_table', {
      documentId: 'test-id',
      tableStartIndex: 100,
    });

    const req = batchUpdate.mock.calls[0][0].requestBody.requests[0];
    expect(req.deleteContentRange.range.startIndex).toBe(100);
    // Mock table sets element.endIndex = startIndex + 50
    expect(req.deleteContentRange.range.endIndex).toBe(150);
  });

  it('errors when no table exists at tableStartIndex', async () => {
    const { handler, batchUpdate } = makeHandlerWithTable({
      tableStartIndex: 100,
      cellContent: [[{ startIndex: 102, endIndex: 108, text: 'A\n' }]],
    });

    const result = await handler.handleTool('gdocs_delete_table', {
      documentId: 'test-id',
      tableStartIndex: 999,
    });
    expect((result as any).isError).toBe(true);
    expect((result as any).content[0].text).toContain('No table found');
    expect(batchUpdate).not.toHaveBeenCalled();
  });
});

// ── gdocs_replace_text — suggestion-collision hint ─────────────────────────

describe('gdocs_replace_text — suggestion-collision hint', () => {
  it('appends a hint about pending suggestions when batchUpdate fails', async () => {
    // Mock: batchUpdate rejects with a 500 (like the real Docs API does on suggestion collisions).
    // documents.get returns a doc with one suggested-deletion ID on a textRun.
    const docWithSuggestion: docs_v1.Schema$Document = {
      title: 'Doc with pending suggestion',
      body: {
        content: [
          {
            paragraph: {
              elements: [
                { textRun: { content: 'Hello ', suggestedDeletionIds: ['sugg-1'] } },
                { textRun: { content: 'world\n' } },
              ],
            },
          },
        ],
      },
    };

    const mockDocs = {
      documents: {
        get: vi.fn().mockResolvedValue({ data: docWithSuggestion }),
        create: vi.fn(),
        batchUpdate: vi.fn().mockRejectedValue(
          Object.assign(new Error('Internal error encountered.'), {
            response: {
              status: 500,
              data: { error: { code: 500, status: 'INTERNAL', message: 'Internal error encountered.' } },
            },
          }),
        ),
      },
    } as unknown as docs_v1.Docs;
    const mockDrive = {} as unknown as drive_v3.Drive;
    const handler = new DocsHandler(mockDocs, mockDrive);

    const result = await handler.handleTool('gdocs_replace_text', {
      documentId: 'test-id',
      find: 'Hello',
      replace: 'Hi',
    });

    expect((result as any).isError).toBe(true);
    const text = (result as any).content[0].text;
    expect(text).toContain('[HTTP 500]');
    expect(text).toContain('[INTERNAL]');
    expect(text).toContain('Internal error encountered.');
    expect(text).toContain('pending suggestion');
    expect(text).toContain('1 pending suggestion');
  });

  it('does not append a hint when the document has no pending suggestions', async () => {
    const cleanDoc: docs_v1.Schema$Document = {
      title: 'Clean doc',
      body: { content: [textParagraph('No suggestions here\n')] },
    };

    const mockDocs = {
      documents: {
        get: vi.fn().mockResolvedValue({ data: cleanDoc }),
        create: vi.fn(),
        batchUpdate: vi.fn().mockRejectedValue(
          Object.assign(new Error('Bad request'), {
            response: { status: 400, data: { error: { code: 400, status: 'INVALID_ARGUMENT', message: 'Bad request' } } },
          }),
        ),
      },
    } as unknown as docs_v1.Docs;
    const mockDrive = {} as unknown as drive_v3.Drive;
    const handler = new DocsHandler(mockDocs, mockDrive);

    const result = await handler.handleTool('gdocs_replace_text', {
      documentId: 'test-id',
      find: 'foo',
      replace: 'bar',
    });

    expect((result as any).isError).toBe(true);
    const text = (result as any).content[0].text;
    expect(text).toContain('[HTTP 400]');
    expect(text).not.toContain('pending suggestion');
  });
});
