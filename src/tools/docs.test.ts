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
