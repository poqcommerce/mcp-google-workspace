# Basic Usage Examples

## Creating a Simple Dashboard

```typescript
// Create a new sheet with KPI data
await mcp.call('gsheets_create_and_populate', {
  title: 'Monthly KPIs',
  data: [
    ['Metric', 'Value', 'Change'],
    ['Revenue', '$50,000', '+15%'],
    ['Users', '1,250', '+8%'],
    ['Conversion', '3.2%', '+0.3%']
  ]
});
Batch Updates
typescript// Update multiple sections at once
await mcp.call('gsheets_batch_update', {
  spreadsheetId: 'your-sheet-id',
  updates: [
    {
      range: 'A1:C1',
      values: [['Sales Report - Q4 2024']]
    },
    {
      range: 'A3:B10',
      values: [
        ['Product', 'Sales'],
        ['Widget A', '1000'],
        ['Widget B', '750'],
        // ... more data
      ]
    }
  ]
});
