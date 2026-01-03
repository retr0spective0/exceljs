// --------------------------------------------------
// Demonstration of multiple value fields with different aggregations
// This showcases the new pivot table functionality added to ExcelJS
// --------------------------------------------------

const Excel = require('../lib/exceljs.nodejs');

async function main() {
  const workbook = new Excel.Workbook();

  // Create source data worksheet
  const dataSheet = workbook.addWorksheet('Sales Data');
  dataSheet.addRows([
    ['Product', 'Region', 'Quarter', 'Sales', 'Units', 'Cost'],
    ['Laptop', 'North', 'Q1', 50000, 100, 30000],
    ['Laptop', 'North', 'Q2', 55000, 110, 33000],
    ['Laptop', 'South', 'Q1', 48000, 96, 28800],
    ['Laptop', 'South', 'Q2', 52000, 104, 31200],
    ['Phone', 'North', 'Q1', 30000, 150, 18000],
    ['Phone', 'North', 'Q2', 35000, 175, 21000],
    ['Phone', 'South', 'Q1', 28000, 140, 16800],
    ['Phone', 'South', 'Q2', 32000, 160, 19200],
    ['Tablet', 'North', 'Q1', 20000, 80, 12000],
    ['Tablet', 'North', 'Q2', 22000, 88, 13200],
    ['Tablet', 'South', 'Q1', 18000, 72, 10800],
    ['Tablet', 'South', 'Q2', 21000, 84, 12600],
  ]);

  // Example 1: Multiple values with different aggregations
  const pivotSheet1 = workbook.addWorksheet('Pivot - Multi Aggregations');
  pivotSheet1.addPivotTable({
    sourceSheet: dataSheet,
    rows: ['Product'],
    columns: ['Region'],
    values: [
      {field: 'Sales', type: 'sum'},
      {field: 'Sales', type: 'average'},
      {field: 'Units', type: 'count'},
      {field: 'Cost', type: 'max'},
    ],
  });

  console.log(
    '✅ Created pivot table with multiple aggregations (sum, average, count, max)'
  );

  // Save the workbook
  const filename = './spec/out/pivot-table-demo.xlsx';
  await workbook.xlsx.writeFile(filename);

  console.log(`\n✅ Successfully created Excel file: ${filename}`);
  console.log('\nFeatures demonstrated:');
  console.log('  • Multiple value fields in a single pivot table');
  console.log('  • Different aggregation types (sum, average, count, max)');
  console.log('  • All 11 OOXML aggregation types supported:');
  console.log('    - average, count, countNums, max, min');
  console.log('    - product, stdDev, stdDevP, sum, var, varP');
  console.log('\nBackward compatibility:');
  console.log(
    '  • Legacy format still works: values: [\'Sales\', \'Units\'], metric: \'sum\''
  );
  console.log('  • New format: values: [{field: \'Sales\', type: \'sum\'}]');
  console.log('  • Mixed format supported\n');
}

main().catch(err => {
  console.error('Error:', err);
  throw err;
});
