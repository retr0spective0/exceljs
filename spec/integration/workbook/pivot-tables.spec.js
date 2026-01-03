// *Note*: `fs.promises` not supported before Node.js 11.14.0;
// ExcelJS version range '>=8.3.0' (as of 2023-10-08).
const fs = require('fs');
const {promisify} = require('util');

const fsReadFileAsync = promisify(fs.readFile);

const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const PIVOT_TABLE_FILEPATHS = [
  'xl/pivotCache/pivotCacheRecords1.xml',
  'xl/pivotCache/pivotCacheDefinition1.xml',
  'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
  'xl/pivotTables/pivotTable1.xml',
  'xl/pivotTables/_rels/pivotTable1.xml.rels',
];

const TEST_XLSX_FILEPATH = './spec/out/wb.test.xlsx';

const TEST_DATA = [
  ['A', 'B', 'C', 'D', 'E'],
  ['a1', 'b1', 'c1', 4, 5],
  ['a1', 'b2', 'c1', 4, 5],
  ['a2', 'b1', 'c2', 14, 24],
  ['a2', 'b2', 'c2', 24, 35],
  ['a3', 'b1', 'c3', 34, 45],
  ['a3', 'b2', 'c3', 44, 45],
];

// =============================================================================
// Tests

describe('Workbook', () => {
  describe('Pivot Tables', () => {
    it('if pivot table added, then certain xml and rels files are added', async () => {
      const workbook = new ExcelJS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      const worksheet2 = workbook.addWorksheet('Sheet2');
      worksheet2.addPivotTable({
        sourceSheet: worksheet1,
        rows: ['A', 'B'],
        columns: ['C'],
        values: ['E'],
        metric: 'sum',
      });

      return workbook.xlsx.writeFile(TEST_XLSX_FILEPATH).then(async () => {
        const buffer = await fsReadFileAsync(TEST_XLSX_FILEPATH);
        const zip = await JSZip.loadAsync(buffer);
        for (const filepath of PIVOT_TABLE_FILEPATHS) {
          expect(zip.files[filepath]).to.not.be.undefined();
        }
      });
    });

    it('if pivot table NOT added, then certain xml and rels files are not added', () => {
      const workbook = new ExcelJS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      workbook.addWorksheet('Sheet2');

      return workbook.xlsx.writeFile(TEST_XLSX_FILEPATH).then(async () => {
        const buffer = await fsReadFileAsync(TEST_XLSX_FILEPATH);
        const zip = await JSZip.loadAsync(buffer);
        for (const filepath of PIVOT_TABLE_FILEPATHS) {
          expect(zip.files[filepath]).to.be.undefined();
        }
      });
    });

    it('supports multiple value fields with same aggregation (legacy format)', async () => {
      const workbook = new ExcelJS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      const worksheet2 = workbook.addWorksheet('Sheet2');
      worksheet2.addPivotTable({
        sourceSheet: worksheet1,
        rows: ['A'],
        columns: ['B'],
        values: ['D', 'E'], // Multiple values with same metric
        metric: 'sum',
      });

      return workbook.xlsx.writeFile(TEST_XLSX_FILEPATH).then(async () => {
        const buffer = await fsReadFileAsync(TEST_XLSX_FILEPATH);
        const zip = await JSZip.loadAsync(buffer);

        // Verify pivot table files exist
        for (const filepath of PIVOT_TABLE_FILEPATHS) {
          expect(zip.files[filepath]).to.not.be.undefined();
        }

        // Verify dataFields contains both value fields
        const pivotTableXml = await zip.files[
          'xl/pivotTables/pivotTable1.xml'
        ].async('string');
        expect(pivotTableXml).to.include('dataFields count="2"');
        expect(pivotTableXml).to.include('Sum of D');
        expect(pivotTableXml).to.include('Sum of E');
      });
    });

    it('supports multiple value fields with different aggregations (new format)', async () => {
      const workbook = new ExcelJS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      const worksheet2 = workbook.addWorksheet('Sheet2');
      worksheet2.addPivotTable({
        sourceSheet: worksheet1,
        rows: ['A'],
        columns: ['B'],
        values: [
          {field: 'D', type: 'sum'},
          {field: 'E', type: 'average'},
        ],
      });

      return workbook.xlsx.writeFile(TEST_XLSX_FILEPATH).then(async () => {
        const buffer = await fsReadFileAsync(TEST_XLSX_FILEPATH);
        const zip = await JSZip.loadAsync(buffer);

        // Verify pivot table files exist
        for (const filepath of PIVOT_TABLE_FILEPATHS) {
          expect(zip.files[filepath]).to.not.be.undefined();
        }

        // Verify dataFields contains both value fields with different aggregations
        const pivotTableXml = await zip.files[
          'xl/pivotTables/pivotTable1.xml'
        ].async('string');
        expect(pivotTableXml).to.include('dataFields count="2"');
        expect(pivotTableXml).to.include('Sum of D');
        expect(pivotTableXml).to.include('Average of E');
        expect(pivotTableXml).to.include('subtotal="average"');
      });
    });

    it('supports all aggregation types', async () => {
      const workbook = new ExcelJS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      const worksheet2 = workbook.addWorksheet('Sheet2');

      // Test each aggregation type
      const aggregations = [
        'average',
        'count',
        'countNums',
        'max',
        'min',
        'product',
        'stdDev',
        'stdDevP',
        'sum',
        'var',
        'varP',
      ];

      for (const aggType of aggregations) {
        worksheet2.addPivotTable({
          sourceSheet: worksheet1,
          rows: ['A'],
          columns: ['B'],
          values: [{field: 'E', type: aggType}],
        });

        // Only one pivot table per workbook, so break after first
        break;
      }

      return workbook.xlsx.writeFile(TEST_XLSX_FILEPATH).then(async () => {
        const buffer = await fsReadFileAsync(TEST_XLSX_FILEPATH);
        const zip = await JSZip.loadAsync(buffer);

        const pivotTableXml = await zip.files[
          'xl/pivotTables/pivotTable1.xml'
        ].async('string');
        expect(pivotTableXml).to.include('Average of E');
      });
    });

    it('validates invalid aggregation types', () => {
      const workbook = new ExcelJS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      const worksheet2 = workbook.addWorksheet('Sheet2');

      expect(() => {
        worksheet2.addPivotTable({
          sourceSheet: worksheet1,
          rows: ['A'],
          columns: ['B'],
          values: [{field: 'E', type: 'invalid'}],
        });
      }).to.throw('Invalid aggregation type "invalid"');
    });

    it('validates at least one value field is required', () => {
      const workbook = new ExcelJS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      const worksheet2 = workbook.addWorksheet('Sheet2');

      expect(() => {
        worksheet2.addPivotTable({
          sourceSheet: worksheet1,
          rows: ['A'],
          columns: ['B'],
          values: [],
        });
      }).to.throw('At least 1 value field needs to be specified');
    });

    it('supports mixed format (some with type, some without)', async () => {
      const workbook = new ExcelJS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      const worksheet2 = workbook.addWorksheet('Sheet2');
      worksheet2.addPivotTable({
        sourceSheet: worksheet1,
        rows: ['A'],
        columns: ['B'],
        values: [
          'D', // Legacy string format - will use default metric
          {field: 'E', type: 'max'}, // New object format
        ],
        metric: 'sum', // Default for string values
      });

      return workbook.xlsx.writeFile(TEST_XLSX_FILEPATH).then(async () => {
        const buffer = await fsReadFileAsync(TEST_XLSX_FILEPATH);
        const zip = await JSZip.loadAsync(buffer);

        const pivotTableXml = await zip.files[
          'xl/pivotTables/pivotTable1.xml'
        ].async('string');
        expect(pivotTableXml).to.include('dataFields count="2"');
        expect(pivotTableXml).to.include('Sum of D');
        expect(pivotTableXml).to.include('Max of E');
      });
    });
  });
});
