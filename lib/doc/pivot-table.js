const {objectFromProps, range, toSortedArray} = require('../utils/utils');

// TK(2023-10-10): turn this into a class constructor.

function makePivotTable(worksheet, model) {
  // Example `model`:
  // {
  //   // Source of data: the entire sheet range is taken,
  //   // akin to `worksheet1.getSheetValues()`.
  //   sourceSheet: worksheet1,
  //
  //   // Pivot table fields: values indicate field names;
  //   // they come from the first row in `worksheet1`.
  //   rows: ['A', 'B'],
  //   columns: ['C'],
  //
  //   // Values can be specified in two ways:
  //   // 1. Array of strings with single metric (legacy):
  //   //    values: ['E', 'F'], metric: 'sum'
  //   // 2. Array of objects with individual metrics (new):
  //   //    values: [{field: 'E', type: 'sum'}, {field: 'F', type: 'average'}]
  //   values: ['E'],
  //   metric: 'sum', // applies to all values if values is array of strings
  // }

  validate(worksheet, model);

  const {sourceSheet} = model;
  let {rows, columns, values} = model;
  const defaultMetric = model.metric || 'sum';

  const cacheFields = makeCacheFields(sourceSheet, [...rows, ...columns]);

  // let {rows, columns, values} use indices instead of names;
  // names can then be accessed via `pivotTable.cacheFields[index].name`.
  // *Note*: Using `reduce` as `Object.fromEntries` requires Node 12+;
  // ExcelJS is >=8.3.0 (as of 2023-10-08).
  const nameToIndex = cacheFields.reduce((result, cacheField, index) => {
    result[cacheField.name] = index;
    return result;
  }, {});
  rows = rows.map(row => nameToIndex[row]);
  columns = columns.map(column => nameToIndex[column]);

  // Normalize values to always be array of objects with field index and type
  values = values.map(value => {
    if (typeof value === 'string') {
      // Legacy format: string field name with global metric
      return {
        field: nameToIndex[value],
        type: defaultMetric,
      };
    }
    // New format: object with field name and individual type
    return {
      field: nameToIndex[value.field],
      type: value.type || defaultMetric,
    };
  });

  // form pivot table object
  return {
    sourceSheet,
    rows,
    columns,
    values,
    cacheFields,
    // defined in <pivotTableDefinition> of xl/pivotTables/pivotTable1.xml;
    // also used in xl/workbook.xml
    cacheId: '10',
  };
}

function validate(worksheet, model) {
  if (worksheet.workbook.pivotTables.length === 1) {
    throw new Error(
      'A pivot table was already added. At this time, ExcelJS supports at most one pivot table per file.'
    );
  }

  // Valid aggregation types per OOXML ST_DataConsolidateFunction
  const validAggregations = new Set([
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
  ]);

  // Validate global metric if provided
  if (model.metric && !validAggregations.has(model.metric)) {
    throw new Error(
      `Invalid metric "${model.metric}". Valid aggregations are: ${Array.from(validAggregations).join(', ')}.`
    );
  }

  const headerNames = model.sourceSheet.getRow(1).values.slice(1);
  const isInHeaderNames = objectFromProps(headerNames, true);

  // Extract field names from values (support both string and object formats)
  const valueFieldNames = model.values.map(value => (typeof value === 'string' ? value : value.field));

  // Validate all field names exist in headers
  for (const name of [...model.rows, ...model.columns, ...valueFieldNames]) {
    if (!isInHeaderNames[name]) {
      throw new Error(`The header name "${name}" was not found in ${model.sourceSheet.name}.`);
    }
  }

  // Validate individual value aggregation types
  for (const value of model.values) {
    if (typeof value === 'object' && value.type && !validAggregations.has(value.type)) {
      throw new Error(
        `Invalid aggregation type "${value.type}" for field "${value.field}". Valid aggregations are: ${Array.from(
          validAggregations
        ).join(', ')}.`
      );
    }
  }

  if (!model.rows.length) {
    throw new Error('No pivot table rows specified.');
  }

  if (!model.columns.length) {
    throw new Error('No pivot table columns specified.');
  }

  if (!model.values.length) {
    throw new Error('At least 1 value field needs to be specified.');
  }
}

function makeCacheFields(worksheet, fieldNamesWithSharedItems) {
  // Cache fields are used in pivot tables to reference source data.
  //
  // Example
  // -------
  // Turn
  //
  //  `worksheet` sheet values [
  //    ['A', 'B', 'C', 'D', 'E'],
  //    ['a1', 'b1', 'c1', 4, 5],
  //    ['a1', 'b2', 'c1', 4, 5],
  //    ['a2', 'b1', 'c2', 14, 24],
  //    ['a2', 'b2', 'c2', 24, 35],
  //    ['a3', 'b1', 'c3', 34, 45],
  //    ['a3', 'b2', 'c3', 44, 45]
  //  ];
  //  fieldNamesWithSharedItems = ['A', 'B', 'C'];
  //
  // into
  //
  //  [
  //    { name: 'A', sharedItems: ['a1', 'a2', 'a3'] },
  //    { name: 'B', sharedItems: ['b1', 'b2'] },
  //    { name: 'C', sharedItems: ['c1', 'c2', 'c3'] },
  //    { name: 'D', sharedItems: null },
  //    { name: 'E', sharedItems: null }
  //  ]

  const names = worksheet.getRow(1).values;
  const nameToHasSharedItems = objectFromProps(fieldNamesWithSharedItems, true);

  const aggregate = columnIndex => {
    const columnValues = worksheet.getColumn(columnIndex).values.splice(2);
    const columnValuesAsSet = new Set(columnValues);
    return toSortedArray(columnValuesAsSet);
  };

  // make result
  const result = [];
  for (const columnIndex of range(1, names.length)) {
    const name = names[columnIndex];
    const sharedItems = nameToHasSharedItems[name] ? aggregate(columnIndex) : null;
    result.push({name, sharedItems});
  }
  return result;
}

module.exports = {makePivotTable};
