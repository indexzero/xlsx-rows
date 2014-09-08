var xlsx = require('xlsx'),
    range = require('r...e');

//
// Regular expression to get the cell header
//
var letters    = range('A', 'Z').toArray(),
    cellHeader = /^([A-Z]+)\d{1,}/,
    isCell     = /^[A-Z]+\d{1,}/;

//
// ### function xlsxRows (options)
// #### @options {string|Object} Options for reading rows
// ####   file      {string} Name of the file to read.
// ####   sheetname {string} Name of the file to read.
// ####   format    {string} Format of the file to read.
// Reads the rows from the XLSX.
//
module.exports = function (options) {
  var file   = options,
      rows   = [],
      row    = [],
      start  = /^A\d{1,}/,
      format = 'w',
      workbook,
      sheetname,
      sheet;

  if (typeof options !== 'string') {
    file      = options.file;
    sheetname = options.sheetname;
    format    = options.format || 'w';
  }

  workbook  = xlsx.readFile(file);
  sheetname = sheetname || workbook.SheetNames[0];
  sheet     = workbook.Sheets[sheetname];

  if (!sheet) {
    throw new Error('No sheet with name: ' + sheetname);
  }

  //
  // Pushes the next row onto the `rows`
  //
  function pushRow() {
    //
    // Fill the row since we prefer the empty string
    // to a value of undefined.
    //
    row = row.map(function (val) {
      return val == undefined ? '' : val;
    });

    rows.push(row.slice());
    row = [];
  }

  Object.keys(sheet).forEach(function (cell) {
    if (!isCell.test(cell)) {
      return;
    }

    //
    // If we are the first cell (i.e. it is A12 or A0)
    // then add the current "row" to the "rows" ONLY
    // if it is not empty.
    //
    if (start.test(cell) && row && row.length) {
      pushRow();
    }

    var index = rowIndex(cell);
    row[index] = sheet[cell][format];
  });

  pushRow();
  return rows;
};

//
// ### function rowIndex (cell)
// Returns the row index for the cell
//
function rowIndex (cell) {
  var header = cellHeader.exec(cell),
      length;

  if (!header) {
    throw new Error('Bad cell header for: ' + cell);
  }

  //
  // TODO: Actually do something with the length
  // to support multi-character headers.
  //
  header = header[1];
  length = header.length;
  return letters.indexOf(header);
}