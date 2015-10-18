var XLSX = require('xlsx');
var XLSX_CALC = require("../");

var workbook = XLSX.readFile('test/testcase.xlsx');

var first_sheet_name = workbook.SheetNames[0];

/* Get worksheet */
var worksheet = workbook.Sheets[first_sheet_name];

var assert = require('assert');

describe('XLSX with XLSX_CALC', function() {
    it('recalc the workbook', function() {
        console.log(worksheet);
        var expected = worksheet['A1'].v;
        worksheet['A1'].v = 0;
        XLSX_CALC(workbook);
        assert.equal(expected, worksheet['A1'].v);
    });  
});
