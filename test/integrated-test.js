var assert = require('assert');
var XLSX = require('xlsx');
var XLSX_CALC = require("../");

describe('XLSX with XLSX_CALC', function() {
    function assert_values(sheet_calculated, sheet_expected) {
        for (var prop in sheet_expected) {
            if(prop.match(/[A-Z]+[0-9]+/)) {
                assert.equal(sheet_expected[prop].v, sheet_calculated[prop].v, "Error: " + prop + ' f="' + sheet_expected[prop].f +'"');
                assert.equal(sheet_expected[prop].w, sheet_calculated[prop].w, "Error: " + prop + ' f="' + sheet_expected[prop].f +'"');
                assert.equal(sheet_expected[prop].t, sheet_calculated[prop].t, "Error: " + prop + ' f="' + sheet_expected[prop].f +'"');
            }
        }
    }
    function read_workbook() {
        return XLSX.readFile('test/testcase.xlsx');
    }
    function read_sheet() {
        return read_workbook().Sheets.Sheet1;
    }
    function erase_values_that_contains_formula(sheet) {
        for (var prop in sheet) {
            if(prop.match(/[A-Z]+[0-9]+/) && sheet[prop].f) {
                sheet[prop].v = null;
            }
        }
    }
    var workbook, original_sheet;
    beforeEach(function() {
        workbook = read_workbook();
        erase_values_that_contains_formula(workbook.Sheets.Sheet1);
        original_sheet = read_sheet();
    });
    it('recalc the workbook', function() {
        XLSX_CALC(workbook);
        assert_values(original_sheet, workbook.Sheets.Sheet1);
    });  
});
