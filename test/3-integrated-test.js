var assert = require('assert');
var XLSX = require('xlsx');
var XLSX_CALC = require("../lib/xlsx-calc");

describe('XLSX with XLSX_CALC', function() {

    function assert_values(sheet_expected, sheet_calculated) {
        for (var prop in sheet_expected) {
            if(prop.match(/[A-Z]+[0-9]+/)) {
                assert.equal(sheet_calculated[prop].v, sheet_expected[prop].v, "Error: " + prop + ' f="' + sheet_expected[prop].f +'"\nexpected ' + sheet_expected[prop].v + " got " + sheet_calculated[prop].v);
                assert.equal(sheet_calculated[prop].w, sheet_expected[prop].w, "Error: " + prop + ' f="' + sheet_expected[prop].f +'"');
                assert.equal(sheet_calculated[prop].t, sheet_expected[prop].t, "Error: " + prop + ' f="' + sheet_expected[prop].f +'"');
            }
        }
    }

    function erase_values_that_contains_formula(sheets) {
        for (var sheet in sheets) {
            for (var prop in sheets[sheet]) {
                if(prop.match(/[A-Z]+[0-9]+/) && sheets[sheet][prop].f) {
                    sheets[sheet][prop].v = null;
                }       
            }
        }
    }

    it('erase_values_that_contains_formula sets values to null', function () {
        var workbook = {
            Sheets: {
                Sheet1: {
                    A1: { v: 1, f: '1*1', t: 'n' }
                }
            }
        };
        erase_values_that_contains_formula(workbook.Sheets);
        assert.equal(workbook.Sheets.Sheet1.A1.v, null);
    });

    it('recalc the workbook Sheet1', function() {
        var workbook = XLSX.readFile('test/testcase.xlsx');
        erase_values_that_contains_formula(workbook.Sheets);
        var original_sheet = XLSX.readFile('test/testcase.xlsx').Sheets.Sheet1;
        XLSX_CALC(workbook);
        assert_values(original_sheet, workbook.Sheets.Sheet1);
    });
    
    it('recalc the workbook Sheet OffSet', function() {
        var workbook = XLSX.readFile('test/testcase.xlsx');
        erase_values_that_contains_formula(workbook.Sheets);
        var original_sheet = XLSX.readFile('test/testcase.xlsx').Sheets.OffSet;
        XLSX_CALC(workbook);
        assert_values(original_sheet, workbook.Sheets.OffSet);
    });

});
