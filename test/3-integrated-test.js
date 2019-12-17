var assert = require('assert');
var XLSX = require('xlsx');
const XLSX_CALC = require("../src");

describe('XLSX with XLSX_CALC', function() {

    function assert_values(sheet_expected, sheet_calculated) {
        for (var prop in sheet_expected) {
            if(prop.match(/[A-Z]+[0-9]+/)) {
                assert.equal(sheet_calculated[prop].t, sheet_expected[prop].t, "Error: " + prop + ' f="' + sheet_expected[prop].f +'"');
                if (sheet_calculated[prop].t === 'e') return;
                if (typeof sheet_calculated[prop].v === 'number' && typeof sheet_expected[prop].v === 'number') {
                    assert.equal(sheet_calculated[prop].v.toFixed(10), sheet_expected[prop].v.toFixed(10), "Error: " + prop + ' f="' + sheet_expected[prop].f +'"\nexpected ' + sheet_expected[prop].v + " got " + sheet_calculated[prop].v);
                } else {
                    assert.equal(sheet_calculated[prop].v, sheet_expected[prop].v, "Error: " + prop + ' f="' + sheet_expected[prop].f +'"\nexpected ' + sheet_expected[prop].v + " got " + sheet_calculated[prop].v);
                }
                assert.equal(sheet_calculated[prop].w, sheet_expected[prop].w, "Error: " + prop + ' f="' + sheet_expected[prop].f +'"');
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

    it('handles the sheet name', function() {
        var workbook = XLSX.readFile('test/tias.xlsx');
        erase_values_that_contains_formula(workbook.Sheets);
        var original_sheet = XLSX.readFile('test/tias.xlsx').Sheets.Sheet2;
        XLSX_CALC(workbook);
        assert_values(original_sheet, workbook.Sheets.Sheet2);
    });

    it('handles transpose', function() {
        var workbook = XLSX.readFile('test/transpose.xlsx');
        var sheet1 = workbook.Sheets.Sheet1;
        //console.log(workbook.Sheets.Sheet1);
        sheet1.G13.v = null;
        XLSX_CALC(workbook);
        assert.equal(sheet1.F13.v, 1);
        assert.equal(sheet1.G13.v, 4);
        assert.equal(sheet1.H13.v, 7);

        assert.equal(sheet1.F14.v, 2);
        assert.equal(sheet1.G14.v, 5);
        assert.equal(sheet1.H14.v, 8);

        assert.equal(sheet1.F15.v, 3);
        assert.equal(sheet1.G15.v, 6);
        assert.equal(sheet1.H15.v, 9);
    });

    // it('fixes the fund.xlsx problem', function() {
    //     var workbook = XLSX.readFile('test/fund-2.xlsx');
    //     var formulajs = require('formulajs');
    //     XLSX_CALC.import_functions(formulajs);

    //     erase_values_that_contains_formula(workbook.Sheets);
    //     var original_sheet = XLSX.readFile('test/fund-2.xlsx').Sheets['Fund Economics'];
    //     XLSX_CALC(workbook);
    //     assert_values(original_sheet, workbook.Sheets['Fund Economics']);
    // });

});
