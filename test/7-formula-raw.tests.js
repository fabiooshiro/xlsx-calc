"use strict";

const XLSX_CALC = require("../src");
const assert = require('assert');

describe.only('raw formulas', () => {

    it('should filter', () => {
        let workbook = {
            Sheets: {
                Sheet1: {
                    A3: {v: 'aa'},
                    A4: {v: 'aaa'},
                    A5: {v: 'aaaa'},
                    B3: {v: 'bb'},
                    B4: {v: 'bbb'},
                    B5: {v: 'bbbb'},
                    C3: {v: 'cc'},
                    C4: {v: 'ccc'},
                    C5: {v: 'cccc'},

                    A9: {v: true},
                    B9: {v: false},
                    C9: {v: false},

                    E3: {f:  "FILTER(A3:C5,A9:C9)"},
                }
            }
        };
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets.Sheet1.E3.v, 'aa');
        assert.equal(workbook.Sheets.Sheet1.E4.v, 'aaa');
        assert.equal(workbook.Sheets.Sheet1.E5.v, 'aaaa');
    })
})