"use strict";

const XLSX_CALC = require("../src");
const assert = require('assert');

describe('matrix formulas', () => {

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
                    B9: {v: true},
                    C9: {v: false},

                    E3: {f:  "FILTER(A3:C5,A9:C9)"},
                }
            }
        };

        XLSX_CALC(workbook);

        assert.equal(workbook.Sheets.Sheet1.E3.v, 'aa');
        assert.equal(workbook.Sheets.Sheet1.E4.v, 'aaa');
        assert.equal(workbook.Sheets.Sheet1.E5.v, 'aaaa');
        assert.equal(workbook.Sheets.Sheet1.F3.v, 'bb');
        assert.equal(workbook.Sheets.Sheet1.F4.v, 'bbb');
        assert.equal(workbook.Sheets.Sheet1.F5.v, 'bbbb');
        assert.equal(workbook.Sheets.Sheet1.G3.v, '');
        assert.equal(workbook.Sheets.Sheet1.G4.v, '');
        assert.equal(workbook.Sheets.Sheet1.G5.v, '');
    });

    it('should filter with no match', () => {
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

                    A9: {v: false},
                    B9: {v: false},
                    C9: {v: false},

                    E3: {f:  "FILTER(A3:C5,A9:C9)"},
                }
            }
        };

        XLSX_CALC(workbook);

        assert.equal(workbook.Sheets.Sheet1.E3.t, 'e');
        //TODO: I dont know the code of this error
        assert.equal(workbook.Sheets.Sheet1.E3.v, 0);

        assert.equal(workbook.Sheets.Sheet1.E3.w, '#CALC!');
    });

    it('should filter with all match', () => {
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
                    B9: {v: true},
                    C9: {v: true},

                    E3: {f:  "FILTER(A3:C5,A9:C9)"},
                }
            }
        };

        XLSX_CALC(workbook);

        assert.equal(workbook.Sheets.Sheet1.E3.v, 'aa');
        assert.equal(workbook.Sheets.Sheet1.E4.v, 'aaa');
        assert.equal(workbook.Sheets.Sheet1.E5.v, 'aaaa');
        assert.equal(workbook.Sheets.Sheet1.F3.v, 'bb');
        assert.equal(workbook.Sheets.Sheet1.F4.v, 'bbb');
        assert.equal(workbook.Sheets.Sheet1.F5.v, 'bbbb');
        assert.equal(workbook.Sheets.Sheet1.G3.v, 'cc');
        assert.equal(workbook.Sheets.Sheet1.G4.v, 'ccc');
        assert.equal(workbook.Sheets.Sheet1.G5.v, 'cccc');
    });
})