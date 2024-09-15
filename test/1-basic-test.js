"use strict";

const XLSX_CALC = require("../src");
const assert = require('assert');
const errorValues = {
    '#NULL!': 0x00,
    '#DIV/0!': 0x07,
    '#VALUE!': 0x0F,
    '#REF!': 0x17,
    '#NAME?': 0x1D,
    '#NUM!': 0x24,
    '#N/A': 0x2A,
    '#GETTING_DATA': 0x2B
};

describe('XLSX_CALC', function() {
    let workbook;
    function create_workbook() {
        return {
            Sheets: {
                Sheet1: {
                    A1: {},
                    A2: {
                        v: 7
                    },
                    C2: {
                        v: 1
                    },
                    C3: {
                        v: 1
                    },
                    C4: {
                        v: 2
                    },
                    C5: {
                        v: 3
                    },
                }
            }
        };
    }
    beforeEach(function() {
        workbook = create_workbook();
    });

    describe('ROUND', () => {
        it('should round to 1', () => {
            workbook.Sheets.Sheet1.A1.f = 'ROUND(1.2)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 1);
        });
        it('should round to 1.2', () => {
            workbook.Sheets.Sheet1.A1.f = 'ROUND(1.23,1)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 1.2);
        });
    });

    describe('plus', function() {
        it('should calc A2+C5', function() {
            workbook.Sheets.Sheet1.A2.v = 7;
            workbook.Sheets.Sheet1.C5.v = 3;
            workbook.Sheets.Sheet1.A1.f = 'A2+C5';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 10);
        });
        it('should calc 1+2', function() {
            workbook.Sheets.Sheet1.A1.f = '1+2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
        });
        it('should calc 1+2+3', function() {
            workbook.Sheets.Sheet1.A1.f = '1+2+3';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 6);
        });

        it('should calc +A2+C5',function() {
            workbook.Sheets.Sheet1.A2.v = 7;
            workbook.Sheets.Sheet1.C5.v = 3;
            workbook.Sheets.Sheet1.A1.f = '+A2+C5';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 10);
        });
        it('should calc C1:C5+2', function() {
            workbook.Sheets.Sheet1.D1 = { f: 'C1:C5+2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 2);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 3);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 3);
            assert.equal(workbook.Sheets.Sheet1.D4.v, 4);
            assert.equal(workbook.Sheets.Sheet1.D5.v, 5);
        });
        it('should calc A1:A3+A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:A3+A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 2);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 3);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 4);
            assert.equal(workbook.Sheets.Sheet1.E1.v, 5);
            assert.equal(workbook.Sheets.Sheet1.E2.v, 6);
            assert.equal(workbook.Sheets.Sheet1.E3.v, 7);
            assert.equal(workbook.Sheets.Sheet1.F1.v, 8);
            assert.equal(workbook.Sheets.Sheet1.F2.v, 9);
            assert.equal(workbook.Sheets.Sheet1.F3.v, 10);
        });
        it('should calc A1:B3+A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:B3+A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 2);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 3);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 4);
            assert.equal(workbook.Sheets.Sheet1.E1.v, 8);
            assert.equal(workbook.Sheets.Sheet1.E2.v, 9);
            assert.equal(workbook.Sheets.Sheet1.E3.v, 10);
            assert.equal(workbook.Sheets.Sheet1.F1.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F2.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F3.v, errorValues['#N/A']);
        });
    });
    describe('minus', function() {
        it('should update the property A1.v with result of formula A2-C4', function() {
            workbook.Sheets.Sheet1.A1.f = 'A2-C4';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
        });
        it('should calc A2-4', function() {
            workbook.Sheets.Sheet1.A1.f = 'A2-4';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
        });
        it('should calc 2-3', function() {
            workbook.Sheets.Sheet1.A1.f = '2-3';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, -1);
        });
        it('should calc 2-3-4', function() {
            workbook.Sheets.Sheet1.A1.f = '2-3-4';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, -5);
        });
        it('should calc -2-3-4', function() {
            workbook.Sheets.Sheet1.A1.f = '-2-3-4';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, -9);
        });
        it('should calc C1:C5-2', function() {
            workbook.Sheets.Sheet1.D1 = { f: 'C1:C5-2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, -2);
            assert.equal(workbook.Sheets.Sheet1.D2.v, -1);
            assert.equal(workbook.Sheets.Sheet1.D3.v, -1);
            assert.equal(workbook.Sheets.Sheet1.D4.v, 0);
            assert.equal(workbook.Sheets.Sheet1.D5.v, 1);
        });
        it('should calc -C1:C5-2', function() {
            workbook.Sheets.Sheet1.D1 = { f: '-C1:C5-2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, -2);
            assert.equal(workbook.Sheets.Sheet1.D2.v, -3);
            assert.equal(workbook.Sheets.Sheet1.D3.v, -3);
            assert.equal(workbook.Sheets.Sheet1.D4.v, -4);
            assert.equal(workbook.Sheets.Sheet1.D5.v, -5);
        });
        it('should calc A1:A3-A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:A3-A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 0);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 1);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 2);
            assert.equal(workbook.Sheets.Sheet1.E1.v, -3);
            assert.equal(workbook.Sheets.Sheet1.E2.v, -2);
            assert.equal(workbook.Sheets.Sheet1.E3.v, -1);
            assert.equal(workbook.Sheets.Sheet1.F1.v, -6);
            assert.equal(workbook.Sheets.Sheet1.F2.v, -5);
            assert.equal(workbook.Sheets.Sheet1.F3.v, -4);
        });
        it('should calc A1:B3-A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:B3-A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 0);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 1);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 2);
            assert.equal(workbook.Sheets.Sheet1.E1.v, 0);
            assert.equal(workbook.Sheets.Sheet1.E2.v, 1);
            assert.equal(workbook.Sheets.Sheet1.E3.v, 2);
            assert.equal(workbook.Sheets.Sheet1.F1.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F2.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F3.v, errorValues['#N/A']);
        });
    });
    describe('multiply', function() {
        it('should calc A2*C5', function() {
            workbook.Sheets.Sheet1.A1.f = 'A2*C5';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 21);
        });
        it('should calc A2*4', function() {
            workbook.Sheets.Sheet1.A1.f = 'A2*4';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 28);
        });
        it('should calc 4*A2', function() {
            workbook.Sheets.Sheet1.A1.f = '4*A2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 28);
        });
        it('should calc 2*3', function() {
            workbook.Sheets.Sheet1.A1.f = '2*3';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 6);
        });
        it('should calc C1:C5*2', function() {
            workbook.Sheets.Sheet1.D1 = { f: 'C1:C5*2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 0);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 2);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 2);
            assert.equal(workbook.Sheets.Sheet1.D4.v, 4);
            assert.equal(workbook.Sheets.Sheet1.D5.v, 6);
        });
        it('should calc A1:A3*A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:A3*A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 1);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 2);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 3);
            assert.equal(workbook.Sheets.Sheet1.E1.v, 4);
            assert.equal(workbook.Sheets.Sheet1.E2.v, 8);
            assert.equal(workbook.Sheets.Sheet1.E3.v, 12);
            assert.equal(workbook.Sheets.Sheet1.F1.v, 7);
            assert.equal(workbook.Sheets.Sheet1.F2.v, 14);
            assert.equal(workbook.Sheets.Sheet1.F3.v, 21);
        });
        it('should calc A1:B3*A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:B3*A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 1);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 2);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 3);
            assert.equal(workbook.Sheets.Sheet1.E1.v, 16);
            assert.equal(workbook.Sheets.Sheet1.E2.v, 20);
            assert.equal(workbook.Sheets.Sheet1.E3.v, 24);
            assert.equal(workbook.Sheets.Sheet1.F1.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F2.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F3.v, errorValues['#N/A']);
        });
    });
    describe('divide', function() {
        it('should calc A2/C4', function() {
            workbook.Sheets.Sheet1.A1.f = 'A2/C4';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 3.5);
        });
        it('should calc A2/14', function() {
            workbook.Sheets.Sheet1.A1.f = 'A2/14';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 0.5);
        });
        it('should calc 7/2/2', function() {
            workbook.Sheets.Sheet1.A1.f = '7/2/2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 1.75);
        });
        it('should divide', function() {
            workbook.Sheets.Sheet1.A1.f = '=20/10';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 2);
        });
        it('should calc C1:C5/2', function() {
            workbook.Sheets.Sheet1.D1 = { f: 'C1:C5/2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 0);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 0.5);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 0.5);
            assert.equal(workbook.Sheets.Sheet1.D4.v, 1);
            assert.equal(workbook.Sheets.Sheet1.D5.v, 1.5);
        });
        it('should calc A1:A3/A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:A3/A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 1);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 2);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 3);
            assert.equal(workbook.Sheets.Sheet1.E1.v, 0.25);
            assert.equal(workbook.Sheets.Sheet1.E2.v, 0.5);
            assert.equal(workbook.Sheets.Sheet1.E3.v, 0.75);
            assert.equal(workbook.Sheets.Sheet1.F1.v, 1/7); // 0.14285714285714285
            assert.equal(workbook.Sheets.Sheet1.F2.v, 2/7); // 0.2857142857142857
            assert.equal(workbook.Sheets.Sheet1.F3.v, 3/7); // 0.42857142857142855
        });
        it('should calc A1:B3/A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:B3/A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 1);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 2);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 3);
            assert.equal(workbook.Sheets.Sheet1.E1.v, 1);
            assert.equal(workbook.Sheets.Sheet1.E2.v, 1.25);
            assert.equal(workbook.Sheets.Sheet1.E3.v, 1.5);
            assert.equal(workbook.Sheets.Sheet1.F1.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F2.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F3.v, errorValues['#N/A']);
        });
    });
    describe('power', function() {
        it('should calc 2^10', function() {
            workbook.Sheets.Sheet1.A1.f = '2^10';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 1024);
        });
        it('should calc C1:C5^2', function() {
            workbook.Sheets.Sheet1.D1 = { f: 'C1:C5^2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 0);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 1);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 1);
            assert.equal(workbook.Sheets.Sheet1.D4.v, 4);
            assert.equal(workbook.Sheets.Sheet1.D5.v, 9);
        });
        it('should calc A1:A3^A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:A3^A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 1);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 2);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 3);
            assert.equal(workbook.Sheets.Sheet1.E1.v, 1);
            assert.equal(workbook.Sheets.Sheet1.E2.v, 16);
            assert.equal(workbook.Sheets.Sheet1.E3.v, 81);
            assert.equal(workbook.Sheets.Sheet1.F1.v, 1);
            assert.equal(workbook.Sheets.Sheet1.F2.v, 128);
            assert.equal(workbook.Sheets.Sheet1.F3.v, 2187);
        });
        it('should calc A1:B3^A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:B3^A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 1);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 2);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 3);
            assert.equal(workbook.Sheets.Sheet1.E1.v, 256);
            assert.equal(workbook.Sheets.Sheet1.E2.v, 625);
            assert.equal(workbook.Sheets.Sheet1.E3.v, 1296);
            assert.equal(workbook.Sheets.Sheet1.F1.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F2.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F3.v, errorValues['#N/A']);
        });
    });
    describe('SQRT', function() {
        it('should calc SQRT(25)', function() {
            workbook.Sheets.Sheet1.A1.f = 'SQRT(25)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
        });
    });
    describe('ABS', function() {
        it('should calc ABS(-3.5)', function() {
            workbook.Sheets.Sheet1.A1.f = 'ABS(-3.5)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 3.5);
        });
    });
    describe('FLOOR', function() {
        it('should calc FLOOR(12.5)', function() {
            workbook.Sheets.Sheet1.A1.f = 'FLOOR(12.5)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 12);
        });
    });
    describe('.set_fx', function() {
        it('sets new function', function() {
            XLSX_CALC.set_fx('ADD_1', function(arg) {
                return arg + 1;
            });
            workbook.Sheets.Sheet1.A1.f = 'ADD_1(123)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 124);
        });
    });
    describe('expression', function() {
        it('should calc 8/2+1', function() {
            workbook.Sheets.Sheet1.A1.f = '8/2+1';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
        });
        it('should calc 1+8/2', function() {
            workbook.Sheets.Sheet1.A1.f = '1+8/2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
        });
        it('should calc 2*3+1', function() {
            workbook.Sheets.Sheet1.A1.f = '2*3+1';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 7);
        });
        it('should calc 2*3-1', function() {
            workbook.Sheets.Sheet1.A1.f = '2*3-1';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
        });
        it('should calc 2*(3-1)', function() {
            workbook.Sheets.Sheet1.A1.f = '2*(3-1)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 4);
        });
        it('should calc (3-1)*5', function() {
            workbook.Sheets.Sheet1.A1.f = '(3-1)*5';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 10);
        });
        it('should calc (3-1)*(4+1)', function() {
            workbook.Sheets.Sheet1.A1.f = '(3-1)*(4+1)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 10);
        });
        it('should calc -1*2', function() {
            workbook.Sheets.Sheet1.A1.f = '-1*2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, -2);
        });
        it('should calc 1*-2', function() {
            workbook.Sheets.Sheet1.A1.f = '1*-2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, -2);
        });
        it('should calc (3*10)-(2-1)', function() {
            workbook.Sheets.Sheet1.A1.f = '(3*10)-(2-1)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 29);
        });
        it('should calc (3*10)-(2-(3*5))', function() {
            workbook.Sheets.Sheet1.A1.f = '(3*10)-(2-(3*5))';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 43);
        });
        it('should calc (3*10)-(-2-(3*5))', function() {
            workbook.Sheets.Sheet1.A1.f = '(3*10)-(-2-(3*5))';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 47);
        });
        it('should calc 8/2*10', function () {
            workbook.Sheets.Sheet1.A1.f = '8/2*10';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 40);
        });
        it('should calc (C1:C5+3)*2', function() {
            workbook.Sheets.Sheet1.D1 = { f: '(C1:C5+3)*2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 6);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 8);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 8);
            assert.equal(workbook.Sheets.Sheet1.D4.v, 10);
            assert.equal(workbook.Sheets.Sheet1.D5.v, 12);
        });
        it('should calc (A1:A3+3)*A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: '(A1:A3+3)*A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 4);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 5);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 6);
            assert.equal(workbook.Sheets.Sheet1.E1.v, 16);
            assert.equal(workbook.Sheets.Sheet1.E2.v, 20);
            assert.equal(workbook.Sheets.Sheet1.E3.v, 24);
            assert.equal(workbook.Sheets.Sheet1.F1.v, 28);
            assert.equal(workbook.Sheets.Sheet1.F2.v, 35);
            assert.equal(workbook.Sheets.Sheet1.F3.v, 42);
        });
        it('should calc (A1:B3+3)*A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: '(A1:B3+3)*A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 4);
            assert.equal(workbook.Sheets.Sheet1.D2.v, 5);
            assert.equal(workbook.Sheets.Sheet1.D3.v, 6);
            assert.equal(workbook.Sheets.Sheet1.E1.v, 28);
            assert.equal(workbook.Sheets.Sheet1.E2.v, 32);
            assert.equal(workbook.Sheets.Sheet1.E3.v, 36);
            assert.equal(workbook.Sheets.Sheet1.F1.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F2.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F3.v, errorValues['#N/A']);
        });
    });
    describe('SUM', function() {
        it('makes the sum', function() {
            workbook.Sheets.Sheet1.A1.f = 'SUM(C3:C4)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
        });
        it('makes the sum of a bigger range', function() {
            workbook.Sheets.Sheet1.A1.f = 'SUM(C3:C5)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 6);
        });
        it('sums numbers', function() {
            workbook.Sheets.Sheet1.A1.f = 'SUM(1,2,3)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 6);
        });
    });
    describe('MAX', function() {
        it('finds the max in range', function() {
            workbook.Sheets.Sheet1.A1.f = 'MAX( C3:C5 )';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
        });
        it('finds the max in range including some cell', function() {
            workbook.Sheets.Sheet1.A1.f = 'MAX(C3:C5 ,A2)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 7);
        });
        it('finds the max in range including some cell', function() {
            workbook.Sheets.Sheet1.A1.f = 'MAX(A2,C3:C5)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 7);
        });
        it('finds the max in args', function() {
            workbook.Sheets.Sheet1.A1.f = 'MAX(1,2,10,3,4)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 10);
        });
        it('finds the max in negative args', function() {
            workbook.Sheets.Sheet1.A1.f = 'MAX(-1,-2,-10,-3,-4)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, -1);
        });
        it('finds the max in range including some negative cell', function() {
            workbook.Sheets.Sheet1.A1.f = 'MAX(C3:C5,-A2)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
        });
        it('finds the max in 2 dimensionnal range', function () {
            workbook.Sheets.Sheet1.A1.f = 'MAX(A2:C5)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 7);
        });
    });
    describe('MIN', function() {
        it('finds the min in range', function() {
            workbook.Sheets.Sheet1.A1.f = 'MIN(C3:C5)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 1);
        });
        it('finds the min in range including some negative cell', function() {
            workbook.Sheets.Sheet1.A1.f = 'MIN(C3:C5,-A2)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, -7);
        });
        it('finds the min in 2 dimensionnal range', function () {
            workbook.Sheets.Sheet1.A1.f = 'MIN(A2:C5)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 1);
        });
    });
    describe('MAX and SUM', function() {
        it('evaluates MAX(1,2,SUM(10,5),7,3,4)', function() {
            workbook.Sheets.Sheet1.A1.f = 'MAX(1,2,SUM(10,5),7,3,4)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 15);
        });
        it('evaluates MAX(1,2, SUM(10,5),7,3,4)', function() {
            workbook.Sheets.Sheet1.A1.f = 'MAX(1,2, SUM(10,5),7,3,4)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 15);
        });
    });
    describe('&', function() {
        it('evaluates "concat "&A2', function() {
            workbook.Sheets.Sheet1.A1.f = '"concat " & A2';
            workbook.Sheets.Sheet1.A2.v = 7;
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'concat 7');
        });
        it('evaluates "concat +1" & A2', function() {
            workbook.Sheets.Sheet1.A1.f = '"concat +1" & A2';
            workbook.Sheets.Sheet1.A2.v = 7;
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'concat +17');
        });
        it('evaluates A2 & "concat"', function() {
            workbook.Sheets.Sheet1.A1.f = 'A2 & "concat +1"';
            workbook.Sheets.Sheet1.A2.v = 7;
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, '7concat +1');
        });
        it('should calc C1:C5 & "concat"', function() {
            workbook.Sheets.Sheet1.D1 = { f: 'C1:C5 & "concat"' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 'concat');
            assert.equal(workbook.Sheets.Sheet1.D2.v, '1concat');
            assert.equal(workbook.Sheets.Sheet1.D3.v, '1concat');
            assert.equal(workbook.Sheets.Sheet1.D4.v, '2concat');
            assert.equal(workbook.Sheets.Sheet1.D5.v, '3concat');
        });
        it('should calc A1:A3&A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:A3&A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, '11');
            assert.equal(workbook.Sheets.Sheet1.D2.v, '21');
            assert.equal(workbook.Sheets.Sheet1.D3.v, '31');
            assert.equal(workbook.Sheets.Sheet1.E1.v, '14');
            assert.equal(workbook.Sheets.Sheet1.E2.v, '24');
            assert.equal(workbook.Sheets.Sheet1.E3.v, '34');
            assert.equal(workbook.Sheets.Sheet1.F1.v, '17');
            assert.equal(workbook.Sheets.Sheet1.F2.v, '27');
            assert.equal(workbook.Sheets.Sheet1.F3.v, '37');
        });
        it('should calc A1:B3&A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:B3&A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, '11');
            assert.equal(workbook.Sheets.Sheet1.D2.v, '21');
            assert.equal(workbook.Sheets.Sheet1.D3.v, '31');
            assert.equal(workbook.Sheets.Sheet1.E1.v, '44');
            assert.equal(workbook.Sheets.Sheet1.E2.v, '54');
            assert.equal(workbook.Sheets.Sheet1.E3.v, '64');
            assert.equal(workbook.Sheets.Sheet1.F1.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F2.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F3.v, errorValues['#N/A']);
        });
    });
    describe('CONCATENATE', function() {
        it('concatenates 1,2,3', function() {
            workbook.Sheets.Sheet1.A1.f = 'CONCATENATE(1,2,3)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, '123');
        });
        it('concatenates A2,"xxx"', function() {
            workbook.Sheets.Sheet1.A1.f = 'CONCATENATE(A2 , "xxx")';
            workbook.Sheets.Sheet1.A2.v = 79;
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, '79xxx');
        });
        it('concatenates null and undefined values as empty', function () {
            workbook.Sheets.Sheet1.A1 = { f: 'CONCATENATE(A2, "-", B2, "-", C2, "-", D2)' };
            workbook.Sheets.Sheet1.A2 = { v: 79 };
            workbook.Sheets.Sheet1.B2 = { v: null };
            workbook.Sheets.Sheet1.C2 = {};
            workbook.Sheets.Sheet1.D2 = { v: 'tutu' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, '79---tutu');
        })
    });
    describe('range', function() {
        it('should eval the expression in range of sum', function() {
            workbook.Sheets.Sheet1.A1.f = 'SUM(C3:C4)';
            workbook.Sheets.Sheet1.C4.f = 'A2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 8);
            assert.equal(workbook.Sheets.Sheet1.C4.v, 7);
        });
        it('should calc range with $', function() {
            workbook.Sheets.Sheet1.A1.f = 'SUM($C$3:$C$4)';
            workbook.Sheets.Sheet1.C4.f = 'A2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 8);
            assert.equal(workbook.Sheets.Sheet1.C4.v, 7);
        });
        it('should calc range like C:C using !ref', function() {
            workbook.Sheets.Sheet1['!ref'] = 'A1:C4';
            workbook.Sheets.Sheet1.A1.f = 'SUM(C:C)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 4);
        });
        it('should calc range like C:C without !ref', function() {
            this.timeout(5000);
            workbook.Sheets.Sheet1.A1.f = 'SUM(C:C)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 7);
        });
        it('A1:A3 formula should return spilled array', function () {
            // workbook.Sheets.Sheet1.A1 = { t: "e", v: 42, w: "#N/A" };
            workbook.Sheets.Sheet1.A1 = { t: 'n', v: 3 };
            workbook.Sheets.Sheet1.A2 = { t: 'n', v: 1 };
            workbook.Sheets.Sheet1.A3 = { t: 'n', v: 2 };
            workbook.Sheets.Sheet1.B1 = { f: 'A1:A3' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.B1.t, 'n');
            assert.equal(workbook.Sheets.Sheet1.B1.v, 3);
            assert.equal(workbook.Sheets.Sheet1.B2.t, 'n');
            assert.equal(workbook.Sheets.Sheet1.B2.v, 1);
            assert.equal(workbook.Sheets.Sheet1.B3.t, 'n');
            assert.equal(workbook.Sheets.Sheet1.B3.v, 2);
        });
        it('should contain error when a cell is in error', function () {
            workbook.Sheets.Sheet1.A1 = { t: "e", v: 42, w: "#N/A" };
            workbook.Sheets.Sheet1.A2 = { t: 'n', v: 1 };
            workbook.Sheets.Sheet1.A3 = { t: 'n', v: 2 };
            workbook.Sheets.Sheet1.B1 = { f: 'A1:A3' };
            var exec_formula = require('../src/exec_formula.js'),
            find_all_cells_with_formulas = require('../src/find_all_cells_with_formulas.js'),
            cache = require('../src/Range.js').cache;
            cache.clear();

            var formula = find_all_cells_with_formulas(workbook, exec_formula)[0];
            var range = exec_formula.build_expression(formula).args[0].calc();
            var expected = [
                [new Error('#N/A')],
                [1],
                [2]
            ];
            assert.deepEqual(range, expected);
        });
    });
    describe('boolean', function() {
        it('evaluates 1<2 as true', function() {
            workbook.Sheets.Sheet1.A1.f = '1<2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, true);
        });
        it('evaluates C1:C5<2 as true,true,true,false,false', function() {
            workbook.Sheets.Sheet1.D1 = { f: 'C1:C5<2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, true);
            assert.equal(workbook.Sheets.Sheet1.D2.v, true);
            assert.equal(workbook.Sheets.Sheet1.D3.v, true);
            assert.equal(workbook.Sheets.Sheet1.D4.v, false);
            assert.equal(workbook.Sheets.Sheet1.D5.v, false);
        });
        it('evaluates A1:A3<A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:A3<A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, false);
            assert.equal(workbook.Sheets.Sheet1.D2.v, false);
            assert.equal(workbook.Sheets.Sheet1.D3.v, false);
            assert.equal(workbook.Sheets.Sheet1.E1.v, true);
            assert.equal(workbook.Sheets.Sheet1.E2.v, true);
            assert.equal(workbook.Sheets.Sheet1.E3.v, true);
            assert.equal(workbook.Sheets.Sheet1.F1.v, true);
            assert.equal(workbook.Sheets.Sheet1.F2.v, true);
            assert.equal(workbook.Sheets.Sheet1.F3.v, true);
        });
        it('evaluates A1:B3<A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:B3<A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, false);
            assert.equal(workbook.Sheets.Sheet1.D2.v, false);
            assert.equal(workbook.Sheets.Sheet1.D3.v, false);
            assert.equal(workbook.Sheets.Sheet1.E1.v, false);
            assert.equal(workbook.Sheets.Sheet1.E2.v, false);
            assert.equal(workbook.Sheets.Sheet1.E3.v, false);
            assert.equal(workbook.Sheets.Sheet1.F1.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F2.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F3.v, errorValues['#N/A']);
        });
        it('evaluates 1>2 as false', function() {
            workbook.Sheets.Sheet1.A1.f = '1>2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, false);
        });
        it('evaluates C1:C5>2 as false,false,false,false,true', function() {
            workbook.Sheets.Sheet1.D1 = { f: 'C1:C5>2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, false);
            assert.equal(workbook.Sheets.Sheet1.D2.v, false);
            assert.equal(workbook.Sheets.Sheet1.D3.v, false);
            assert.equal(workbook.Sheets.Sheet1.D4.v, false);
            assert.equal(workbook.Sheets.Sheet1.D5.v, true);
        });
        it('evaluates A1:A3>A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:A3>A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, false);
            assert.equal(workbook.Sheets.Sheet1.D2.v, true);
            assert.equal(workbook.Sheets.Sheet1.D3.v, true);
            assert.equal(workbook.Sheets.Sheet1.E1.v, false);
            assert.equal(workbook.Sheets.Sheet1.E2.v, false);
            assert.equal(workbook.Sheets.Sheet1.E3.v, false);
            assert.equal(workbook.Sheets.Sheet1.F1.v, false);
            assert.equal(workbook.Sheets.Sheet1.F2.v, false);
            assert.equal(workbook.Sheets.Sheet1.F3.v, false);
        });
        it('evaluates A1:B3>A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:B3>A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, false);
            assert.equal(workbook.Sheets.Sheet1.D2.v, true);
            assert.equal(workbook.Sheets.Sheet1.D3.v, true);
            assert.equal(workbook.Sheets.Sheet1.E1.v, false);
            assert.equal(workbook.Sheets.Sheet1.E2.v, true);
            assert.equal(workbook.Sheets.Sheet1.E3.v, true);
            assert.equal(workbook.Sheets.Sheet1.F1.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F2.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F3.v, errorValues['#N/A']);
        });
        it('evaluates 2=2 as true', function() {
            workbook.Sheets.Sheet1.A1.f = '2=2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, true);
        });
        it('evaluates C1:C5=2 as false,false,false,true,false', function() {
            workbook.Sheets.Sheet1.D1 = { f: 'C1:C5=2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, false);
            assert.equal(workbook.Sheets.Sheet1.D2.v, false);
            assert.equal(workbook.Sheets.Sheet1.D3.v, false);
            assert.equal(workbook.Sheets.Sheet1.D4.v, true);
            assert.equal(workbook.Sheets.Sheet1.D5.v, false);
        });
        it('evaluates A1:A3=A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:A3=A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, true);
            assert.equal(workbook.Sheets.Sheet1.D2.v, false);
            assert.equal(workbook.Sheets.Sheet1.D3.v, false);
            assert.equal(workbook.Sheets.Sheet1.E1.v, false);
            assert.equal(workbook.Sheets.Sheet1.E2.v, false);
            assert.equal(workbook.Sheets.Sheet1.E3.v, false);
            assert.equal(workbook.Sheets.Sheet1.F1.v, false);
            assert.equal(workbook.Sheets.Sheet1.F2.v, false);
            assert.equal(workbook.Sheets.Sheet1.F3.v, false);
        });
        it('evaluates A1:B3=A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:B3=A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, true);
            assert.equal(workbook.Sheets.Sheet1.D2.v, false);
            assert.equal(workbook.Sheets.Sheet1.D3.v, false);
            assert.equal(workbook.Sheets.Sheet1.E1.v, true);
            assert.equal(workbook.Sheets.Sheet1.E2.v, false);
            assert.equal(workbook.Sheets.Sheet1.E3.v, false);
            assert.equal(workbook.Sheets.Sheet1.F1.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F2.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F3.v, errorValues['#N/A']);
        });
        it('evaluates 2>=2 as true', function() {
            workbook.Sheets.Sheet1.A1.f = '2>=2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, true);
        });
        it('evaluates 1>=2 as false', function() {
            workbook.Sheets.Sheet1.A1.f = '1>=2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, false);
        });
        it('evaluates C1:C5>=2 as false,false,false,true,true', function() {
            workbook.Sheets.Sheet1.D1 = { f: 'C1:C5>=2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, false);
            assert.equal(workbook.Sheets.Sheet1.D2.v, false);
            assert.equal(workbook.Sheets.Sheet1.D3.v, false);
            assert.equal(workbook.Sheets.Sheet1.D4.v, true);
            assert.equal(workbook.Sheets.Sheet1.D5.v, true);
        });
        it('evaluates A1:A3>=A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:A3>=A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, true);
            assert.equal(workbook.Sheets.Sheet1.D2.v, true);
            assert.equal(workbook.Sheets.Sheet1.D3.v, true);
            assert.equal(workbook.Sheets.Sheet1.E1.v, false);
            assert.equal(workbook.Sheets.Sheet1.E2.v, false);
            assert.equal(workbook.Sheets.Sheet1.E3.v, false);
            assert.equal(workbook.Sheets.Sheet1.F1.v, false);
            assert.equal(workbook.Sheets.Sheet1.F2.v, false);
            assert.equal(workbook.Sheets.Sheet1.F3.v, false);
        });
        it('evaluates A1:B3>=A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:B3>=A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, true);
            assert.equal(workbook.Sheets.Sheet1.D2.v, true);
            assert.equal(workbook.Sheets.Sheet1.D3.v, true);
            assert.equal(workbook.Sheets.Sheet1.E1.v, true);
            assert.equal(workbook.Sheets.Sheet1.E2.v, true);
            assert.equal(workbook.Sheets.Sheet1.E3.v, true);
            assert.equal(workbook.Sheets.Sheet1.F1.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F2.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F3.v, errorValues['#N/A']);
        });
        it('evaluates 2<=2 as true', function() {
            workbook.Sheets.Sheet1.A1.f = '2<=2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, true);
        });
        it('evaluates 3<=2 as false', function() {
            workbook.Sheets.Sheet1.A1.f = '3<=2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, false);
        });
        it('evaluates C3<=C5 as false', function() {
            workbook.Sheets.Sheet1.C3.v = 3;
            workbook.Sheets.Sheet1.C5.v = 2;
            workbook.Sheets.Sheet1.A1.f = 'C3<=C5';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, false);
        });
        it('evaluates C1:C5<=2 as true,true,true,true,false', function() {
            workbook.Sheets.Sheet1.D1 = { f: 'C1:C5<=2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, true);
            assert.equal(workbook.Sheets.Sheet1.D2.v, true);
            assert.equal(workbook.Sheets.Sheet1.D3.v, true);
            assert.equal(workbook.Sheets.Sheet1.D4.v, true);
            assert.equal(workbook.Sheets.Sheet1.D5.v, false);
        });
        it('evaluates A1:A3<=A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:A3<=A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, true);
            assert.equal(workbook.Sheets.Sheet1.D2.v, false);
            assert.equal(workbook.Sheets.Sheet1.D3.v, false);
            assert.equal(workbook.Sheets.Sheet1.E1.v, true);
            assert.equal(workbook.Sheets.Sheet1.E2.v, true);
            assert.equal(workbook.Sheets.Sheet1.E3.v, true);
            assert.equal(workbook.Sheets.Sheet1.F1.v, true);
            assert.equal(workbook.Sheets.Sheet1.F2.v, true);
            assert.equal(workbook.Sheets.Sheet1.F3.v, true);
        });
        it('evaluates A1:B3<=A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:B3<=A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, true);
            assert.equal(workbook.Sheets.Sheet1.D2.v, false);
            assert.equal(workbook.Sheets.Sheet1.D3.v, false);
            assert.equal(workbook.Sheets.Sheet1.E1.v, true);
            assert.equal(workbook.Sheets.Sheet1.E2.v, false);
            assert.equal(workbook.Sheets.Sheet1.E3.v, false);
            assert.equal(workbook.Sheets.Sheet1.F1.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F2.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F3.v, errorValues['#N/A']);
        });
        it('evaluates 1<>1 as false', function() {
            workbook.Sheets.Sheet1.A1.f = '1<>1';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, false);
        });
        it('evaluates 2<>1 as true', function() {
            workbook.Sheets.Sheet1.A1.f = '2<>1';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, true);
        });
        it('evaluates C1:C5<>2 as true,true,true,false,true', function() {
            workbook.Sheets.Sheet1.D1 = { f: 'C1:C5<>2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, true);
            assert.equal(workbook.Sheets.Sheet1.D2.v, true);
            assert.equal(workbook.Sheets.Sheet1.D3.v, true);
            assert.equal(workbook.Sheets.Sheet1.D4.v, false);
            assert.equal(workbook.Sheets.Sheet1.D5.v, true);
        });
        it('evaluates A1:A3<>A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:A3<>A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, false);
            assert.equal(workbook.Sheets.Sheet1.D2.v, true);
            assert.equal(workbook.Sheets.Sheet1.D3.v, true);
            assert.equal(workbook.Sheets.Sheet1.E1.v, true);
            assert.equal(workbook.Sheets.Sheet1.E2.v, true);
            assert.equal(workbook.Sheets.Sheet1.E3.v, true);
            assert.equal(workbook.Sheets.Sheet1.F1.v, true);
            assert.equal(workbook.Sheets.Sheet1.F2.v, true);
            assert.equal(workbook.Sheets.Sheet1.F3.v, true);
        });
        it('evaluates A1:B3<>A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:B3<>A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, false);
            assert.equal(workbook.Sheets.Sheet1.D2.v, true);
            assert.equal(workbook.Sheets.Sheet1.D3.v, true);
            assert.equal(workbook.Sheets.Sheet1.E1.v, false);
            assert.equal(workbook.Sheets.Sheet1.E2.v, true);
            assert.equal(workbook.Sheets.Sheet1.E3.v, true);
            assert.equal(workbook.Sheets.Sheet1.F1.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F2.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F3.v, errorValues['#N/A']);
        });
        it('evaluates C1:C5=2 as false,false,false,true,false', function() {
            workbook.Sheets.Sheet1.D1 = { f: 'C1:C5=2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, false);
            assert.equal(workbook.Sheets.Sheet1.D2.v, false);
            assert.equal(workbook.Sheets.Sheet1.D3.v, false);
            assert.equal(workbook.Sheets.Sheet1.D4.v, true);
            assert.equal(workbook.Sheets.Sheet1.D5.v, false);
        });
        it('evaluates A1:A3=A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:A3=A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, true);
            assert.equal(workbook.Sheets.Sheet1.D2.v, false);
            assert.equal(workbook.Sheets.Sheet1.D3.v, false);
            assert.equal(workbook.Sheets.Sheet1.E1.v, false);
            assert.equal(workbook.Sheets.Sheet1.E2.v, false);
            assert.equal(workbook.Sheets.Sheet1.E3.v, false);
            assert.equal(workbook.Sheets.Sheet1.F1.v, false);
            assert.equal(workbook.Sheets.Sheet1.F2.v, false);
            assert.equal(workbook.Sheets.Sheet1.F3.v, false);
        });
        it('evaluates A1:B3=A1:C1', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 6 };
            workbook.Sheets.Sheet1.C1 = { v: 7 };
            workbook.Sheets.Sheet1.D1 = { f: 'A1:B3=A1:C1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, true);
            assert.equal(workbook.Sheets.Sheet1.D2.v, false);
            assert.equal(workbook.Sheets.Sheet1.D3.v, false);
            assert.equal(workbook.Sheets.Sheet1.E1.v, true);
            assert.equal(workbook.Sheets.Sheet1.E2.v, false);
            assert.equal(workbook.Sheets.Sheet1.E3.v, false);
            assert.equal(workbook.Sheets.Sheet1.F1.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F2.v, errorValues['#N/A']);
            assert.equal(workbook.Sheets.Sheet1.F3.v, errorValues['#N/A']);
        });
        it('evaluates undefined<>"" as false', function() {
            workbook.Sheets.Sheet1.A2.f = 'A1<>""';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, false);
        })
        it('evaluates undefined="" as true', function() {
            workbook.Sheets.Sheet1.A2.f = 'A1=""';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, true);
        });
        it('evaluates 0<>"" as true', function() {
            workbook.Sheets.Sheet1.A1.f = '0<>""';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, true);
        });
    });
    describe('dates', function () {
        it('dateA - dateB should calc day diff', function() {
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: new Date('2019-01-10'),
                w: '2019-01-10'
            };
            workbook.Sheets.Sheet1.A2 = {
                t: 'd',
                v: new Date('2019-01-15'),
                w: '2019-01-15'
            };
            workbook.Sheets.Sheet1.A3 = { f: 'A2 - A1' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A3.v, 5);
            assert.equal(workbook.Sheets.Sheet1.A3.t, 'n');
        });
        it('DateA + 5 should add 5 days to dateA', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: new Date('2019-01-10'),
                w: '2019-01-10'
            };
            workbook.Sheets.Sheet1.A2 = { f: 'A1 + 5' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.t, 'n');
            assert.equal(workbook.Sheets.Sheet1.A2.v, Date.parse(new Date('2019-01-15')));
        });
        it('< operator should work for dates', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: new Date('2019-01-10'),
                w: '2019-01-10'
            };
            workbook.Sheets.Sheet1.A2 = {
                t: 'd',
                v: new Date('2019-01-15'),
                w: '2019-01-15'
            };
            workbook.Sheets.Sheet1.A3 = {
                t: 'd',
                v: new Date('2019-01-10'),
                w: '2019-01-10'
            };
            workbook.Sheets.Sheet1.A4 = { f: 'A1 < A2' };
            workbook.Sheets.Sheet1.A5 = { f: 'A1 < A3' };
            workbook.Sheets.Sheet1.A6 = { f: 'A2 < A3' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A4.v, true);
            assert.equal(workbook.Sheets.Sheet1.A5.v, false);
            assert.equal(workbook.Sheets.Sheet1.A6.v, false);
        });
        it('> operator should work for dates', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: new Date('2019-01-10'),
                w: '2019-01-10'
            };
            workbook.Sheets.Sheet1.A2 = {
                t: 'd',
                v: new Date('2019-01-15'),
                w: '2019-01-15'
            };
            workbook.Sheets.Sheet1.A3 = {
                t: 'd',
                v: new Date('2019-01-10'),
                w: '2019-01-10'
            };
            workbook.Sheets.Sheet1.A4 = { f: 'A2 > A1' };
            workbook.Sheets.Sheet1.A5 = { f: 'A3 > A1' };
            workbook.Sheets.Sheet1.A6 = { f: 'A3 > A2' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A4.v, true);
            assert.equal(workbook.Sheets.Sheet1.A5.v, false);
            assert.equal(workbook.Sheets.Sheet1.A6.v, false);
        });
        it('<> operator should work for dates', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: new Date('2019-01-10'),
                w: '2019-01-10'
            };
            workbook.Sheets.Sheet1.A2 = {
                t: 'd',
                v: new Date('2019-01-15'),
                w: '2019-01-15'
            };
            workbook.Sheets.Sheet1.A3 = {
                t: 'd',
                v: new Date('2019-01-10'),
                w: '2019-01-10'
            };
            workbook.Sheets.Sheet1.A4 = { f: 'A1 <> A2' };
            workbook.Sheets.Sheet1.A5 = { f: 'A1 <> A3' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A4.v, true);
            assert.equal(workbook.Sheets.Sheet1.A5.v, false);
        });
        it('= operator should work for dates', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: new Date('2019-01-10'),
                w: '2019-01-10'
            };
            workbook.Sheets.Sheet1.A2 = {
                t: 'd',
                v: new Date('2019-01-15'),
                w: '2019-01-15'
            };
            workbook.Sheets.Sheet1.A3 = {
                t: 'd',
                v: new Date('2019-01-10'),
                w: '2019-01-10'
            };
            workbook.Sheets.Sheet1.A4 = { f: 'A1 = A2' };
            workbook.Sheets.Sheet1.A5 = { f: 'A1 = A3' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A4.v, false);
            assert.equal(workbook.Sheets.Sheet1.A5.v, true);
        });
        it('>= should compare with today', function () {
            var today = new Date();
            today.setHours(0,0,0,0);
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: today
            };
            workbook.Sheets.Sheet1.A2 = { f: 'A1 >= TODAY()'};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, true);
        });
        it('<= should compare with today', function () {
            var today = new Date();
            today.setHours(0,0,0,0);
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: today
            };
            workbook.Sheets.Sheet1.A2 = { f: 'A1 <= TODAY()'};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, true);
        });
        it('date + TIME(0, integer, 0) should return a date with increased minutes', function () {
            workbook.Sheets.Sheet1 = {
                A1: {
                    t: "d",
                    v: new Date('2019-02-18 08:00'),
                    w: "2019-02-18 08:00"
                },
                A2: { v: "30" },
                B1: { f: "A1+TIME(0,A2,0)"}
            };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.B1.t, 'n');
            assert.equal(workbook.Sheets.Sheet1.B1.v, Date.parse(new Date('2019-02-18 08:30')));
        });
        xit('MIN, MAX should work for dates', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: new Date('2019-01-10'),
                w: '2019-01-10'
            };
            workbook.Sheets.Sheet1.A2 = {
                t: 'd',
                v: new Date('2019-01-11'),
                w: '2019-01-11'
            };
            workbook.Sheets.Sheet1.A3 = {
                t: 'd',
                v: new Date('2019-01-12'),
                w: '2019-01-12'
            };
            workbook.Sheets.Sheet1.A4 = { f: 'MAX(A1:A3)' };
            workbook.Sheets.Sheet1.A5 = { f: 'MIN(A1:A3)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A4.v instanceof Date, true);
            assert.equal(workbook.Sheets.Sheet1.A4.v.getTime(), 1547251200000);
            assert.equal(workbook.Sheets.Sheet1.A5.v instanceof Date, true);
            assert.equal(workbook.Sheets.Sheet1.A5.v.getTime(), 1547078400000);
        });
    });
    describe('SUBSTITUTE', function() {
        it('should throw #VALUE if occurence value is 0', function() {
            workbook.Sheets.Sheet1.A1 = { f: 'SUBSTITUTE("a","b","c",0)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.w, '#VALUE!');
            assert.equal(workbook.Sheets.Sheet1.A1.v, errorValues['#VALUE!']);
        });
        it('should throw #VALUE if occurence value is negative', function() {
            workbook.Sheets.Sheet1.A1 = { f: 'SUBSTITUTE("a","b","c",-5)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.w, '#VALUE!');
            assert.equal(workbook.Sheets.Sheet1.A1.v, errorValues['#VALUE!']);
        });
        it('should transform Jim to James', function() {
            workbook.Sheets.Sheet1.A1 = { f: 'SUBSTITUTE("Jim Alateras","im","ames")' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'James Alateras');
        });
        it('should transform nothing', function() {
            workbook.Sheets.Sheet1.A1 = { f: 'SUBSTITUTE("Jim Alateras","","ames")' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'Jim Alateras');
        });
        it('should equals empty string', function() {
            workbook.Sheets.Sheet1.A1 = { f: 'SUBSTITUTE("","im","ames")' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, '');
        });
        it('should equal Quarter 2, 2008', function() {
            workbook.Sheets.Sheet1.A1 = { f: 'SUBSTITUTE("Quarter 1, 2008","1","2", 1)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'Quarter 2, 2008');
        });
        it('should equal 07792 526879', function() {
            workbook.Sheets.Sheet1.A1 = { f: 'SUBSTITUTE("t:07792 526879","t:","")' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, '07792 526879');
        });
        it('should handle especial chars like dot', function() {
            workbook.Sheets.Sheet1.A1 = { f: 'SUBSTITUTE("my text","...","")' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'my text');
        });
    });
    describe('DAY', function () {
        xit('should return day of date value', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: new Date('2019-05-29'),
                w: '2019-05-29'
            };
            workbook.Sheets.Sheet1.A2 = { f: 'DAY(A1)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, 29);
            assert.equal(workbook.Sheets.Sheet1.A2.t, 'n');
        });
        it('should throw #VALUE error if applied to an invalid date', function () {
            workbook.Sheets.Sheet1.A1 = { v: 'AAA' };
            workbook.Sheets.Sheet1.A2 = { f: 'DAY(A1)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.t, 'e');
            assert.equal(workbook.Sheets.Sheet1.A2.w, '#VALUE!');
            assert.equal(workbook.Sheets.Sheet1.A2.v, errorValues['#VALUE!']);
        });
    });
    describe('MONTH', function () {
        it('should return month of date value', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: new Date('2019-05-29'),
                w: '2019-05-29'
            };
            workbook.Sheets.Sheet1.A2 = { f: 'MONTH(A1)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, 5);
            assert.equal(workbook.Sheets.Sheet1.A2.t, 'n');
        });
        it('should throw #VALUE error if applied to an invalid date', function () {
            workbook.Sheets.Sheet1.A1 = { v: 'AAA' };
            workbook.Sheets.Sheet1.A2 = { f: 'MONTH(A1)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.t, 'e');
            assert.equal(workbook.Sheets.Sheet1.A2.w, '#VALUE!');
            assert.equal(workbook.Sheets.Sheet1.A2.v, errorValues['#VALUE!']);
        });
    });
    describe('YEAR', function () {
        it('should return year of date value', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: new Date('2019-05-01'),
                w: '2019-05-01'
            };
            workbook.Sheets.Sheet1.A2 = { f: 'YEAR(A1)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, 2019);
            assert.equal(workbook.Sheets.Sheet1.A2.t, 'n');
        });
        it('should throw #VALUE error if applied to an invalid date', function () {
            workbook.Sheets.Sheet1.A1 = { v: 'AAA' };
            workbook.Sheets.Sheet1.A2 = { f: 'YEAR(A1)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.t, 'e');
            assert.equal(workbook.Sheets.Sheet1.A2.w, '#VALUE!');
            assert.equal(workbook.Sheets.Sheet1.A2.v, errorValues['#VALUE!']);
        });
    });

    describe('DATEDIF', function () {
        it('calcs DATEDIF("2000/1/1", "2001/1/1", "D")', function() {
            workbook.Sheets.Sheet1.A1.f = `DATEDIF('2000-01-01', '2001-01-01', 'D')`;
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 366);
        });

        it('calcs DATEDIF("2000/1/1", "2001/1/1", "M")', function() {
            workbook.Sheets.Sheet1.A1.f = `DATEDIF('2000-01-01', '2001-01-01', 'M')`;
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 12);
        });

        it('calcs DATEDIF("2000/1/1", "2001/1/1", "Y")', function() {
            workbook.Sheets.Sheet1.A1.f = `DATEDIF('2000-01-01', '2001-01-01', 'Y')`;
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 1);
        });

        it('should throw #VALUE error if applied to an invalid date', function () {
            workbook.Sheets.Sheet1.A1.f = `DATEDIF('NOT A DAY', '2001-01-01', 'Y')`;
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.w, "#VALUE!");
        });

        it('should return days between two values', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: new Date('2019-05-01'),
                w: '2019-05-01'
            };
            workbook.Sheets.Sheet1.A2 = {
                t: 'd',
                v: new Date('2019-06-01'),
                w: '2019-06-01'
            };
            workbook.Sheets.Sheet1.A3 = { f: 'DATEDIF(A1,A2,"D")' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A3.v, 31);
            assert.equal(workbook.Sheets.Sheet1.A3.t, 'n');
        });
    });

    describe('EOMONTH', function () {
        it('calcs EOMONTH("2000/1/1", 0)', function() {
            workbook.Sheets.Sheet1.A1.f = `EOMONTH("2000/1/1", 0)`;
            XLSX_CALC(workbook);
            assert.deepEqual(workbook.Sheets.Sheet1.A1.v, new Date("2000-01-31"));
        });

        it('calcs EOMONTH("2000/1/1", 1)', function() {
            workbook.Sheets.Sheet1.A1.f = `EOMONTH("2000/1/1", 1)`;
            XLSX_CALC(workbook);
            assert.deepEqual(workbook.Sheets.Sheet1.A1.v, new Date("2000-02-29"));
        });

        it('calcs EOMONTH("2000/12/1", 1)', function() {
            workbook.Sheets.Sheet1.A1.f = `EOMONTH("2000/12/1", 10)`;
            XLSX_CALC(workbook);
            assert.deepEqual(workbook.Sheets.Sheet1.A1.v, new Date("2001-10-31"));
        });

        it('should throw #VALUE error if applied to an invalid date', function () {
            workbook.Sheets.Sheet1.A1.f = `EOMONTH('NOT A DAY', 0)`;
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.w, "#VALUE!");
        });

        it('should return end of a month', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: new Date('2019-05-01'),
                w: '2019-05-01'
            };
            workbook.Sheets.Sheet1.A3 = { f: 'EOMONTH(A1,1)' };
            XLSX_CALC(workbook);
            assert.deepEqual(workbook.Sheets.Sheet1.A3.v, new Date('2019-06-30'));
        });
    });

    describe('RIGHT', function () {
        it('should return n last characters of a string value', function () {
            workbook.Sheets.Sheet1.A1.v = 'test value';
            workbook.Sheets.Sheet1.A1.t = 'n';
            workbook.Sheets.Sheet1.A2.f = 'RIGHT(A1, 2)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, 'ue');
            assert.equal(workbook.Sheets.Sheet1.A2.t, 's');
        });
        it('should return n last characters of a numeric value', function () {
            workbook.Sheets.Sheet1.A1.v = 2019;
            workbook.Sheets.Sheet1.A1.t = 'n';
            workbook.Sheets.Sheet1.A2.f = 'RIGHT(A1, 2)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, '19');
            assert.equal(workbook.Sheets.Sheet1.A2.t, 's');
        });
        it('should work in combination with YEAR formula', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: new Date('2019-05-01'),
                w: '2019-05-01'
            };
            workbook.Sheets.Sheet1.B1 = { f: 'YEAR(A1)' };
            workbook.Sheets.Sheet1.C1 = { f: 'RIGHT(B1, 2)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.C1.v, '19');
            assert.equal(workbook.Sheets.Sheet1.C1.t, 's');
        });
        it('should throw a #VALUE! error when used with an incorrect nb_car parameter', function () {
            workbook.Sheets.Sheet1.A1.v = 'test value';
            workbook.Sheets.Sheet1.A1.t = 'n';
            workbook.Sheets.Sheet1.A2.f = 'RIGHT(A1, "tutu")';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.t, 'e');
            assert.equal(workbook.Sheets.Sheet1.A2.w, '#VALUE!');
            assert.equal(workbook.Sheets.Sheet1.A2.v, errorValues['#VALUE!']);
        });
    });

    describe('LEFT', function () {
        it('should return n first characters of a string value', function () {
            workbook.Sheets.Sheet1.A1.v = 'test value';
            workbook.Sheets.Sheet1.A1.t = 'n';
            workbook.Sheets.Sheet1.A2.f = 'LEFT(A1, 2)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, 'te');
            assert.equal(workbook.Sheets.Sheet1.A2.t, 's');
        });
        it('should return n first characters of a numeric value', function () {
            workbook.Sheets.Sheet1.A1.v = 2019;
            workbook.Sheets.Sheet1.A1.t = 'n';
            workbook.Sheets.Sheet1.A2.f = 'LEFT(A1, 2)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, '20');
            assert.equal(workbook.Sheets.Sheet1.A2.t, 's');
        });
        it('should work in combination with YEAR formula', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'd',
                v: new Date('2019-05-01'),
                w: '2019-05-01'
            };
            workbook.Sheets.Sheet1.B1 = { f: 'YEAR(A1)' };
            workbook.Sheets.Sheet1.C1 = { f: 'LEFT(B1, 2)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.C1.v, '20');
            assert.equal(workbook.Sheets.Sheet1.C1.t, 's');
        });
        it('should throw a #VALUE! error when used with an incorrect nb_car parameter', function () {
            workbook.Sheets.Sheet1.A1.v = 'test value';
            workbook.Sheets.Sheet1.A1.t = 'n';
            workbook.Sheets.Sheet1.A2.f = 'LEFT(A1, "tutu")';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.t, 'e');
            assert.equal(workbook.Sheets.Sheet1.A2.w, '#VALUE!');
            assert.equal(workbook.Sheets.Sheet1.A2.v, errorValues['#VALUE!']);
        });
    });

    describe('IF', function() {
        it('should exec true', function() {
            workbook.Sheets.Sheet1.A1.f = 'IF(1<2,123,0)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 123);
        });
        it('should exec false', function() {
            workbook.Sheets.Sheet1.A1.f = 'IF(1>2,0,123)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 123);
        });
        it('should handle matrices of the same dimension', function() {
            workbook.Sheets.Sheet1.A1 = { v: true };
            workbook.Sheets.Sheet1.A2 = { v: false };
            workbook.Sheets.Sheet1.B1 = { v: false };
            workbook.Sheets.Sheet1.B2 = { v: true };
            workbook.Sheets.Sheet1.C1 = { v: 1 };
            workbook.Sheets.Sheet1.C2 = { v: 2 };
            workbook.Sheets.Sheet1.D1 = { v: 3 };
            workbook.Sheets.Sheet1.D2 = { v: 4 };
            workbook.Sheets.Sheet1.E1 = { v: 5 };
            workbook.Sheets.Sheet1.E2 = { v: 6 };
            workbook.Sheets.Sheet1.F1 = { v: 7 };
            workbook.Sheets.Sheet1.F2 = { v: 8 };
            workbook.Sheets.Sheet1.G1 = { f: 'IF(A1:B2,C1:D2,E1:F2)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.G1.v, 1);
            assert.equal(workbook.Sheets.Sheet1.G2.v, 6);
            assert.equal(workbook.Sheets.Sheet1.H1.v, 7);
            assert.equal(workbook.Sheets.Sheet1.H2.v, 4);
        });
        it('should handle matrices of different dimensions', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 2 };
            workbook.Sheets.Sheet1.C1 = { v: 3 };
            workbook.Sheets.Sheet1.D1 = { v: 4 };
            workbook.Sheets.Sheet1.D2 = { v: 5 };
            workbook.Sheets.Sheet1.D3 = { v: 6 };
            workbook.Sheets.Sheet1.E1 = { v: 7 };
            workbook.Sheets.Sheet1.E2 = { v: 8 };
            workbook.Sheets.Sheet1.E3 = { v: 9 };
            workbook.Sheets.Sheet1.F1 = { f: 'IF(A1:A3=A1:C1,D1:D3,E1:E3)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.F1.v, 4);
            assert.equal(workbook.Sheets.Sheet1.F2.v, 8);
            assert.equal(workbook.Sheets.Sheet1.F3.v, 9);
            assert.equal(workbook.Sheets.Sheet1.G1.v, 7);
            assert.equal(workbook.Sheets.Sheet1.G2.v, 5);
            assert.equal(workbook.Sheets.Sheet1.G3.v, 9);
            assert.equal(workbook.Sheets.Sheet1.H1.v, 7);
            assert.equal(workbook.Sheets.Sheet1.H2.v, 8);
            assert.equal(workbook.Sheets.Sheet1.H3.v, 6);
        });
        it('should handle missing references', function() {
            workbook.Sheets.Sheet1.A1 = { v: 1 };
            workbook.Sheets.Sheet1.A2 = { v: 2 };
            workbook.Sheets.Sheet1.A3 = { v: 3 };
            workbook.Sheets.Sheet1.B1 = { v: 1 };
            workbook.Sheets.Sheet1.B2 = { v: 3 };
            workbook.Sheets.Sheet1.C1 = { v: 4 };
            workbook.Sheets.Sheet1.C2 = { v: 5 };
            workbook.Sheets.Sheet1.C3 = { v: 6 };
            workbook.Sheets.Sheet1.D1 = { v: 7 };
            workbook.Sheets.Sheet1.D2 = { v: 8 };
            workbook.Sheets.Sheet1.D3 = { v: 9 };
            workbook.Sheets.Sheet1.E1 = { f: 'IF(A1:A3=B1:B2,C1:C3,D1:D3)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.E1.v, 4);
            assert.equal(workbook.Sheets.Sheet1.E2.v, 8);
            assert.equal(workbook.Sheets.Sheet1.E3.v, errorValues['#N/A']);
        });
    });
    it('calcs ref with space', function() {
        workbook.Sheets.Sheet1.A1.f = 'A2 ';
        workbook.Sheets.Sheet1.A2.v = 1979;
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets.Sheet1.A1.v, 1979);
    });
    it('calcs form with space after parentheses', function() {
        workbook.Sheets.Sheet1.A1.f = '(1) * 2';
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets.Sheet1.A1.v, 2);
    });
    it('calcs ref with $', function() {
        workbook.Sheets.Sheet1.A1.f = '$A$2 ';
        workbook.Sheets.Sheet1.A2.v = 1979;
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets.Sheet1.A1.v, 1979);
    });
    it('calcs ref chain', function() {
        workbook.Sheets.Sheet1.C4.f = 'A1';
        workbook.Sheets.Sheet1.A1.f = 'A2';
        workbook.Sheets.Sheet1.A2.v = 1979;
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets.Sheet1.C4.v, 1979);
    });
    it('calcs ref chain 2', function() {
        workbook.Sheets.Sheet1.C4.f = 'C3';
        workbook.Sheets.Sheet1.C3.f = 'C2';
        workbook.Sheets.Sheet1.C2.f = 'A2';
        workbook.Sheets.Sheet1.A2.f = 'A1';
        workbook.Sheets.Sheet1.A1.v = 1979;
        workbook.Sheets.Sheet1.C5.f = 'C3';
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets.Sheet1.C4.v, 1979);
    });
    it('throws a circular exception', function() {
        workbook.Sheets.Sheet1.C4.f = 'A1';
        workbook.Sheets.Sheet1.A1.f = 'C4';
        assert.throws(
            function() {
                XLSX_CALC(workbook);
            },
            /Circular ref/
        );
    });
    it('throws a function XPTO not found', function() {
        workbook.Sheets.Sheet1.A1.f = 'XPTO()';
        assert.throws(
            function() {
                XLSX_CALC(workbook);
            },
            /"Sheet1"!A1.*Function XPTO not found/
        );
    });
    it('handles error values', function () {
        workbook.Sheets.Sheet1.A1.f = '1/0';
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets.Sheet1.A1.v, errorValues['#DIV/0!']);
        assert.equal(workbook.Sheets.Sheet1.A1.w, '#DIV/0!');
        assert.equal(workbook.Sheets.Sheet1.A1.t, 'e');
    });
    it('propagates error values', function () {
        workbook.Sheets.Sheet1.A1 = {
            t: 'e',
            w: '#N/A',
            v: errorValues['#N/A']
        };
        workbook.Sheets.Sheet1.A2.f = '2*A1';
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets.Sheet1.A2.v, errorValues['#N/A']);
        assert.equal(workbook.Sheets.Sheet1.A2.w, '#N/A');
        assert.equal(workbook.Sheets.Sheet1.A2.t, 'e');

        workbook.Sheets.Sheet1.B1 = {
            f: '1/0'
        };
        workbook.Sheets.Sheet1.B2 = {
            f: '2*B1'
        };
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets.Sheet1.B2.v, errorValues['#DIV/0!']);
        assert.equal(workbook.Sheets.Sheet1.B2.w, '#DIV/0!');
        assert.equal(workbook.Sheets.Sheet1.B2.t, 'e');
    });
    describe('PTM', function() {
        it('calcs PMT(0.07/12, 24, 1000)', function() {
            workbook.Sheets.Sheet1.A1.f = 'PMT(0.07/12, 24, 1000)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, -44.77257910314528);
        });
        it('calcs PMT(0.07/12, 24, 1000,2000,0)', function() {
            workbook.Sheets.Sheet1.A1.f = 'PMT(0.07/12, 24, 1000,2000,0)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, -122.6510706427692);
        });

    });
    describe('COUNTA', function() {
        it('counts non empty cells', function() {
            workbook.Sheets.Sheet1.A1.f = 'COUNTA(B1:B3)';
            workbook.Sheets.Sheet1.B1 = {v:1};
            workbook.Sheets.Sheet1.B2 = {};
            workbook.Sheets.Sheet1.B3 = {v:1};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 2);
        });
    });
    describe('NORM.INV', function() {
        it('should call normsInv', function() {
            workbook.Sheets.Sheet1.A1.f = 'NORM.INV(0.05, -0.0015, 0.0175)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, -0.030284938471650775);
        });
    });
    describe('STDEV', function() {
        it('should calc STDEV', function() {
            workbook.Sheets.Sheet1.A1.f = 'STDEV(6.2,5,4.5,6,6,6.9,6.4,7.5)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 0.96204766736670300000);
        });
    });
    describe('AVERAGE', function() {
        it('should calc AVERAGE', function() {
            workbook.Sheets.Sheet1.A1.f = 'AVERAGE(1,2,3,4,5)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
        });
        it('should calc AVERAGE of range', function() {
            workbook.Sheets.Sheet1.A1 = {v: 0.1};
            workbook.Sheets.Sheet1.A2 = {v: 0.5};
            workbook.Sheets.Sheet1.A3 = {v: 0.2};
            workbook.Sheets.Sheet1.A4 = {v: 0.3};
            workbook.Sheets.Sheet1.A5 = {v: 0.2};
            workbook.Sheets.Sheet1.A6 = {v: 0.2};
            workbook.Sheets.Sheet1.A7 = {f: 'AVERAGE(A1:A6)'};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A7.v, 0.25);
        });
        it('should calc AVERAGE of empty cells as div by zero', function() {
            workbook.Sheets.Sheet1 = {};
            workbook.Sheets.Sheet1.B1 = {f: 'AVERAGE(A1:A6)'};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.B1.v, errorValues["#DIV/0!"]);
            assert.equal(workbook.Sheets.Sheet1.B1.w, '#DIV/0!');
            assert.equal(workbook.Sheets.Sheet1.B1.t, 'e');
        });
    });
    describe('IRR', function() {
        it('calcs IRR', function() {
            workbook.Sheets.Sheet1.A1.f = 'IRR(B1:B3)';
            workbook.Sheets.Sheet1.B1 = {v: -10.0};
            workbook.Sheets.Sheet1.B2 = {v:  -1.0};
            workbook.Sheets.Sheet1.B3 = {v:   2.9};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, -0.5091672986745834);
        });
        it('calcs IRR 2', function() {
            workbook.Sheets.Sheet1.A1.f = 'IRR(B1:B3)';
            workbook.Sheets.Sheet1.B1 = {v: -100.0};
            workbook.Sheets.Sheet1.B2 = {v:   10.0};
            workbook.Sheets.Sheet1.B3 = {v:  100000.0};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 30.672816276550293);
        });
    });
    describe('VAR.P', function() {
        it('calcs VAR.P', function() {
            workbook.Sheets.Sheet1.A1 = {v: 0.1};
            workbook.Sheets.Sheet1.A2 = {v: 0.5};
            workbook.Sheets.Sheet1.A3 = {v: 0.2};
            workbook.Sheets.Sheet1.A4 = {v: 0.3};
            workbook.Sheets.Sheet1.A5 = {v: 0.2};
            workbook.Sheets.Sheet1.A6 = {v: 0.2};
            workbook.Sheets.Sheet1.A7 = {f: 'VAR.P(A1:A6)'};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A7.v.toFixed(8), 0.01583333);
        });
        it('calls the VAR.P', function() {
            var x = XLSX_CALC.exec_fx('VAR.P', [0.1, 0.5, 0.2, 0.3, 0.2, 0.2]);
            assert.equal(x.toFixed(8), 0.01583333);
        });
    });
    describe('COVARIANCE.P', function() {
        it('computes COVARIANCE.P', function() {
            workbook.Sheets.Sheet1.A1 = {v: 0.01};
            workbook.Sheets.Sheet1.A2 = {v: 0.05};
            workbook.Sheets.Sheet1.A3 = {v: 0.02};
            workbook.Sheets.Sheet1.A4 = {v: 0.03};
            workbook.Sheets.Sheet1.A5 = {v: 0.02};
            workbook.Sheets.Sheet1.A6 = {v: 0.02};

            workbook.Sheets.Sheet1.B1 = {v: 0.1};
            workbook.Sheets.Sheet1.B2 = {v: 0.5};
            workbook.Sheets.Sheet1.B3 = {v: 0.2};
            workbook.Sheets.Sheet1.B4 = {v: 0.3};
            workbook.Sheets.Sheet1.B5 = {v: 0.2};
            workbook.Sheets.Sheet1.B6 = {v: 0.2};

            workbook.Sheets.Sheet1.A7 = {f: 'COVARIANCE.P(A1:A6,B1:B6)'};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A7.v.toFixed(8), 0.00158333);
        });
    });
    describe('#int_2_col_str', function() {
        it('should returns A', function() {
            assert.equal(XLSX_CALC.int_2_col_str(0), 'A');
        });
        it('should returns B', function() {
            assert.equal(XLSX_CALC.int_2_col_str(1), 'B');
        });
        it('should returns AZ', function() {
            assert.equal(XLSX_CALC.int_2_col_str(51), 'AZ');
        });
        it('should returns BA', function() {
            assert.equal(XLSX_CALC.int_2_col_str(52), 'BA');
        });
    });
    describe('EXP', function() {
        it('calculates EXP', function() {
            workbook.Sheets.Sheet1.A1.f = 'EXP(2)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 7.3890560989306495);
        });
    });
    describe('LN', function() {
        it('calculates LN of a number', function() {
            workbook.Sheets.Sheet1.A1.f = 'LN(EXP(2))';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 2);
        });
    });
    describe('ISBLANK', function() {
        it('calculates ISBLANK as false', function() {
            workbook.Sheets.Sheet1.A1.f = 'ISBLANK(B1)';
            workbook.Sheets.Sheet1.B1 = {v: ' '};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, false);
        });
        it('calculates ISBLANK as true', function() {
            workbook.Sheets.Sheet1.A1.f = 'ISBLANK(B1)';
            workbook.Sheets.Sheet1.B1 = {v: ''};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, true);
        });
        it('calculates ISBLANK as false for a ref to an error cell', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'e',
                w: '#N/A',
                v: errorValues['#N/A']
            };
            workbook.Sheets.Sheet1.B1 = {
                f: 'ISBLANK(A1)'
            };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.B1.v, false);
        });
    });
    describe('Sheet ref references', function() {
        it('calculates the sum of Sheet2!A1+Sheet2!A2', function() {
            workbook.Sheets.Sheet1.A1.f = 'Sheet2!A1+Sheet2!A2';
            workbook.Sheets.Sheet2 = { A1: {v:1}, A2: {v:2}};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
        });
        it('calculates the sum of Sheet2!A1:A2', function() {
            workbook.Sheets.Sheet1.A1.f = 'SUM(Sheet2!A1:A2)';
            workbook.Sheets.Sheet2 = { A1: {v:1}, A2: {v:2}};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
        });
        it('calculates the sum of Sheet2!A:B', function() {
            this.timeout(5000);
            workbook.Sheets.Sheet1.A1.f = 'SUM(Sheet2!A:B)';
            workbook.Sheets.Sheet2 = { A1: {v:1}, B1: {v:2}, A2: {v: 3}};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 6);
        });
    });
    describe('Cell type: A2.t = "s" or A2.t = "n"', function() {
        it('should set t = "s" for string values', function() {
            workbook.Sheets.Sheet1.A1 = { v: " some string " };
            workbook.Sheets.Sheet1.A2 = { f: "TRIM(A1)" };

            /* calculate */
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.t, 's');
            assert.equal(workbook.Sheets.Sheet1.A2.v, 'some string');
        });
        it('should set t = "n" for numeric values', function() {
            workbook.Sheets.Sheet1.A1 = { v: " some string " };
            workbook.Sheets.Sheet1.A2 = { f: "LEN(TRIM(A1))" };

            /* calculate */
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.t, 'n');
            assert.equal(workbook.Sheets.Sheet1.A2.v, 11);
        });
    });
    describe('raw function importer', function() {
        it('should sends the raw argument', function() {
            workbook.Sheets.Sheet1.A1 = { f: "MYRAWFN(A2,3-2,0)"};
            workbook.Sheets.Sheet1.A2 = { v: "VaLuE"};
            workbook.Sheets.Sheet1.B1 = { v: 1};
            XLSX_CALC.import_raw_functions({
                MYRAWFN: function(expr1, expr2, expr3) {
                    console.log(expr1.name); // Expression
                    console.log(expr1.args[0].name); // RefValue
                    console.log(expr1.args[0].str_expression); // A2
                    console.log(expr1.args[0].calc()); // VaLuE
                    return [expr1.args[0].str_expression, expr2.calc(), expr3.calc()];
                },
            });
            XLSX_CALC(workbook);
            assert.deepEqual(workbook.Sheets.Sheet1.A1.v, ['A2',1,0]);
        });
    });

    describe('ROW', function () {
       it('returns row in which the formula appears', function () {
           workbook.Sheets.Sheet1.A1.f = 'ROW()';
           XLSX_CALC(workbook);
           assert.equal(workbook.Sheets.Sheet1.A1.v, 1);
       });
        it('returns row of the reference', function () {
            workbook.Sheets.Sheet1.A1.f = 'ROW(A2)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 2);
        });
        it('returns row as an array if the reference is an array', function () {
            workbook.Sheets.Sheet1.A1.f = 'ROW(A1:A3)';
            XLSX_CALC(workbook);
            assert.deepEqual(workbook.Sheets.Sheet1.A1.v, [1, 2, 3]);
        });
    });

    describe('COLUMN', function () {
        it('returns column in which the formula appears', function () {
            workbook.Sheets.Sheet1.A1.f = 'COLUMN()';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 1);
        });
        it('returns column of the reference', function () {
            workbook.Sheets.Sheet1.A1.f = 'COLUMN(B1)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 2);
        });
        it('returns column as an array if the reference is an array', function () {
            workbook.Sheets.Sheet1.A1.f = 'COLUMN(A1:C1)';
            XLSX_CALC(workbook);
            assert.deepEqual(workbook.Sheets.Sheet1.A1.v, [1, 2, 3]);
        });
    });

    describe('ISERROR', function () {
        it('returns true if in error', function () {
            workbook.Sheets.Sheet1.A1 = { f: "0/0" };
            workbook.Sheets.Sheet1.A2 = { f: "ISERROR(A1)" };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, true);
        });
        it('returns false if in not in error', function () {
            workbook.Sheets.Sheet1.A1 = { f: "2*3" };
            workbook.Sheets.Sheet1.A2 = { f: "ISERROR(A1)" };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, false);
        });
    });

    describe('IFERROR', function() {
        it('returns the string Error', function() {
            workbook.Sheets.Sheet1.A1 = { f: "IFERROR(A2,\"Error\")"};
            workbook.Sheets.Sheet1.A2 = { f: "0/0"};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'Error');
        });
        it('returns the string Error when res is Infinity', function() {
            workbook.Sheets.Sheet1.A1 = { f: "IFERROR(1/0,\"Error\")"};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'Error');
        });
        it('returns the string Error when res is -Infinity', function() {
            workbook.Sheets.Sheet1.A1 = { f: "IFERROR(-1/0,\"Error\")"};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'Error');
        });
        it('returns the string boston', function() {
            workbook.Sheets.Sheet1.A1 = { f: "IFERROR(A2,\"Error\")"};
            workbook.Sheets.Sheet1.A2 = { v: "boston"};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'boston');
        });
        it('returns the string Error when VLOOKUP fail', function() {
            workbook.Sheets.Sheet1.A1 = { f: "IFERROR(A2,\"Error\")"};
            workbook.Sheets.Sheet1.A2 = { f: "VLOOKUP(\"void\",\"A3:B7\",2)"};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'Error');
        });
        it('returns the string Error when in reference to an error cell', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'e',
                w: '#N/A',
                v: errorValues['#N/A']
            };
            workbook.Sheets.Sheet1.A2 = { f: "IFERROR(A1, \"Error\")" };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, 'Error');
            assert.equal(workbook.Sheets.Sheet1.A2.t, 's');
        });
        it('returns the string Error when using an error cell in a formula', function () {
            workbook.Sheets.Sheet1.A1 = {
                t: 'e',
                w: '#N/A',
                v: errorValues['#N/A']
            };
            workbook.Sheets.Sheet1.A2 = { f: "IFERROR(2*A1, \"Error\")" };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A2.v, 'Error');
            assert.equal(workbook.Sheets.Sheet1.A2.t, 's');
        });
    });

    describe('HLOOKUP', function () {
        it('searches for a key in the top row of a matrix and returns the value in the same column at the specified return_index row', function () {
            workbook.Sheets.Sheet1.A1 = { v: 'Axles' };
            workbook.Sheets.Sheet1.B1 = { v: 'Bearings' };
            workbook.Sheets.Sheet1.C1 = { v: 'Bolts' };

            workbook.Sheets.Sheet1.A2 = { v: 4 };
            workbook.Sheets.Sheet1.B2 = { v: 4 };
            workbook.Sheets.Sheet1.C2 = { v: 9 };
            workbook.Sheets.Sheet1.A3 = { v: 5 };
            workbook.Sheets.Sheet1.B3 = { v: 7 };
            workbook.Sheets.Sheet1.C3 = { v: 10 };
            workbook.Sheets.Sheet1.A4 = { v: 6 };
            workbook.Sheets.Sheet1.B4 = { v: 8 };
            workbook.Sheets.Sheet1.C4 = { v: 11 };

            workbook.Sheets.Sheet1.D1 = { f: "HLOOKUP(\"Bearings\", A1:C4, 3, FALSE)" };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.D1.v, 7);
        });
        it('returns #N/A error when used with an empty needle', function () {
            workbook.Sheets.Sheet1 = {
                A1: { v: "a" },
                A2: { v: "b" },
                A3: { v: "c" },
                B1: { v: "1" },
                B2: { v: "2" },
                B3: { v: "3" },
                C1: { f: 'HLOOKUP(C2, A1:B3, 2, FALSE)'}
            };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.C1.t, "e");
            assert.equal(workbook.Sheets.Sheet1.C1.v, 42);
            assert.equal(workbook.Sheets.Sheet1.C1.w, "#N/A");
        });
    });

    describe('VLOOKUP', function() {
        it('exact match', function() {
            workbook.Sheets.Sheet1 = {
                A1: { v: 'A' },
                A2: { v: 'B' },
                A3: { v: 'C' },
                A4: { v: 'D' },
                A5: { v: 'E' },
                B1: { v: 1996 },
                B2: { v: 1997 },
                B3: { v: 1999 },
                B4: { v: 1995 },
                B5: { v: 1992 },
                C1: { v: 5 },
                C2: { v: 4 },
                C3: { v: 1 },
                C4: { v: 2 },
                C5: { v: 3 },
                D1: { v: 'D' },
                D2: { f: 'VLOOKUP("D", A1:C5, 2, FALSE)' },
            }
            XLSX_CALC(workbook)
            assert.equal(workbook.Sheets.Sheet1.D2.v, 1995)
        })
        it('approximate match', function() {
            workbook.Sheets.Sheet1 = {
                A1: { v: 171900 },
                A2: { v: 93500 },
                A3: { v: 151200 },
                A4: { v: 119850 },
                A5: { v: 89450 },
                A6: { v: 124500 },
                A7: { v: 131100 },
                A8: { v: 201500 },
                B1: { f: 'VLOOKUP(A1, C1:D6, 2, TRUE)' },
                B2: { f: 'VLOOKUP(A2, C1:D6, 2, TRUE)' },
                B3: { f: 'VLOOKUP(A3, C1:D6, 2, TRUE)' },
                B4: { f: 'VLOOKUP(A4, C1:D6, 2, TRUE)' },
                B5: { f: 'VLOOKUP(A5, C1:D6, 2, TRUE)' },
                B6: { f: 'VLOOKUP(A6, C1:D6, 2, TRUE)' },
                B7: { f: 'VLOOKUP(A7, C1:D6, 2, TRUE)' },
                B8: { f: 'VLOOKUP(A8, C1:D6, 2, TRUE)' },
                C1: { v: 50000 },
                C2: { v: 75000 },
                C3: { v: 100000 },
                C4: { v: 125000 },
                C5: { v: 175000 },
                C6: { v: 200000 },
                D1: { v: 1 },
                D2: { v: 2 },
                D3: { v: 3 },
                D4: { v: 4 },
                D5: { v: 5 },
                D6: { v: 6 },
            }
            XLSX_CALC(workbook)
            assert.equal(workbook.Sheets.Sheet1.B1.v, 4)
            assert.equal(workbook.Sheets.Sheet1.B2.v, 2)
            assert.equal(workbook.Sheets.Sheet1.B3.v, 4)
            assert.equal(workbook.Sheets.Sheet1.B4.v, 3)
            assert.equal(workbook.Sheets.Sheet1.B5.v, 2)
            assert.equal(workbook.Sheets.Sheet1.B6.v, 3)
            assert.equal(workbook.Sheets.Sheet1.B7.v, 4)
            assert.equal(workbook.Sheets.Sheet1.B8.v, 6)
        })
    })

    describe('INDEX', function () {
        it('returns the value of an element in a matrix, selected by the row and column number indexes', function () {
            workbook.Sheets.Sheet1.A1 = { v: 'Data' };
            workbook.Sheets.Sheet1.B1 = { v: 'Data' };

            workbook.Sheets.Sheet1.A2 = { v: 'Apples' };
            workbook.Sheets.Sheet1.B2 = { v: 'Lemons' };
            workbook.Sheets.Sheet1.A3 = { v: 'Bananas' };
            workbook.Sheets.Sheet1.B3 = { v: 'Pears' };

            workbook.Sheets.Sheet1.C1 = { f: "INDEX(A2:B3, 2, 2)" };
            workbook.Sheets.Sheet1.C2 = { f: "INDEX(A2:B3, 2, 1)" };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.C1.v, "Pears");
            assert.equal(workbook.Sheets.Sheet1.C2.v, "Bananas");
        });
    });

    describe('MATCH_GREATER_THAN_OR_EQUAL', function () {
        it('return position of element in range (row or column)', function () {
            workbook.Sheets.Sheet1.A1 = { v: 100 };
            workbook.Sheets.Sheet1.A2 = { v: 80 };
            workbook.Sheets.Sheet1.A3 = { v: 60 };
            workbook.Sheets.Sheet1.A4 = { v: 0 };

            workbook.Sheets.Sheet1.B1 = { v: 100 };
            workbook.Sheets.Sheet1.C1 = { v: 80 };
            workbook.Sheets.Sheet1.D1 = { v: 60 };
            workbook.Sheets.Sheet1.E1 = { v: 0 };

            workbook.Sheets.Sheet1.B2 = { v: 80 };
            workbook.Sheets.Sheet1.B3 = { v: 15 };
            workbook.Sheets.Sheet1.C2 = { v: 75 };
            workbook.Sheets.Sheet1.C3 = { v: 20 };
            workbook.Sheets.Sheet1.D2 = { v: 100 };
            workbook.Sheets.Sheet1.D3 = { v: 12 };
            workbook.Sheets.Sheet1.B4 = { f: 'MATCH(B2, A1:A4, -1)' };
            workbook.Sheets.Sheet1.B5 = { f: 'MATCH(B3, B1:F1, -1)' };
            workbook.Sheets.Sheet1.C4 = { f: 'MATCH(C2, B1:F1, -1)' };
            workbook.Sheets.Sheet1.C5 = { f: 'MATCH(C3, B1:F1, -1)' };
            workbook.Sheets.Sheet1.D4 = { f: 'MATCH(D2, B1:F1, -1)' };
            workbook.Sheets.Sheet1.D5 = { f: 'MATCH(D3, B1:F1, -1)' };
            workbook.Sheets.Sheet1.C6 = { f: 'MATCH(C2, A1:A4, -1)' };
            workbook.Sheets.Sheet1.C7 = { f: 'MATCH(C3, A1:A4, -1)' };
            workbook.Sheets.Sheet1.D6 = { f: 'MATCH(D2, A1:A4, -1)' };
            workbook.Sheets.Sheet1.D7 = { f: 'MATCH(D3, A1:A4, -1)' };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.B4.v, 2);
            assert.equal(workbook.Sheets.Sheet1.B5.v, 3);
            assert.equal(workbook.Sheets.Sheet1.C4.v, 2);
            assert.equal(workbook.Sheets.Sheet1.C5.v, 3);
            assert.equal(workbook.Sheets.Sheet1.D4.v, 1);
            assert.equal(workbook.Sheets.Sheet1.D5.v, 3);
            assert.equal(workbook.Sheets.Sheet1.C6.v, 2);
            assert.equal(workbook.Sheets.Sheet1.C7.v, 3);
            assert.equal(workbook.Sheets.Sheet1.D6.v, 1);
            assert.equal(workbook.Sheets.Sheet1.D7.v, 3);
        });
    });

    describe('MATCH', function () {
        it('returns position of element in range (row or column)', function () {
            workbook.Sheets.Sheet1.A1 = { v: 'Apple' };
            workbook.Sheets.Sheet1.A2 = { v: 'Raspberry' };
            workbook.Sheets.Sheet1.A3 = { v: 'Carambola' };
            workbook.Sheets.Sheet1.A4 = { v: 'Pear' };

            workbook.Sheets.Sheet1.B1 = { v: 'Cantaloupe' };
            workbook.Sheets.Sheet1.C1 = { v: 'Longan' };
            workbook.Sheets.Sheet1.D1 = { v: 'Lime' };
            workbook.Sheets.Sheet1.E1 = { v: 'Carambola' };
            workbook.Sheets.Sheet1.F1 = { v: 'Grape' };

            workbook.Sheets.Sheet1.B2 = { v: 'Carambola' };
            workbook.Sheets.Sheet1.B3 = { f: "MATCH(B2, A1:A4, 0)" };
            workbook.Sheets.Sheet1.B4 = { f: "MATCH(B2, A1:F1, 0)" };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.B3.v, 3);
            assert.equal(workbook.Sheets.Sheet1.B4.v, 5);
        });
        it('should show "#N/A" error when a multi-dimensional array is passed', function () {
            workbook.Sheets.Sheet1.B3 = { v: 'Carambola' };
            workbook.Sheets.Sheet1.A3 = { f: "MATCH(B3, A1:B2, 0)" };
            XLSX_CALC(workbook);

            assert.equal(workbook.Sheets.Sheet1.A3.t, 'e');
            assert.equal(workbook.Sheets.Sheet1.A3.w, '#N/A');
            assert.equal(workbook.Sheets.Sheet1.A3.v, errorValues['#N/A']);
        });
    });

    describe('SUMPRODUCT', function () {
        it('Multiplies corresponding components in the given arrays, and returns the sum of those products', function () {
            workbook.Sheets.Sheet1 = {
                A1: { v: 'Array 1' },
                A2: { v: 3 },
                A3: { v: 8 },
                A4: { v: 1 },
                B2: { v: 4 },
                B3: { v: 6 },
                B4: { v: 9 },
                D1: { v: 'Array 2' },
                D2: { v: 2 },
                D3: { v: 6 },
                D4: { v: 5 },
                E2: { v: 7 },
                E3: { v: 7 },
                E4: { v: 3 },
                C1: { f: "SUMPRODUCT(A2:B4, D2:E4)" }
            };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.C1.v, 156);
        });
        it('should handle empty values in the given arrays', function () {
            workbook.Sheets.Sheet1 = {
                A1: { v: 'Array 1' },
                A2: { v: 3 },
                A4: { v: 8 },
                D1: { v: 'Array 2' },
                D2: { v: 2 },
                D4: { v: 6 },
                C1: { f: "SUMPRODUCT(A2:A4, D2:D4)" }
            };

            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.C1.v, 54);
        });
        it('should handle no values for the given arrays', function () {
            workbook.Sheets.Sheet1 = {
                A1: { v: 'Array 1' },
                A2: { v: 3 },
                A4: { v: 8 },
                D1: { v: 'Array 2' },
                D2: { v: 2 },
                D4: { v: 6 },
                C1: { f: "SUMPRODUCT(,)" }
            };

            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.C1.w, '#VALUE!');
        });
        it('shows "#VALUE!" error value if the array arguments dont have the same dimensions', function () {
            workbook.Sheets.Sheet1 = {
                A1: { v: 'Array 1' },
                A2: { v: 3 },
                A3: { v: 8 },
                D1: { v: 'Array 2' },
                D2: { v: 2 },
                D3: { v: 4 },
                D4: { v: 6 },
                C1: { f: "SUMPRODUCT(A2:A3, D2:D4)" }
            };
            XLSX_CALC(workbook);

            assert.equal(workbook.Sheets.Sheet1.C1.t, 'e');
            assert.equal(workbook.Sheets.Sheet1.C1.w, '#VALUE!');
            assert.equal(workbook.Sheets.Sheet1.C1.v, errorValues['#VALUE!']);
        });
        it('treats array entries that are not numeric as if they were zeros', function () {
            workbook.Sheets.Sheet1 = {
                A1: { v: 'Array 1' },
                A2: { v: 3 },
                A3: { v: 8 },
                D1: { v: 'Array 2' },
                D2: { v: 2 },
                D3: { v: 6 },
                C1: { f: "SUMPRODUCT(A1:A3, D1:D3)" }
            };

            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.C1.v, 54);
        });
        it('should transmit error if any cell in range is in error', function () {
            workbook.Sheets.Sheet1 = {
                A1: { t: "e", v: 42, w: "#N/A" },
                A2: { t: "n", v: 0.5 },
                A3: { t: "n", v: 0 },
                B1: { t: "n", v: 0.1, w: "10%" },
                B2: { t: "n", v: 0.2, w: "20%" },
                B3: { t: "n", v: 0, w: "0%" },
                C1: { f: "SUMPRODUCT(A1:A3, B1:B3)" }
            };

            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.C1.t, "e");
            assert.equal(workbook.Sheets.Sheet1.C1.v, 42);
            assert.equal(workbook.Sheets.Sheet1.C1.w, "#N/A");
        });
        it('should transmit error if any cell in range is in error even in a range with empty cells', function () {
            workbook.Sheets.Sheet1 = {
                A16: { t: "e", v: 42, w: "#N/A" },
                A17: { t: "n", v: 0.5, w: "0.5" },
                A18: { t: "n", v: 0, w: " -    " },
                A21: { t: "n", v: 1, w: "0.5" },
                A22: { t: "n", v: 1, w: " 1.00    " },
                A24: { t: "n", v: 1, w: "0.5" },
                A26: { t: "n", v: 1, w: "1" },
                A28: { t: "n", v: 1, w: "0" },
                B16: { t: "n", v: 0.1, w: "10%" },
                B17: { t: "n", v: 0.2, w: "20%" },
                B18: { t: "n", v: 0, w: "0%" },
                B21: { t: "n", v: 0.15, w: "15%" },
                B22: { t: "n", v: 0, w: "0%" },
                B24: { t: "n", v: 0.05, w: "5%" },
                B26: { t: "n", v: 0.1, w: "10%" },
                B28: { t: "n", v: 0.1, w: "10%" },
                C1: { f: "SUMPRODUCT(A16:A29, B16:B29)" }
            };

            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.C1.t, "e");
            assert.equal(workbook.Sheets.Sheet1.C1.v, 42);
            assert.equal(workbook.Sheets.Sheet1.C1.w, "#N/A");
        });
    });

    describe('AND', () => {
        it('evaluates false', () => {
            workbook.Sheets.Sheet1 = {
                A1: { f: 'IF(AND(1,0),"err","ok")' }
            };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'ok');
        });
        it('evaluates true', () => {
            workbook.Sheets.Sheet1 = {
                A1: { f: 'IF(AND(1,1),"ok","err")' }
            };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'ok');
        });
    });

    describe('localizeFunctions', () => {
        it('makes an alias to CONCATENATE', () => {
            workbook.Sheets.Sheet1 = {
                A1: { f: 'CONCAT("Hello"," ","world")' }
            };
            XLSX_CALC.localizeFunctions({
                CONCAT: 'CONCATENATE'
            });
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'Hello world');
        });
    });

    describe('IFS', () => {
        it("for every pair of values, returns the first value that resolves to true", () => {
            workbook.Sheets.Sheet1 = {
                A1: { f: 'IFS(0, "a", 0, "b", 1, "c")' }
            };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 'c');
        });
    });

    describe('OR', () => {
        it('evaluates an OR(0,0) as false', () => {
            workbook.Sheets.Sheet1 = {
                A1: { f: 'OR(0,0)' }
            };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, false);
        });
        it('evaluates an OR(1,0) as true', () => {
            workbook.Sheets.Sheet1 = {
                A1: { f: 'OR(1,0)' }
            };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, true);
        });
        it('evaluates an OR(0,1) as true', () => {
            workbook.Sheets.Sheet1 = {
                A1: { f: 'OR(0,1)' }
            };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, true);
        });
    });

    describe('CHOOSE', () => {
        it('evaluates an CHOOSE(2,"a","b","c") as true', () => {
            workbook.Sheets.Sheet1 = {
                A1: { f: 'CHOOSE(2,"a","b","c")' }
            };
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, "b");
        });
    });

    describe('CEILING', () => {
        it('should ceiling to 150', () => {
            workbook.Sheets.Sheet1.A1.f = 'CEILING(141,10)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 150);
        });
    });

    // describe('HELLO', function() {
    //     it('says: Hello, World!', function() {
    //         workbook.Sheets.Sheet1.A1.f = 'HELLO("World")';
    //         XLSX_CALC(workbook);
    //         assert.equal(workbook.Sheets.Sheet1.A1.v, "Hello, World!");
    //     });
    // });
});
