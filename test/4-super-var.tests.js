"use strict";

const XLSX_CALC = require("../lib/xlsx-calc");
const assert = require('assert');

describe('trocar variavel', () => {
    let workbook;
    beforeEach(() => {
        workbook = {
            Sheets: {
                Sheet1: {
                    A1: {
                        f: '1+[a]'
                    },
                    B1: {
                        f: '1+[a]'
                    }
                }
            }
        }; 
    });
    it('troca o valor da variavel', () => {
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a]', 1);
        calculator.execute();
        assert.equal(workbook.Sheets.Sheet1.A1.v, 2);
    });
    it('troca o valor da variavel duas vezes', () => {
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a]', 1);
        calculator.setVar('[a]', 2);
        calculator.execute();
        assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
    });
    xit('calcula normal', () => {
        for (let i = 0; i < 10000; i++) {
            workbook.Sheets.Sheet1.A1.f = '1+4';
            workbook.Sheets.Sheet1.B1.f = '1+4';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
            assert.equal(workbook.Sheets.Sheet1.B1.v, 5);
            workbook.Sheets.Sheet1.A1.f = '1+2';
            workbook.Sheets.Sheet1.B1.f = '1+2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
            assert.equal(workbook.Sheets.Sheet1.B1.v, 3);
        }
    });
    it('troca o valor da variavel duas vezes nas duas celulas', () => {
        let calculator = XLSX_CALC.calculator(workbook);
        for (let i = 0; i < 1000; i++) {
            calculator.setVar('[a]', 4);
            calculator.execute();
            assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
            assert.equal(workbook.Sheets.Sheet1.B1.v, 5);
            calculator.setVar('[a]', 2);
            calculator.execute();
            assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
            assert.equal(workbook.Sheets.Sheet1.B1.v, 3);
        }
    });
    it('troca o valor da variavel dentro de outras expressoes', () => {
        workbook.Sheets.Sheet1.A1.f = '1+[a]+([a2]+[a3])';
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a]', 1);
        calculator.setVar('[a2]', 1);
        calculator.setVar('[a3]', 2);
        calculator.execute();
        assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
    });
    it('troca o valor da variavel dentro de argumentos de funcoes', () => {
        workbook.Sheets.Sheet1.A1.f = '1+SUM([a2],[a3])';
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a]', 1);
        calculator.setVar('[a2]', 2);
        calculator.setVar('[a3]', 3);
        calculator.execute();
        assert.equal(workbook.Sheets.Sheet1.A1.v, 6);
    });
    it('troca o valor da variavel dentro de argumentos de funcoes', () => {
        workbook.Sheets.Sheet1.A1.f = '1+SUM([a2],([a3]+[a]))';
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a]', 1);
        calculator.setVar('[a2]', 2);
        calculator.setVar('[a3]', 3);
        calculator.execute();
        assert.equal(workbook.Sheets.Sheet1.A1.v, 7);
    });
});
