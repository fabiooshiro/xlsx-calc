"use strict";

const XLSX_CALC = require("../");
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
    it('troca o valor da variavel duas vezes nas duas celulas', () => {
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a]', 1);
        calculator.setVar('[a]', 2);
        calculator.execute();
        assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
        assert.equal(workbook.Sheets.Sheet1.B1.v, 3);
    });
});
