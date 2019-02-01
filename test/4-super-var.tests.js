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
    it('cria um erro inteligivel quando a variavel nao for setada', () => {
        workbook.Sheets.Sheet1.A1.f = '1+SUM([a2],([a3]+[a]))';
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a]', 1);
        try {
            calculator.execute();
            throw new Error('Where is the error?');
        } catch(e) {
            assert.equal(e.message, 'Undefined [a3]');
        }
    });
    it('cria um erro inteligivel quando a variavel nao for setada e ela estiver com o sinal de menos na frente', () => {
        workbook.Sheets.Sheet1.A1.f = '[a]-[a3]';
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a]', 1);
        try {
            calculator.execute();
            throw new Error('Where is the error?');
        } catch(e) {
            assert.equal(e.message, 'Undefined [a3]');
        }
    });
    it('gets all variables setted', () => {
        workbook.Sheets.Sheet1 = {A1: {f: '[a1]-(([a3]))'}};
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a1]', 1);
        calculator.setVar('[a3]', 3);
        calculator.execute();
        let vars = calculator.getVars();
        assert.equal(workbook.Sheets.Sheet1.A1.v, -2);
        assert.equal(vars['[a1]'], 1);
        assert.equal(vars['[a3]'], 3);
    });
    
    it('sets named cell', () => {
        workbook.Workbook = {
            Names: [{Name: 'XPTO', Ref: 'Sheet1!A2'}]
        };
        workbook.Sheets.Sheet1 = {
            A1: { f: 'XPTO+B1' }, A2: { v: 3 }, B1: { v: 2 }
        };
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.execute();
        assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
    });
    
    it('sets named cell range', () => {
        workbook.Workbook = {
            Names: [{Name: 'XPTO', Ref: 'Sheet1!A2:A3'}]
        };
        workbook.Sheets.Sheet1 = {
            A1: { f: 'SUM(XPTO)' }, A2: { v: 3 }, A3: { v: 2 }
        };
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.execute();
        assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
    });
});
