"use strict";

const XLSX_CALC = require("../");
const assert = require('assert');

describe('trocar variavel', () => {
    let workbook;
    function createWorkbook() {
        return {
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
    }
    beforeEach(() => {
        workbook = createWorkbook();
    });
    it('troca o valor da variavel', (done) => {
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a]', 1);
        calculator.execute().then(() => {
            assert.equal(workbook.Sheets.Sheet1.A1.v, 2);
            done();
        }).catch(done);
    });
    it('troca o valor da variavel duas vezes', (done) => {
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a]', 1);
        calculator.setVar('[a]', 2);
        calculator.execute().then(() => {
            assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
            done();
        }).catch(done);
    });
    it('calcula normal', (done) => {
        let promises = [];
        for (let i = 0; i < 2001; i++) {
            promises.push(new Promise((resolve, reject) => {
                let workbook = createWorkbook();
                workbook.Sheets.Sheet1.A1.f = '1+4';
                workbook.Sheets.Sheet1.B1.f = '1+6';
                XLSX_CALC(workbook).then(() => {
                    assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
                    assert.equal(workbook.Sheets.Sheet1.B1.v, 7);
                    workbook.Sheets.Sheet1.A1.f = '1+2';
                    workbook.Sheets.Sheet1.B1.f = '3+6';
                    return XLSX_CALC(workbook);
                }).then(() => {
                    assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
                    assert.equal(workbook.Sheets.Sheet1.B1.v, 9);
                    resolve();
                }).catch(done);
            }));
        }
        Promise.all(promises).then(() => {
            done();
        });
    });
    it('troca o valor da variavel duas vezes nas duas celulas', (done) => {
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a]', 4);
        calculator.execute().then(() => {
            assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
            assert.equal(workbook.Sheets.Sheet1.B1.v, 5);
            calculator.setVar('[a]', 2);
            return calculator.execute();
        }).then(() => {
            assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
            assert.equal(workbook.Sheets.Sheet1.B1.v, 3);
            done();
        }).catch(done);
    });
    it('troca o valor da variavel dentro de outras expressoes', (done) => {
        workbook.Sheets.Sheet1.A1.f = '1+[a]+([a2]+[a3])';
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a]', 1);
        calculator.setVar('[a2]', 1);
        calculator.setVar('[a3]', 2);
        calculator.execute().then(() => {
            assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
            done();
        }).catch(done);
    });
    it('troca o valor da variavel dentro de argumentos de funcoes', (done) => {
        workbook.Sheets.Sheet1.A1.f = '1+SUM([a2],[a3])';
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a]', 1);
        calculator.setVar('[a2]', 2);
        calculator.setVar('[a3]', 3);
        calculator.execute().then(() => {
            assert.equal(workbook.Sheets.Sheet1.A1.v, 6);
            done();
        }).catch(done);
    });
    it('troca o valor da variavel dentro de argumentos de funcoes', (done) => {
        workbook.Sheets.Sheet1.A1.f = '1+SUM([a2],([a3]+[a]))';
        let calculator = XLSX_CALC.calculator(workbook);
        calculator.setVar('[a]', 1);
        calculator.setVar('[a2]', 2);
        calculator.setVar('[a3]', 3);
        calculator.execute().then(() => {
            assert.equal(workbook.Sheets.Sheet1.A1.v, 7);
            done();
        }).catch(done);
    });
});
