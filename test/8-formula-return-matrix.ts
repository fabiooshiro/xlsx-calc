import * as assert from 'assert';

import XLSX_CALC from "../src";

describe('formula that returns a matrix', () => {

    it('should set a matrix 3x3', () => {
        let workbook: any = {
            Sheets: {
                Sheet1: {
                    E3: {f:  "SET_MATRIX()"},
                }
            }
        };
        XLSX_CALC.set_fx('SET_MATRIX', () => {
            return [
                ['aa', 'bb', 'cc'],
                ['aaa', 'bbb', 'ccc'],
                ['aaaa', 'bbbb', 'cccc'],
            ]
        })
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

    it('should replace empty blocks', () => {
        let workbook: any = {
            Sheets: {
                Sheet1: {
                    E3: {f:  "SET_MATRIX()"},
                }
            }
        };
        XLSX_CALC.set_fx('SET_MATRIX', () => {
            return [
                ['aa', 'bb', 'cc'],
                ['aaa', 'bbb', 'ccc'],
                ['aaaa', 'bbbb', ,],
            ]
        })
        XLSX_CALC(workbook);

        assert.equal(workbook.Sheets.Sheet1.E3.v, 'aa');
        assert.equal(workbook.Sheets.Sheet1.E4.v, 'aaa');
        assert.equal(workbook.Sheets.Sheet1.E5.v, 'aaaa');
        assert.equal(workbook.Sheets.Sheet1.F3.v, 'bb');
        assert.equal(workbook.Sheets.Sheet1.F4.v, 'bbb');
        assert.equal(workbook.Sheets.Sheet1.F5.v, 'bbbb');
        assert.equal(workbook.Sheets.Sheet1.G3.v, 'cc');
        assert.equal(workbook.Sheets.Sheet1.G4.v, 'ccc');
        assert.equal(workbook.Sheets.Sheet1.G5.v, undefined);
    });

    
})