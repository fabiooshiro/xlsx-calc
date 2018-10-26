"use strict";

var XLSX_CALC = require('../lib/xlsx-calc');
var XLSX = require('xlsx');
var assert = require('assert');

describe('Bugs', function() {
    var workbook;
    beforeEach(function() {
        workbook = {
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
    });
    it('should consider the end of string', function() {
        workbook.Sheets.Sheet1.A1.f = 'IF($C$3<=0,"Tempo de Investimento Invalido",IF($C$3<=24,"x","y"))';
        workbook.Sheets.Sheet1.C3 = { v: 24 };
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets.Sheet1.A1.v, 'x');
    });
    it('should eval 10%', function() {
        workbook.Sheets.Sheet1.A1.f = '(B3*10%)/12';
        workbook.Sheets.Sheet1.B3 = { v: 120 };
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets.Sheet1.A1.v, 1);
    });
    it('should works', function() {
        workbook.Sheets.Sheet1.A1.f = '-1-2';
        workbook.Sheets.Sheet1.B1 = {f: '4^5'};
        workbook.Sheets.Sheet1.C1 = {v: 33};
        workbook.Sheets.Sheet1.A2 = {f: 'SUM(A1:C1)'};
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets.Sheet1.A2.v, 1054);
    });
    it('should ignore spaces before (', function() {
        workbook.Sheets.Sheet1.A1.f = '- 1 - (1+1)';
        workbook.Sheets.Sheet1.B1 = {f: '4^5'};
        workbook.Sheets.Sheet1.C1 = {v: 33};
        workbook.Sheets.Sheet1.A2 = {f: 'SUM(A1:C1)'};
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets.Sheet1.A2.v, 1054);
    });
    it('returns the correct string for column', function() {
        assert.equal(XLSX_CALC.int_2_col_str(130), 'EA');
    });
    it('resolves the bug of quoted sheet names', function() {
        workbook = XLSX.readFile('test/abc.xlsx');
        workbook.Sheets['B C'].A1.v = 2000;
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets['A'].A1.v, 2000);
    });
    
    it('some bug?', () => {
        let workbook = {
            Sheets: {
                Sample: {
                    A1: { f: 'CONCATENATE("I"," told"," you ","10"," times",".")' },
                    A2: { f: 'CONCATENATE("I"," told"," you ","10"," times",".","")' },
        
                    B1: { f: 'CONCATENATE("I"," told"," you ",SUM(1,2,3,4)," times",".","")' },
                    B2: { f: 'CONCATENATE("I"," told"," you ",SUM(1,2,3,SUM(4))," times",".","")' },
                    B3: { f: 'CONCATENATE("I"," told"," you ",SUM(1,2,3,SUM(4),0)," times",".","")' }
                }
            }
        };
        
        XLSX_CALC(workbook);
        
        let sheet = workbook.Sheets.Sample;
        
        assert.equal(sheet.A1.v, "I told you 10 times.");
        assert.equal(sheet.A2.v, "I told you 10 times.");
        
        assert.equal(sheet.B1.v, "I told you 10 times.");
        assert.equal(sheet.B2.v, "I told you 10 times.");
        assert.equal(sheet.B3.v, "I told you 10 times.");
    });
    
    describe('"ref is an error with new formula" error thrown when executing a formula containing a number division by a blank cell', () => {
        it('should not run that division', () => {
            let workbook = {
                Sheets: {
                    Sample: {
                        A1: {  },
                        A2: { f: 'IF(AND(ISNUMBER(A1),A1<>0),100/A1,"Number cannot be divided by 0")' },
                    }
                }
            };
            XLSX_CALC(workbook);
            let sheet = workbook.Sheets.Sample;
            assert.equal(sheet.A2.v, "Number cannot be divided by 0");
        });
        it('should not run that multiplication', () => {
            let workbook = {
                Sheets: {
                    Sample: {
                        A1: {  },
                        A2: { f: 'IF(AND(ISNUMBER(A1),A1<>0),100*A1,"Number cannot be 0")' },
                    }
                }
            };
            XLSX_CALC(workbook);
            let sheet = workbook.Sheets.Sample;
            assert.equal(sheet.A2.v, "Number cannot be 0");
        });
    });
    
});
