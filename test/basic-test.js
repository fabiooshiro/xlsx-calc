"use strict";

const XLSX_CALC = require("../");
const assert = require('assert');

describe('XLSX_CALC', function() {
    let workbook;
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
    describe('plus', function(done) {
        it('should calc A2+C5', function() {
            workbook.Sheets.Sheet1.A2.v = 7;
            workbook.Sheets.Sheet1.C5.v = 3;
            workbook.Sheets.Sheet1.A1.f = 'A2+C5';
            XLSX_CALC(workbook).then(function() {
                assert.equal(workbook.Sheets.Sheet1.A1.v, 10);
                done();
            }).catch(done);
        });
        it('should calc 1+2', function(done) {
            workbook.Sheets.Sheet1.A1.f = '1+2';
            XLSX_CALC(workbook).then(function() {
                assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
                done();
            }).catch(done);
        });
        it('should calc 1+2+3', function() {
            workbook.Sheets.Sheet1.A1.f = '1+2+3';
            XLSX_CALC(workbook).then(function() {
                assert.equal(workbook.Sheets.Sheet1.A1.v, 6);
                done();
            }).catch(done);
        });
    });
    describe('minus', function() {
        it('should update the property A1.v with result of formula A2-C4', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'A2-C4';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
                done();
            }).catch(done);
        });
        it('should calc A2-4', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'A2-4';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
                done();
            }).catch(done);
        });
        it('should calc 2-3', function(done) {
            workbook.Sheets.Sheet1.A1.f = '2-3';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, -1);
                done();
            }).catch(done);
        });
        it('should calc 2-3-4', function(done) {
            workbook.Sheets.Sheet1.A1.f = '2-3-4';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, -5);
                done();
            }).catch(done);
        });
        it('should calc -2-3-4', function(done) {
            workbook.Sheets.Sheet1.A1.f = '-2-3-4';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, -9);
                done();
            }).catch(done);
        });
    });
    describe('multiply', function() {
        it('should calc A2*C5', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'A2*C5';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 21);
                done();
            }).catch(done);
        });
        it('should calc A2*4', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'A2*4';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 28);
                done();
            }).catch(done);
        });
        it('should calc 4*A2', function(done) {
            workbook.Sheets.Sheet1.A1.f = '4*A2';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 28);
                done();
            }).catch(done);
        });
        it('should calc 2*3', function(done) {
            workbook.Sheets.Sheet1.A1.f = '2*3';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 6);
                done();
            }).catch(done);
        });
    });
    describe('divide', function() {
        it('should calc A2/C4', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'A2/C4';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 3.5);
                done();
            }).catch(done);
        });
        it('should calc A2/14', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'A2/14';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 0.5);
                done();
            }).catch(done);
        });
        it('should calc 7/2/2', function(done) {
            workbook.Sheets.Sheet1.A1.f = '7/2/2';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 1.75);
                done();
            }).catch(done);
        });
        it('should divide', function(done) {
            workbook.Sheets.Sheet1.A1.f = '=20/10';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 2);
                done();
            }).catch(done);
        });
    });
    describe('power', function() {
        it('should calc 2^10', function(done) {
            workbook.Sheets.Sheet1.A1.f = '2^10';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 1024);
                done();
            }).catch(done);
        });
    });
    describe('SQRT', function() {
        it('should calc SQRT(25)', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'SQRT(25)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
                done();
            }).catch(done);
        });
    });
    describe('ABS', function() {
        it('should calc ABS(-3.5)', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'ABS(-3.5)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 3.5);
                done();
            }).catch(done);
        });
    });
    describe('FLOOR', function() {
        it('should calc FLOOR(12.5)', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'FLOOR(12.5)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 12);
                done();
            }).catch(done);
        });
    });
    describe('.set_fx', function() {
        it('sets new function', function(done) {
            XLSX_CALC.set_fx('ADD_1', function(arg) {
                return arg + 1;
            });
            workbook.Sheets.Sheet1.A1.f = 'ADD_1(123)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 124);
                done();
            }).catch(done);
        });
        it('sets promise function', function(done) {
            XLSX_CALC.set_fx('ADD_1', function(arg) {
                return new Promise((resolve, reject) => {
                    setTimeout(() => {
                        resolve(arg + 1);
                    }, 50);
                });
            });
            workbook.Sheets.Sheet1.A1.f = 'ADD_1(123)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 124);
                done();
            }).catch(done);
        });
    });
    describe('expression', function() {
        it('should calc 8/2+1', function(done) {
            workbook.Sheets.Sheet1.A1.f = '8/2+1';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
                done();
            }).catch(done);
        });
        it('should calc 1+8/2', function(done) {
            workbook.Sheets.Sheet1.A1.f = '1+8/2';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
                done();
            }).catch(done);
        });
        it('should calc 2*3+1', function(done) {
            workbook.Sheets.Sheet1.A1.f = '2*3+1';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 7);
                done();
            }).catch(done);
        });
        it('should calc 2*3-1', function(done) {
            workbook.Sheets.Sheet1.A1.f = '2*3-1';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 5);
                done();
            }).catch(done);
        });
        it('should calc 2*(3-1)', function(done) {
            workbook.Sheets.Sheet1.A1.f = '2*(3-1)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 4);
                done();
            }).catch(done);
        });
        it('should calc (3-1)*5', function(done) {
            workbook.Sheets.Sheet1.A1.f = '(3-1)*5';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 10);
                done();
            }).catch(done);
        });
        it('should calc (3-1)*(4+1)', function(done) {
            workbook.Sheets.Sheet1.A1.f = '(3-1)*(4+1)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 10);
                done();
            }).catch(done);
        });
        it('should calc -1*2', function(done) {
            workbook.Sheets.Sheet1.A1.f = '-1*2';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, -2);
                done();
            }).catch(done);
        });
        it('should calc 1*-2', function(done) {
            workbook.Sheets.Sheet1.A1.f = '1*-2';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, -2);
                done();
            }).catch(done);
        });
        it('should calc (3*10)-(2-1)', function(done) {
            workbook.Sheets.Sheet1.A1.f = '(3*10)-(2-1)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 29);
                done();
            }).catch(done);
        });
        it('should calc (3*10)-(2-(3*5))', function(done) {
            workbook.Sheets.Sheet1.A1.f = '(3*10)-(2-(3*5))';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 43);
                done();
            }).catch(done);
        });
        it('should calc (3*10)-(-2-(3*5))', function(done) {
            workbook.Sheets.Sheet1.A1.f = '(3*10)-(-2-(3*5))';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 47);
                done();
            }).catch(done);
        });
    });
    describe('SUM', function() {
        it('makes the sum', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'SUM(C3:C4)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
                done();
            }).catch(done);
        });
        it('makes the sum of a bigger range', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'SUM(C3:C5)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 6);
                done();
            }).catch(done);
        });
        it('sums numbers', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'SUM(1,2,3)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 6);
                done();
            }).catch(done);
        });
    });
    describe('MAX formula', function() {
        it('finds the max in range', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'MAX( C3:C5 )';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
                done();
            }).catch(done);
        });
        it('finds the max in range including some cell', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'MAX(C3:C5 ,A2)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 7);
                done();
            }).catch(done);
        });
        it('finds the max in range including some cell', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'MAX(A2,C3:C5)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 7);
                done();
            }).catch(done);
        });
        it('finds the max in args', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'MAX(1,2,10,3,4)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 10);
                done();
            }).catch(done);
        });
        it('finds the max in negative args', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'MAX(-1,-2,-10,-3,-4)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, -1);
                done();
            }).catch(done);
        });
        it('finds the max in range including some negative cell', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'MAX(C3:C5,-A2)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
                done();
            }).catch(done);
        });
    });
    describe('MIN formula', function() {
        it('finds the min in range', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'MIN(C3:C5)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 1);
                done();
            }).catch(done);
        });
        it('finds the min in range including some negative cell', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'MIN(C3:C5,-A2)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, -7);
                done();
            }).catch(done);
        });
    });
    describe('MAX and SUM', function() {
        it('evaluates MAX(1,2,SUM(10,5),7,3,4)', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'MAX(1,2,SUM(10,5),7,3,4)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 15);
                done();
            }).catch(done);
        });
        it('evaluates MAX(1,2, SUM(10,5),7,3,4)', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'MAX(1,2, SUM(10,5),7,3,4)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 15);
                done();
            }).catch(done);
        });
    });
    describe('&', function() {
        it('evaluates "concat "&A2', function(done) {
            workbook.Sheets.Sheet1.A1.f = '"concat " & A2';
            workbook.Sheets.Sheet1.A2.v = 7;
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 'concat 7');
                done();
            }).catch(done);
        });
        it('evaluates "concat +1" & A2', function(done) {
            workbook.Sheets.Sheet1.A1.f = '"concat +1" & A2';
            workbook.Sheets.Sheet1.A2.v = 7;
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 'concat +17');
                done();
            }).catch(done);
        });
        it('evaluates A2 & "concat"', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'A2 & "concat +1"';
            workbook.Sheets.Sheet1.A2.v = 7;
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, '7concat +1');
                done();
            }).catch(done);
        });
    });
    describe('CONCATENATE', function() {
        it('concatenates 1,2,3', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'CONCATENATE(1,2,3)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, '123');
                done();
            }).catch(done);
        });
        it('concatenates A2,"xxx"', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'CONCATENATE(A2 , "xxx")';
            workbook.Sheets.Sheet1.A2.v = 79;
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, '79xxx');
                done();
            }).catch(done);
        });
    });
    describe('range', function() {
        function cell(name, obj) {
            workbook.Sheets.Sheet1[name] = obj;
        }
        it('should eval the expression in range of sum', function(done) {
            cell('A1', {f: 'SUM(C3:C5)'});
            cell('A2', {v: 7});
            cell('C3', {v: 1});
            cell('C4', {v: 2, f: 'A2'});
            cell('C5', {v: 1, f: 'A2'});
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 15);
                assert.equal(workbook.Sheets.Sheet1.C4.v, 7);
                assert.equal(workbook.Sheets.Sheet1.C5.v, 7);
                done();
            }).catch(done);
        });
        it('should calc range with $', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'SUM($C$3:$C$4)';
            workbook.Sheets.Sheet1.C4.f = 'A2';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 8);
                assert.equal(workbook.Sheets.Sheet1.C4.v, 7);
                done();
            }).catch(done);
        });
        it('should calc range like C:C using !ref', function(done) {
            workbook.Sheets.Sheet1['!ref'] = 'A1:C4';
            workbook.Sheets.Sheet1.A1.f = 'SUM(C:C)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 4);
                done();
            }).catch(done);
        });
        it('should calc range like C:C without !ref', function(done) {
            this.timeout(5000);
            workbook.Sheets.Sheet1.A1.f = 'SUM(C:C)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 7);
                done();
            }).catch(done);
        });
    });
    describe('boolean', function() {
        it('evaluates 1<2 as true', function(done) {
            workbook.Sheets.Sheet1.A1.f = '1<2';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, true);
                done();
            }).catch(done);
        });
        it('evaluates 1>2 as false', function(done) {
            workbook.Sheets.Sheet1.A1.f = '1>2';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, false);
                done();
            }).catch(done);
        });
        it('evaluates 2=2 as true', function(done) {
            workbook.Sheets.Sheet1.A1.f = '2=2';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, true);
                done();
            }).catch(done);
        });
        it('evaluates 2>=2 as true', function(done) {
            workbook.Sheets.Sheet1.A1.f = '2>=2';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, true);
                done();
            }).catch(done);
        });
        it('evaluates 1>=2 as true', function(done) {
            workbook.Sheets.Sheet1.A1.f = '1>=2';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, false);
                done();
            }).catch(done);
        });
        it('evaluates 2<=2 as true', function(done) {
            workbook.Sheets.Sheet1.A1.f = '2<=2';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, true);
                done();
            }).catch(done);
        });
        it('evaluates 3<=2 as false', function(done) {
            workbook.Sheets.Sheet1.A1.f = '3<=2';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, false);
                done();
            }).catch(done);
        });
        it('evaluates C3<=C5 as false', function(done) {
            workbook.Sheets.Sheet1.C3.v = 3;
            workbook.Sheets.Sheet1.C5.v = 2;
            workbook.Sheets.Sheet1.A1.f = 'C3<=C5';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, false);
                done();
            }).catch(done);
        });
        it('evaluates 1<>1 as false', function(done) {
            workbook.Sheets.Sheet1.A1.f = '1<>1';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, false);
                done();
            }).catch(done);
        });
        it('evaluates 2<>1 as true', function(done) {
            workbook.Sheets.Sheet1.A1.f = '2<>1';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, true);
                done();
            }).catch(done);
        });
    });
    
    describe('IF function', function() {
        it('should exec true', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'IF(1<2,123,0)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 123);
                done();
            }).catch(done);
        });
        it('should exec false', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'IF(1>2,0,123)';
            XLSX_CALC(workbook).then(r=>{
                assert.equal(workbook.Sheets.Sheet1.A1.v, 123);
                done();
            }).catch(done);
        });
    });
    
    describe('references', function() {
        it('calcs ref with space', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'A2 ';
            workbook.Sheets.Sheet1.A2.v = 1979;
            XLSX_CALC(workbook).then(res => {
                assert.equal(workbook.Sheets.Sheet1.A1.v, 1979);
                done();
            }).catch(done);
        });
        it('calcs ref with $', function(done) {
            workbook.Sheets.Sheet1.A1.f = '$A$2 ';
            workbook.Sheets.Sheet1.A2.v = 1979;
            XLSX_CALC(workbook).then(res => {
                assert.equal(workbook.Sheets.Sheet1.A1.v, 1979);
                done();
            }).catch(done);
        });
        it('calcs simple chain', function(done) {
            workbook.Sheets.Sheet1.C4.f = 'A1';
            workbook.Sheets.Sheet1.A1.f = 'A2';
            workbook.Sheets.Sheet1.A2.v = 1979;
            XLSX_CALC(workbook).then(res => {
                assert.equal(workbook.Sheets.Sheet1.C4.v, 1979);
                done();
            }).catch(done);
        });
        it('calcs long chain', function(done) {
            workbook.Sheets.Sheet1.C4.f = 'C3';
            workbook.Sheets.Sheet1.C3.f = 'C2';
            workbook.Sheets.Sheet1.C2.f = 'A2';
            workbook.Sheets.Sheet1.A2.f = 'A1';
            workbook.Sheets.Sheet1.A1.v = 1979;
            workbook.Sheets.Sheet1.C5.f = 'C3';
            XLSX_CALC(workbook).then(res => {
                assert.equal(workbook.Sheets.Sheet1.C4.v, 1979);
                done();
            }).catch(done);
        });
        it('throws a circular exception', function(done) {
            workbook.Sheets.Sheet1.C4.f = 'A1';
            workbook.Sheets.Sheet1.A1.f = 'C4';
            XLSX_CALC(workbook).then(x=> {
                done(new Error('Where is the error?'));
            }).catch(err => {
                assert.throws(
                    function() {
                        throw err;
                    },
                    /Circular ref/
                );
                done();
            });
        });
    });
    
    it('throws a function XPTO not found', function(done) {
        workbook.Sheets.Sheet1.A1.f = 'XPTO()';
        XLSX_CALC(workbook).then(x => {
            done('Missing expected exception.');
        }).catch(err => {
            assert.throws(
                function() {
                    throw err;
                },
                /"Sheet1"!A1.*Function XPTO not found/
            );
            done();
        });
    });
    
    describe('Common excel formulas', function() {
        describe('PTM', function() {
            it('calcs PMT(0.07/12, 24, 1000)', function(done) {
                workbook.Sheets.Sheet1.A1.f = 'PMT(0.07/12, 24, 1000)';
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A1.v, -44.77257910314528);
                    done();
                }).catch(done);
            });
            it('calcs PMT(0.07/12, 24, 1000,2000,0)', function(done) {
                workbook.Sheets.Sheet1.A1.f = 'PMT(0.07/12, 24, 1000,2000,0)';
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A1.v, -122.6510706427692);
                    done();
                }).catch(done);
            });
        });
        
        describe('COUNTA', function() {
            it('counts non empty cells', function(done) {
                workbook.Sheets.Sheet1.A1.f = 'COUNTA(B1:B3)';
                workbook.Sheets.Sheet1.B1 = {v:1};
                workbook.Sheets.Sheet1.B2 = {};
                workbook.Sheets.Sheet1.B3 = {v:1};
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A1.v, 2);
                    done();
                }).catch(done);
            });
        });
        describe('NORM.INV', function() {
            it('should call normsInv', function(done) {
                workbook.Sheets.Sheet1.A1.f = 'NORM.INV(0.05, -0.0015, 0.0175)';
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A1.v, -0.030284938471650775);
                    done();
                }).catch(done);
            });
        });
        describe('STDEV', function() {
            it('should calc STDEV', function(done) {
                workbook.Sheets.Sheet1.A1.f = 'STDEV(6.2,5,4.5,6,6,6.9,6.4,7.5)';
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A1.v, 0.96204766736670300000);
                    done();
                }).catch(done);
            });
        });
        describe('AVERAGE', function() {
            it('should calc AVERAGE', function(done) {
                workbook.Sheets.Sheet1.A1.f = 'AVERAGE(1,2,3,4,5)';
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
                    done();
                }).catch(done);
            });
            it('should calc AVERAGE of range', function(done) {
                workbook.Sheets.Sheet1.A1 = {v: 0.1};
                workbook.Sheets.Sheet1.A2 = {v: 0.5};
                workbook.Sheets.Sheet1.A3 = {v: 0.2};
                workbook.Sheets.Sheet1.A4 = {v: 0.3};
                workbook.Sheets.Sheet1.A5 = {v: 0.2};
                workbook.Sheets.Sheet1.A6 = {v: 0.2};
                workbook.Sheets.Sheet1.A7 = {f: 'AVERAGE(A1:A6)'};
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A7.v, 0.25);
                    done();
                }).catch(done);
            });
        });
        describe('IRR', function() {
            it('calcs IRR', function(done) {
                workbook.Sheets.Sheet1.A1.f = 'IRR(B1:B3)';
                workbook.Sheets.Sheet1.B1 = {v: -10.0};
                workbook.Sheets.Sheet1.B2 = {v:  -1.0};
                workbook.Sheets.Sheet1.B3 = {v:   2.9};
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A1.v, -0.5091672986745834);
                    done();
                }).catch(done);
            });
            it('calcs IRR 2', function(done) {
                workbook.Sheets.Sheet1.A1.f = 'IRR(B1:B3)';
                workbook.Sheets.Sheet1.B1 = {v: -100.0};
                workbook.Sheets.Sheet1.B2 = {v:   10.0};
                workbook.Sheets.Sheet1.B3 = {v:  100000.0};
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A1.v, 30.672816276550293);
                    done();
                }).catch(done);
            });
        });
        describe('VAR.P', function() {
            it('calcs VAR.P', function(done) {
                workbook.Sheets.Sheet1.A1 = {v: 0.1};
                workbook.Sheets.Sheet1.A2 = {v: 0.5};
                workbook.Sheets.Sheet1.A3 = {v: 0.2};
                workbook.Sheets.Sheet1.A4 = {v: 0.3};
                workbook.Sheets.Sheet1.A5 = {v: 0.2};
                workbook.Sheets.Sheet1.A6 = {v: 0.2};
                workbook.Sheets.Sheet1.A7 = {f: 'VAR.P(A1:A6)'};
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A7.v.toFixed(8), 0.01583333);
                    done();
                }).catch(done);
            });
            it('calls the VAR.P', function() {
                var x = XLSX_CALC.exec_fx('VAR.P', [0.1, 0.5, 0.2, 0.3, 0.2, 0.2]);
                assert.equal(x.toFixed(8), 0.01583333);
            });
        });
        describe('EXP', function() {
            it('calculates EXP', function(done) {
                workbook.Sheets.Sheet1.A1.f = 'EXP(2)';
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A1.v, 7.3890560989306495);
                    done();
                }).catch(done);
            });
        });
        describe('LN', function() {
            it('calculates LN of a number', function(done) {
                workbook.Sheets.Sheet1.A1.f = 'LN(EXP(2))';
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A1.v, 2);
                    done();
                }).catch(done);
            });
        });
        describe('ISBLANK', function() {
            it('calculates ISBLANK as false', function(done) {
                workbook.Sheets.Sheet1.A1.f = 'ISBLANK(B1)';
                workbook.Sheets.Sheet1.B1 = {v: ' '};
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A1.v, false);
                    done();
                }).catch(done);
            });
            it('calculates ISBLANK as true', function(done) {
                workbook.Sheets.Sheet1.A1.f = 'ISBLANK(B1)';
                workbook.Sheets.Sheet1.B1 = {v: ''};
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A1.v, true);
                    done();
                }).catch(done);
            });
        });
        describe('COVARIANCE.P', function() {
            it('computes COVARIANCE.P', function(done) {
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
                XLSX_CALC(workbook).then(x => {
                    assert.equal(workbook.Sheets.Sheet1.A7.v.toFixed(8), 0.00158333);
                    done();
                }).catch(done);
            });
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
    
    describe('Sheet ref references', function(done) {
        it('calculates the sum of Sheet2!A1+Sheet2!A2', function() {
            workbook.Sheets.Sheet1.A1.f = 'Sheet2!A1+Sheet2!A2';
            workbook.Sheets.Sheet2 = { A1: {v:1}, A2: {v:2}};
            XLSX_CALC(workbook).then(x => {
                assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
                done();
            }).catch(done);
        });
        it('calculates the sum of Sheet2!A1:A2', function(done) {
            workbook.Sheets.Sheet1.A1.f = 'SUM(Sheet2!A1:A2)';
            workbook.Sheets.Sheet2 = { A1: {v:1}, A2: {v:2}};
            XLSX_CALC(workbook).then(x => {
                assert.equal(workbook.Sheets.Sheet1.A1.v, 3);
                done();
            }).catch(done);
        });
        it('calculates the sum of Sheet2!A:B', function(done) {
            this.timeout(5000);
            workbook.Sheets.Sheet1.A1.f = 'SUM(Sheet2!A:B)';
            workbook.Sheets.Sheet2 = { A1: {v:1}, B1: {v:2}, A2: {v: 3}};
            XLSX_CALC(workbook).then(x => {
                assert.equal(workbook.Sheets.Sheet1.A1.v, 6);
                done();
            }).catch(done);
        });
    });
    describe('Cell type: A2.t = "s" or A2.t = "n"', function() {
        it('should set t = "s" for string values', function(done) {
            workbook.Sheets.Sheet1.A1 = { v: " some string " };
            workbook.Sheets.Sheet1.A2 = { f: "TRIM(A1)" };
            
            /* calculate */
            XLSX_CALC(workbook).then(x => {
                assert.equal(workbook.Sheets.Sheet1.A2.t, 's');
                assert.equal(workbook.Sheets.Sheet1.A2.v, 'some string');
                done();
            }).catch(done);
        });
        it('should set t = "n" for numeric values', function(done) {
            workbook.Sheets.Sheet1.A1 = { v: " some string " };
            workbook.Sheets.Sheet1.A2 = { f: "LEN(TRIM(A1))" };
            
            /* calculate */
            XLSX_CALC(workbook).then(x => {
                assert.equal(workbook.Sheets.Sheet1.A2.t, 'n');
                assert.equal(workbook.Sheets.Sheet1.A2.v, 11);
                done();
            }).catch(done);
        });
    });
    describe('raw function importer', function() {
        it('should sends the raw argument', function(done) {
            workbook.Sheets.Sheet1.A1 = { f: "MYRAWFN(A2,3-2,0)"};
            workbook.Sheets.Sheet1.A2 = { v: "VaLuE"};
            workbook.Sheets.Sheet1.B1 = { v: 1};
            XLSX_CALC.import_raw_functions({
                MYRAWFN: function(expr1, expr2, expr3) {
                    console.log(expr1.name); // Expression
                    console.log(expr1.args[0].name); // RefValue
                    console.log(expr1.args[0].str_expression); // A2
                    console.log(expr1.args[0].calc()); // Promise { 'VaLuE' }
                    return new Promise((resolve, reject) => {
                        Promise.all([expr2.calc(),expr3.calc()]).then(rs => {
                            resolve([expr1.args[0].str_expression, rs[0], rs[1]]);
                        });
                    });
                },
            });
            XLSX_CALC(workbook).then(x => {
                assert.deepEqual(workbook.Sheets.Sheet1.A1.v, ['A2',1,0]);
                done();
            }).catch(done);
        });
    });
    
    describe('IFERROR', function() {
        it('returns the string Error', function(done) {
            workbook.Sheets.Sheet1.A1 = { f: "IFERROR(A2,\"Error\")"};
            workbook.Sheets.Sheet1.A2 = { f: "0/0"};
            XLSX_CALC(workbook).then(x => {
                assert.equal(workbook.Sheets.Sheet1.A1.v, 'Error');
                done();
            }).catch(done);
        });
        it('returns the string Error when res is Infinity', function(done) {
            workbook.Sheets.Sheet1.A1 = { f: "IFERROR(1/0,\"Error\")"};
            XLSX_CALC(workbook).then(x => {
                assert.equal(workbook.Sheets.Sheet1.A1.v, 'Error');
                done();
            }).catch(done);
        });
        it('returns the string Error when res is -Infinity', function(done) {
            workbook.Sheets.Sheet1.A1 = { f: "IFERROR(-1/0,\"Error\")"};
            XLSX_CALC(workbook).then(x => {
                assert.equal(workbook.Sheets.Sheet1.A1.v, 'Error');
                done();
            }).catch(done);
        });
        it('returns the string boston', function(done) {
            workbook.Sheets.Sheet1.A1 = { f: "IFERROR(A2,\"Error\")"};
            workbook.Sheets.Sheet1.A2 = { v: "boston"};
            XLSX_CALC(workbook).then(x => {
                assert.equal(workbook.Sheets.Sheet1.A1.v, 'boston');
                done();
            }).catch(done);
        });
        it('returns the string Error when VLOOKUP fail', function(done) {
            workbook.Sheets.Sheet1.A1 = { f: "IFERROR(A2,\"Error\")"};
            workbook.Sheets.Sheet1.A2 = { f: "VLOOKUP(\"void\",\"A3:B7\",2)"};
            XLSX_CALC(workbook).then(x => {
                //console.log('calculou');
                assert.equal(workbook.Sheets.Sheet1.A1.v, 'Error');
                done();
            }).catch(e => {
                //console.log('Test error?', e);
                done(e);
            });
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