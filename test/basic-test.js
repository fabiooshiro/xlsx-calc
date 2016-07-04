var XLSX_CALC = require("../");
var assert = require('assert');

describe('XLSX_CALC', function() {
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
    });
    describe('power', function() {
        it('should calc 2^10', function() {
            workbook.Sheets.Sheet1.A1.f = '2^10';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 1024);
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
            workbook.Sheets.Sheet1.A1.f = 'SUM(C:C)';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, 7);
        });
    });
    describe('boolean', function() {
        it('evaluates 1<2 as true', function() {
            workbook.Sheets.Sheet1.A1.f = '1<2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, true);
        });
        it('evaluates 1>2 as false', function() {
            workbook.Sheets.Sheet1.A1.f = '1>2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, false);
        });
        it('evaluates 2=2 as true', function() {
            workbook.Sheets.Sheet1.A1.f = '2=2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, true);
        });
        it('evaluates 2>=2 as true', function() {
            workbook.Sheets.Sheet1.A1.f = '2>=2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, true);
        });
        it('evaluates 1>=2 as true', function() {
            workbook.Sheets.Sheet1.A1.f = '1>=2';
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A1.v, false);
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
    });
    it('calcs ref with space', function() {
        workbook.Sheets.Sheet1.A1.f = 'A2 ';
        workbook.Sheets.Sheet1.A2.v = 1979;
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets.Sheet1.A1.v, 1979);
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
            /Function XPTO not found/
        );
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
    describe('bug fix', function() {
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
    
    // describe('HELLO', function() {
    //     it('says: Hello, World!', function() {
    //         workbook.Sheets.Sheet1.A1.f = 'HELLO("World")';
    //         XLSX_CALC(workbook);
    //         assert.equal(workbook.Sheets.Sheet1.A1.v, "Hello, World!");
    //     });
    // });
});