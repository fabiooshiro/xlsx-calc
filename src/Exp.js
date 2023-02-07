"use strict";

const RawValue = require('./RawValue.js');
const Range = require('./Range.js');
const str_2_val = require('./str_2_val.js');
const MS_PER_DAY = 24 * 60 * 60 * 1000;
const col_str_2_int = require('./col_str_2_int.js');
const int_2_col_str = require('./int_2_col_str.js');
const dynamic_array_compatible = require('./dynamic_array_compatible.js');
const { getErrorValueByMessage } = require('./errors')
var exp_id = 0;

function isMatrix(obj) {
    return Array.isArray(obj) && (obj.length === 0 || Array.isArray(obj[0]));
}

module.exports = function Exp(formula) {
    var self = this;
    self.id = ++exp_id;
    self.args = [];
    self.name = 'Expression';
    self.update_cell_value = update_cell_value;
    self.formula = formula;
    
    function update_cell_value() {
        try {
            formula.cell.v = self.calc();
            formula.cell.t = getCellType(formula.cell.v);
            if (isMatrix(formula.cell.v)) {
                const array = formula.cell.v;
                formula.cell.v = undefined;
                let cellsName = formula.name;
                let colAndRow = cellsName.match(/([A-Z]+)([0-9]+)/);
                let colNumber = col_str_2_int(colAndRow[1]);
                let rowNumber = +colAndRow[2];
                for (let i = 0; i < array.length; i++) {
                    const newCellNumber = rowNumber + i;
                    for (let j = 0; j < array[i].length; j++) {
                        const newCellValue = array[i][j];
                        const destinationColumn = j + colNumber;
                        const destinationCellName = int_2_col_str(destinationColumn) + newCellNumber;
                        let cell = formula.sheet[destinationCellName];
                        if (!cell) {
                            cell = {};
                            formula.sheet[destinationCellName] = cell;
                        }
                        applyCellValue(cell, newCellValue);
                    }
                }
            }
        }
        catch (e) {
            if (!applyCellError(formula.cell, e)) {
                throw e;
            }
        }
        finally {
            formula.status = 'done';
        }
    }

    function applyCellError(cell, cellValueOrError) {
        const error = cellValueOrError || {};
        cell.t = 'e';
        const errorValue = getErrorValueByMessage(error.message);
        if (errorValue !== undefined) {
            cell.w = error.message;
            cell.v = errorValue;
            return true;
        } else {
            return false;
        }
    }

    function applyCellValue(cell, cellValueOrError) {
        if (cellValueOrError instanceof Error) {
            applyCellError(cell, cellValueOrError)
        } else {
            const newCellType = getCellType(cellValueOrError);
            cell.v = cellValueOrError;
            if (newCellType) cell.t = newCellType;
        }
    }

    function getCellType(cellValue) {
        if (typeof(cellValue) === 'string') {
            return 's';
        }
        else if (typeof(cellValue) === 'number') {
            return 'n';
        }
        else if (cellValue instanceof Error) {
            return 'e';
        }
    }

    function isEmpty(value) {
        return value === undefined || value === null || value === "";
    }
    
    function checkVariable(obj) {
        if (typeof obj.calc !== 'function') {
            throw new Error('Undefined ' + obj);
        }
    }
    
    function exec(op, args, fn) {
        for (var i = 0; i < args.length; i++) {
            if (args[i] === op) {
                try {
                    if (i===0 && op==='+') {
                        checkVariable(args[i + 1]);
                        let r = args[i + 1].calc();
                        args.splice(i, 2, new RawValue(r));
                    } else {
                        checkVariable(args[i - 1]);
                        checkVariable(args[i + 1]);

                        let a = args[i - 1].calc();
                        let b = args[i + 1].calc();
                        if (a instanceof Error) {
                            throw a;
                        }
                        if (b instanceof Error) {
                            throw b;
                        }

                        let r = fn(a, b);
                        args.splice(i - 1, 3, new RawValue(r));
                        i--;
                    }
                }
                catch (e) {
                    // console.log('[Exp.js] - ' + formula.name + ': evaluating ' + formula.cell.f + '\n' + e.message);
                    throw e;
                }
            }
        }
    }

    function toNegative(value) {
        return Array.isArray(value) ? value.map(toNegative) : -value;
    }

    function exec_minus(args) {
        for (var i = args.length; i--;) {
            if (args[i] === '-') {
                checkVariable(args[i + 1]);
                var b = args[i + 1].calc();
                if (i > 0 && typeof args[i - 1] !== 'string') {
                    args.splice(i, 1, '+');
                    if (b instanceof Date) {
                        b = Date.parse(b);
                        checkVariable(args[i - 1]);
                        var a = args[i - 1].calc();
                        if (a instanceof Date) {
                            a = Date.parse(a) / MS_PER_DAY;
                            b = b / MS_PER_DAY;
                            args.splice(i - 1, 1, new RawValue(a));
                        }
                    }
                    args.splice(i + 1, 1, new RawValue(toNegative(b)));
                }
                else {
                    if (typeof b === 'string') {
                        throw new Error('#VALUE!');
                    }
                    args.splice(i, 2, new RawValue(toNegative(b)));
                }
            }
        }
    }

    self.calc = function() {
        let args = self.args.concat();
        exec('^', args, dynamic_array_compatible(function(a, b) {
            return Math.pow(+a, +b);
        }));
        exec_minus(args);
        exec('/', args, dynamic_array_compatible(function(a, b) {
            if (b == 0) {
                throw Error('#DIV/0!');
            }
            return (+a) / (+b);
        }));
        exec('*', args, dynamic_array_compatible(function(a, b) {
            return (+a) * (+b);
        }));
        exec('+', args, dynamic_array_compatible(function(a, b) {
            if (a instanceof Date && typeof b === 'number') {
                b = b * MS_PER_DAY;
            }
            return (+a) + (+b);
        }));
        exec('&', args, dynamic_array_compatible(function(a, b) {
            var string = '';
            if (a !== null) {
                string += a;
            }
            if (b !== null) {
                string += b;
            }
            return string;
        }));
        exec('<', args, dynamic_array_compatible(function(a, b) {
            return a < b;
        }));
        exec('>', args, dynamic_array_compatible(function(a, b) {
            return a > b;
        }));
        exec('>=', args, dynamic_array_compatible(function(a, b) {
            return a >= b;
        }));
        exec('<=', args, dynamic_array_compatible(function(a, b) {
            return a <= b;
        }));
        exec('<>', args, dynamic_array_compatible(function(a, b) {
            if (a instanceof Date && b instanceof Date) {
                return a.getTime() !== b.getTime();
            }
            if (isEmpty(a) && isEmpty(b)) {
                return false;
            }
            return a !== b;
        }));
        exec('=', args, dynamic_array_compatible(function(a, b) {
            if (a instanceof Date && b instanceof Date) {
                return a.getTime() === b.getTime();
            }
            if (isEmpty(a) && isEmpty(b)) {
                return true;
            }
            if ((a == null && b === 0) || (a === 0 && b == null)) {
                return true;
            }
            if (typeof a === 'string' && typeof b === 'string' && a.toLowerCase() === b.toLowerCase()) {
                return true;
            }
            return a === b;
        }));
        if (args.length == 1) {
            if (typeof(args[0].calc) !== 'function') {
                return args[0];
            }
            else {
                return args[0].calc();
            }
        }
    };

    var last_arg;
    self.push = function(buffer) {
        if (buffer) {
            var v = str_2_val(buffer, formula);
            if (((v === '=') && (last_arg == '>' || last_arg == '<')) || (last_arg == '<' && v === '>')) {
                self.args[self.args.length - 1] += v;
            }
            else {
                self.args.push(v);
            }
            last_arg = v;
            //console.log(self.id, '-->', v);
        }
    };
};