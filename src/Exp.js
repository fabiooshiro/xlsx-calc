"use strict";

const RawValue = require('./RawValue.js');
const RefValue = require('./RefValue.js');
const Range = require('./Range.js');
const str_2_val = require('./str_2_val.js');

var exp_id = 0;

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
            if (typeof(formula.cell.v) === 'string') {
                formula.cell.t = 's';
            }
            else if (typeof(formula.cell.v) === 'number') {
                formula.cell.t = 'n';
            }
        }
        catch (e) {
            var errorValues = {
                '#NULL!': 0x00,
                '#DIV/0!': 0x07,
                '#VALUE!': 0x0F,
                '#REF!': 0x17,
                '#NAME?': 0x1D,
                '#NUM!': 0x24,
                '#N/A': 0x2A,
                '#GETTING_DATA': 0x2B
            };
            if (errorValues[e.message] !== undefined) {
                formula.cell.t = 'e';
                formula.cell.w = e.message;
                formula.cell.v = errorValues[e.message];
            }
            else {
                throw e;
            }
        }
        finally {
            formula.status = 'done';
        }
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
                        var r = args[i + 1].calc();
                        args.splice(i, 2, new RawValue(r));
                    } else {
                        checkVariable(args[i - 1]);
                        checkVariable(args[i + 1]);
                        var r = fn(args[i - 1].calc(), args[i + 1].calc());
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

    function exec_minus(args) {
        for (var i = args.length; i--;) {
            if (args[i] === '-') {
                checkVariable(args[i + 1]);
                var r = -args[i + 1].calc();
                if (typeof args[i - 1] !== 'string' && i > 0) {
                    args.splice(i, 1, '+');
                    args.splice(i + 1, 1, new RawValue(r));
                }
                else {
                    args.splice(i, 2, new RawValue(r));
                }
            }
        }
    }

    self.calc = function() {
        let args = self.args.concat();
        exec_minus(args);
        exec('^', args, function(a, b) {
            return Math.pow(+a, +b);
        });
        exec('/', args, function(a, b) {
            if (b == 0) {
                throw Error('#DIV/0!');
            }
            return (+a) / (+b);
        });
        exec('*', args, function(a, b) {
            return (+a) * (+b);
        });
        exec('+', args, function(a, b) {
            return (+a) + (+b);
        });
        exec('&', args, function(a, b) {
            return '' + a + b;
        });
        exec('<', args, function(a, b) {
            return a < b;
        });
        exec('>', args, function(a, b) {
            return a > b;
        });
        exec('>=', args, function(a, b) {
            return a >= b;
        });
        exec('<=', args, function(a, b) {
            return a <= b;
        });
        exec('<>', args, function(a, b) {
            return a != b;
        });
        exec('=', args, function(a, b) {
            return a == b;
        });
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