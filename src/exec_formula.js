"use strict";

const Exp = require('./Exp.js');
const RawValue = require('./RawValue.js');
const UserFnExecutor = require('./UserFnExecutor.js');
const UserRawFnExecutor = require('./UserRawFnExecutor.js');

var xlsx_Fx = {};
var xlsx_raw_Fx = {};

import_functions(require('./formulas.js'));
import_raw_functions(require('./formulas-raw.js'));

const common_operations = {
    '*': 'multiply',
    '+': 'plus',
    '-': 'minus',
    '/': 'divide',
    '^': 'power',
    '&': 'concat',
    '<': 'lt',
    '>': 'gt',
    '=': 'eq'
};

function import_raw_functions(functions, opts) {
    for (var key in functions) {
        xlsx_raw_Fx[key] = functions[key];
    }
}

function import_functions(formulajs, opts) {
    opts = opts || {};
    var prefix = opts.prefix || '';
    for (var key in formulajs) {
        var obj = formulajs[key];
        if (typeof(obj) === 'function') {
            if (opts.override || !xlsx_Fx[prefix + key]) {
                xlsx_Fx[prefix + key] = obj;
            }
            // else {
            //     console.log(prefix + key, 'already exists.');
            //     console.log('  to override:');
            //     console.log('    XLSX_CALC.import_functions(yourlib, {override: true})');
            // }
        }
        else if (typeof(obj) === 'object') {
            import_functions(obj, my_assign(opts, { prefix: key + '.' }));
        }
    }
}

function my_assign(dest, source) {
    var obj = JSON.parse(JSON.stringify(dest));
    for (var k in source) {
        obj[k] = source[k];
    }
    return obj;
}

function build_expression(formula) {
    formula.status = 'working';
    var root_exp;
    var str_formula = formula.cell.f;
    if (str_formula[0] == '=') {
        str_formula = str_formula.substr(1);
    }
    var exp_obj = root_exp = new Exp(formula);
    var buffer = '',
        is_string = false,
        was_string = false;
    var fn_stack = [{
        exp: exp_obj
    }];
    for (var i = 0; i < str_formula.length; i++) {
        if (str_formula[i] == '"') {
            if (is_string) {
                exp_obj.push(new RawValue(buffer));
                is_string = false;
                was_string = true;
            }
            else {
                is_string = true;
            }
            buffer = '';
        }
        else if (is_string) {
            buffer += str_formula[i];
        }
        else if (str_formula[i] == '(') {
            var o, trim_buffer = buffer.trim(),
                special = xlsx_Fx[trim_buffer];
            var special_raw = xlsx_raw_Fx[trim_buffer];
            if (special_raw) {
                special = new UserRawFnExecutor(special_raw, formula);
            }
            else if (special) {
                special = new UserFnExecutor(special, formula);
            }
            else if (trim_buffer) {
                //Error: "Worksheet 1"!D145: Function INDEX not found
                throw new Error('"' + formula.sheet_name + '"!' + formula.name + ': Function ' + buffer + ' not found');
            }
            o = new Exp(formula);
            fn_stack.push({
                exp: o,
                special: special
            });
            exp_obj = o;
            buffer = '';
        }
        else if (common_operations[str_formula[i]]) {
            if (!was_string) {
                exp_obj.push(buffer);
            }
            was_string = false;
            exp_obj.push(str_formula[i]);
            buffer = '';
        }
        else if (str_formula[i] === ',' && fn_stack[fn_stack.length - 1].special) {
            was_string = false;
            fn_stack[fn_stack.length - 1].exp.push(buffer);
            fn_stack[fn_stack.length - 1].special.push(fn_stack[fn_stack.length - 1].exp);
            fn_stack[fn_stack.length - 1].exp = exp_obj = new Exp(formula);
            buffer = '';
        }
        else if (str_formula[i] == ')') {
            var v, stack = fn_stack.pop();
            exp_obj = stack.exp;
            exp_obj.push(buffer);
            v = exp_obj;
            buffer = '';
            exp_obj = fn_stack[fn_stack.length - 1].exp;
            if (stack.special) {
                stack.special.push(v);
                exp_obj.push(stack.special);
            }
            else {
                exp_obj.push(v);
            }
        }
        else {
            buffer += str_formula[i];
        }
    }
    root_exp.push(buffer);
    return root_exp;
}

function exec_formula(formula) {
    let root_exp = build_expression(formula);
    root_exp.update_cell_value();
}

exec_formula.set_fx = function set_fx(name, fn) {
    xlsx_Fx[name] = fn;
};

exec_formula.exec_fx = function exec_fx(name, args) {
    return xlsx_Fx[name].apply(this, args);
};

exec_formula.import_functions = import_functions;
exec_formula.import_raw_functions = import_raw_functions;
exec_formula.build_expression = build_expression;
exec_formula.xlsx_Fx = xlsx_Fx;
module.exports = exec_formula;
