"use strict";

const expression_builder = require('./expression_builder.js');

let xlsx_Fx = {};
let xlsx_raw_Fx = {};

import_functions(require('./formulas.js'));
import_raw_functions(require('./formulas-raw.js'));

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
    return expression_builder(formula, {xlsx_Fx: xlsx_Fx, xlsx_raw_Fx: xlsx_raw_Fx});
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

exec_formula.localizeFunctions = function(dic) {
    for (let newName in dic) {
        let oldName = dic[newName];
        if (xlsx_Fx[oldName]) {
            xlsx_Fx[newName] = xlsx_Fx[oldName];
        }
        if (xlsx_raw_Fx[oldName]) {
            xlsx_raw_Fx[newName] = xlsx_raw_Fx[oldName];
        }
    }
};

exec_formula.import_functions = import_functions;
exec_formula.import_raw_functions = import_raw_functions;
exec_formula.build_expression = build_expression;
exec_formula.xlsx_Fx = xlsx_Fx;
module.exports = exec_formula;
