"use strict";

const int_2_col_str = require('./int_2_col_str.js');
const col_str_2_int = require('./col_str_2_int.js');
const exec_formula = require('./exec_formula.js');
const find_all_cells_with_formulas = require('./find_all_cells_with_formulas.js');
const Calculator = require('./Calculator.js');


function exec_next(formulas, i, resolve, reject) {
    if (i === formulas.length) {
        resolve();
        return;
    }
    //console.log('executing', i+1, 'of', formulas.length,'...');
    exec_formula(formulas[i]).then(x => {
        exec_next(formulas, i + 1, resolve, reject);
    }).catch(err => {
        //console.error(err);
        reject(err);
    });
}

var XLSX_CALC = function(workbook) {
    return new Promise((resolve, reject) => {
        var formulas = find_all_cells_with_formulas(workbook, exec_formula);
        exec_next(formulas, 0, resolve, reject);
        //for (var i = formulas.length - 1; i >= 0; i--) {
        //    exec_formula(formulas[i]);
        //}
        //resolve();
    });
};

XLSX_CALC.calculator = function calculator(workbook) {
    return new Calculator(workbook, exec_formula);
};

XLSX_CALC.set_fx = exec_formula.set_fx;
XLSX_CALC.exec_fx = exec_formula.exec_fx;
XLSX_CALC.col_str_2_int = col_str_2_int;
XLSX_CALC.int_2_col_str = int_2_col_str;
XLSX_CALC.import_functions = exec_formula.import_functions;
XLSX_CALC.import_raw_functions = exec_formula.import_raw_functions;



module.exports = XLSX_CALC;