"use strict";

const int_2_col_str = require('./int_2_col_str.js');
const col_str_2_int = require('./col_str_2_int.js');
const exec_formula = require('./exec_formula.js');
const find_all_cells_with_formulas = require('./find_all_cells_with_formulas.js');
const Calculator = require('./Calculator.js');

function XLSX_CALC(workbook, options) {
    let opts = options || {}
    var formulas = find_all_cells_with_formulas(workbook, exec_formula);
    for (var i = formulas.length - 1; i >= 0; i--) {
        try {
            exec_formula(formulas[i]);
        } catch (e) {
            if (opts.throwErrors === false) {
                if (opts.logErrors !== false) {
                    console.error(e)
                }
            } else {
                throw e
            }
        }
    }
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
XLSX_CALC.xlsx_Fx = exec_formula.xlsx_Fx;
XLSX_CALC.localizeFunctions = exec_formula.localizeFunctions;

XLSX_CALC.XLSX_CALC = XLSX_CALC

module.exports = XLSX_CALC;
