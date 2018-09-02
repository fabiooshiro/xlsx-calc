"use strict";

const int_2_col_str = require('./int_2_col_str.js');
const col_str_2_int = require('./col_str_2_int.js');
const exec_formula = require('./exec_formula.js');
const find_all_cells_with_formulas = require('./find_all_cells_with_formulas.js');
const Calculator = require('./Calculator.js');

var mymodule = function(workbook) {
    var formulas = find_all_cells_with_formulas(workbook, exec_formula);
    for (var i = formulas.length - 1; i >= 0; i--) {
        exec_formula(formulas[i]);
    }
};

mymodule.calculator = function calculator(workbook) {
    return new Calculator(workbook, exec_formula);
};

mymodule.set_fx = exec_formula.set_fx;
mymodule.exec_fx = exec_formula.exec_fx;
mymodule.col_str_2_int = col_str_2_int;
mymodule.int_2_col_str = int_2_col_str;
mymodule.import_functions = exec_formula.import_functions;
mymodule.import_raw_functions = exec_formula.import_raw_functions;
mymodule.xlsx_Fx = exec_formula.xlsx_Fx;



module.exports = mymodule;