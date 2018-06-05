"use strict";

const col_str_2_int = require('./col_str_2_int.js');
const int_2_col_str = require('./int_2_col_str.js');
const getSanitizedSheetName = require('./getSanitizedSheetName.js');

module.exports = function Range(str_expression, formula) {
    this.calc = function() {
        var range_expression, sheet_name, sheet;
        if (str_expression.indexOf('!') != -1) {
            var aux = str_expression.split('!');
            sheet_name = getSanitizedSheetName(aux[0]);
            range_expression = aux[1];
        }
        else {
            sheet_name = formula.sheet_name;
            range_expression = str_expression;
        }
        sheet = formula.wb.Sheets[sheet_name];
        var arr = range_expression.split(':');
        var min_row = parseInt(arr[0].replace(/^[A-Z]+/, ''), 10) || 0;
        var str_max_row = arr[1].replace(/^[A-Z]+/, '');
        var max_row;
        if (str_max_row === '' && sheet['!ref']) {
            str_max_row = sheet['!ref'].split(':')[1].replace(/^[A-Z]+/, '');
        }
        // the max is 1048576, but TLE
        max_row = parseInt(str_max_row == '' ? '500000' : str_max_row, 10);
        var min_col = col_str_2_int(arr[0]);
        var max_col = col_str_2_int(arr[1]);
        var matrix = [];
        for (var i = min_row; i <= max_row; i++) {
            var row = [];
            matrix.push(row);
            for (var j = min_col; j <= max_col; j++) {
                var cell_name = int_2_col_str(j) + i;
                var cell_full_name = sheet_name + '!' + cell_name;
                if (formula.formula_ref[cell_full_name]) {
                    if (formula.formula_ref[cell_full_name].status === 'new') {
                        formula.exec_formula(formula.formula_ref[cell_full_name]);
                    }
                    else if (formula.formula_ref[cell_full_name].status === 'working') {
                        throw new Error('Circular ref');
                    }
                    row.push(sheet[cell_name].v);
                }
                else if (sheet[cell_name]) {
                    row.push(sheet[cell_name].v);
                }
                else {
                    row.push(null);
                }
            }
        }
        return matrix;
    };
};
