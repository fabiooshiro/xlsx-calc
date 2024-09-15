"use strict";

const LRUCache = require('./LRUCache.js');
const col_str_2_int = require('./col_str_2_int.js');
const int_2_col_str = require('./int_2_col_str.js');
const getSanitizedSheetName = require('./getSanitizedSheetName.js');

const Cache = new LRUCache()

function Range(str_expression, formula) {
    this.parse = function() {
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
            str_max_row = (sheet['!ref'].includes(':') ? sheet['!ref'].split(':')[1] : sheet['!ref']).replace(/^[A-Z]+/, '');
        }
        // the max is 1048576, but TLE
        max_row = parseInt(str_max_row == '' ? '500000' : str_max_row, 10);
        var min_col = col_str_2_int(arr[0]);
        var max_col = col_str_2_int(arr[1]);
        return {
            sheet_name: sheet_name,
            sheet: sheet,
            min_row: min_row,
            min_col: min_col,
            max_row: max_row,
            max_col: max_col,
        };
    };

    this._calc = function() {
        var results = this.parse();
        var sheet_name = results.sheet_name;
        var sheet = results.sheet;
        var min_row = results.min_row;
        var min_col = results.min_col;
        var max_row = results.max_row;
        var max_col = results.max_col;
        var matrix = [];
        for (var i = min_row; i <= max_row; i++) {
            var row = [];
            matrix.push(row);
            for (var j = min_col; j <= max_col; j++) {
                var cell_name = int_2_col_str(j) + i;
                var cell_full_name = sheet_name + '!' + cell_name;
                var formula_ref = formula.formula_ref[cell_full_name];
                if (formula_ref) {
                    if (formula_ref.status === 'new') {
                        formula.exec_formula(formula_ref);
                    } else if (formula_ref.status === 'working') {
                        if (formula_ref.cell.f.includes(formula.name)) {
                            throw new Error('Circular ref');
                        }
                        formula.exec_formula(formula_ref);
                    }
                    if (sheet[cell_name].t === 'e') {
                        row.push(new Error(sheet[cell_name].w));
                    }
                    else {
                        row.push(sheet[cell_name].v);
                    }
                }
                else if (sheet[cell_name]) {
                    if (sheet[cell_name].t === 'e') {
                        row.push(new Error(sheet[cell_name].w));
                    }
                    else {
                        row.push(sheet[cell_name].v);
                    }
                }
                else {
                    row.push(null);
                }
            }
        }
        return matrix;
    };

    this.calc = function() {
        const cached = Cache.get(str_expression);
        if(cached) {
            return cached;
        }
        else {
            const result =  this._calc();
            Cache.set(str_expression, result);
            return result;
        }
    }
};

Range.cache = Cache

module.exports = Range;
