"use strict";

const col_str_2_int = require('./col_str_2_int.js');
const int_2_col_str = require('./int_2_col_str.js');
const getSanitizedSheetName = require('./getSanitizedSheetName.js');

module.exports = function Range(str_expression, formula) {
    
    function promiseInSeq(sheet, matrix, min_row, max_row, min_col, max_col, sheet_name, resolve, reject, _row) {
        //console.log('min_row =', min_row, 'max_row =', max_row);
        //console.log('min_col =', min_col, 'max_col =', max_col);
        for (let i = min_row; i <= max_row; i++) {
            //console.log('ok i =', i);
            let row;
            if (_row) {
                row = _row;
            }
            else {
                row = [];
                matrix.push(row);
            }
            for (let j = min_col; j <= max_col; j++) {
                let cell_name = int_2_col_str(j) + i;
                let cell_full_name = sheet_name + '!' + cell_name;
                //console.log('range <<', cell_name, 'i =', i, 'j =', j);
                if (formula.formula_ref[cell_full_name]) {
                    if (formula.formula_ref[cell_full_name].status === 'working') {
                        //console.log('Circular ref in range');
                        reject('Circular ref');
                        return;
                    } else if (formula.formula_ref[cell_full_name].status === 'new') {
                        formula.exec_formula(formula.formula_ref[cell_full_name]).then(r=>{
                            row.push(sheet[cell_name].v);
                            //console.log('recursao asincrona');
                            j++;
                            if (j > max_col) {
                                j = 0;
                                i++;
                            }
                            promiseInSeq(sheet, matrix, i, max_row, j, max_col, sheet_name, resolve, reject, row);
                        }).catch(reject);
                        return;
                    } else if (formula.formula_ref[cell_full_name].status === 'done') {
                        row.push(sheet[cell_name].v);
                    }
                }
                else if (sheet[cell_name]) {
                    row.push(sheet[cell_name].v);
                }
                else {
                    row.push(null);
                }
            }
        }
        resolve();
    }
    
    this.calc = function() {
        return new Promise((resolve, reject) => {
            try {
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
                new Promise((res, rej) => {
                    promiseInSeq(sheet, matrix, min_row, max_row, min_col, max_col, sheet_name, res, rej);
                }).then(() => {
                    resolve(matrix);
                }).catch(reject);
            } catch(e) {
                reject(e);
            }
        });
    };
};
