"use strict";

module.exports = function int_2_col_str(n) {
    var dividend = n + 1;
    var columnName = '';
    var modulo;
    var guard = 10;
    while (dividend > 0 && guard--) {
        modulo = (dividend - 1) % 26;
        columnName = String.fromCharCode(modulo + 65) + columnName;
        dividend = (dividend - modulo - 1) / 26;
    }
    return columnName;
};