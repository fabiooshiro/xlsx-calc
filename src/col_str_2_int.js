"use strict";

module.exports = function col_str_2_int(col_str) {
    var r = 0;
    var colstr = col_str.replace(/[0-9]+$/, '');
    for (var i = colstr.length; i--;) {
        r += Math.pow(26, colstr.length - i - 1) * (colstr.charCodeAt(i) - 64);
    }
    return r - 1;
};