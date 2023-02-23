"use strict";

const int_2_col_str = require('./int_2_col_str.js');
const col_str_2_int = require('./col_str_2_int.js');
const dynamic_array_compatible = require('./dynamic_array_compatible.js');
const RawValue = require('./RawValue.js');
const Range = require('./Range.js');
const RefValue = require('./RefValue.js');

function raw_offset(cell_ref, rows, columns, height, width) {
    height = (height || new RawValue(1)).calc();
    width = (width || new RawValue(1)).calc();
    if (cell_ref.args.length === 1 && cell_ref.args[0].name === 'RefValue') {
        var ref_value = cell_ref.args[0];
        var parsed_ref = ref_value.parseRef();
        var col = col_str_2_int(parsed_ref.cell_name) + columns.calc();
        var col_str = int_2_col_str(col);
        var row = +parsed_ref.cell_name.replace(/^[A-Z]+/g, '') + rows.calc();
        var cell_name = parsed_ref.sheet_name + '!' + col_str + row;
        if (height === 1 && width === 1) {
            return new RefValue(cell_name, ref_value.formula).calc();
        }
        else {
            var end_range_col = int_2_col_str(col + width - 1);
            var end_range_row = row + height - 1;
            var end_range = end_range_col + end_range_row;
            var str_expression = cell_name + ':' + end_range;
            return new Range(str_expression, ref_value.formula).calc();
        }
    }
}

function iferror(cell_ref, onerrorvalue) {
    try {
        var value = cell_ref.calc();
        if (typeof value === 'number' && (isNaN(value) || value === Infinity || value === -Infinity)) {
            return onerrorvalue.calc();
        }
        return value;
    } catch (e) {
        return onerrorvalue.calc();
    }
}

function _if(condition, _then, _else) {
    var condition_results;
    var then_results;
    var else_results;
    try {
        condition_results = condition.calc();
    } catch (e) {
        condition_results = e;
    }
    try {
        then_results = _then.calc();
    } catch (e) {
        then_results = e;
    }
    try {
        else_results = typeof _else === 'undefined' ? false : _else.calc();
    } catch (e) {
        else_results = e;
    }
    return dynamic_array_compatible(function (condition_result, then_result, else_result) {
        if (condition_result instanceof Error) {
            return condition_result;
        }
        return condition_result ? then_result : else_result;
    })(condition_results, then_results, else_results);
}

function and() {
    for (var i = 0; i < arguments.length; i++) {
        if (!arguments[i].calc()) return false;
    }
    return true;
}

function _or() {
    for (var i = 0; i < arguments.length; i++) {
        if (arguments[i].calc()) return true;
    }
    return false;
}

function transpose(expressionWithRange) {
    let range = expressionWithRange.args[0];
    // console.log(expressionWithRange.args[0])
    // console.log(expressionWithRange.formula.wb.Sheets.Sheet1)
    // console.log(range.calc())
    let matrix = range.calc();
    let cellName = expressionWithRange.formula.name;
    let colRow = cellName.match(/([A-Z]+)([0-9]+)/);
    let sheet = expressionWithRange.formula.sheet;
    // console.log(colRow[1], colRow[2]);
    // console.log(col_str_2_int(colRow[1]));
    let colNumber = col_str_2_int(colRow[1]);
    let rowNumber = +colRow[2];
    for (let i = 0; i < matrix.length; i++) {
        let matrixRow = matrix[i];
        for (let j = 0; j < matrixRow.length; j++) {
            let destinationColumn = colNumber + i;
            let destinationRow = rowNumber + j;
            let value = matrixRow[j];
            // console.log(int_2_col_str(destinationColumn), destinationRow, value);
            sheet[int_2_col_str(destinationColumn) + destinationRow].v = value;
        }
    }
    // console.log(expressionWithRange.formula.name)
    return matrix[0][0];
}

module.exports = {
    'OFFSET': raw_offset,
    'IFERROR': iferror,
    'IF': _if,
    'AND': and,
    'OR': _or,
    'TRANSPOSE': transpose,
};
