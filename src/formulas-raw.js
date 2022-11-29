"use strict";

const int_2_col_str = require('./int_2_col_str.js');
const col_str_2_int = require('./col_str_2_int.js');
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
    if (condition.calc()) {
        // console.log(condition.formula.name)
        // if (condition.formula.name === 'P40') {
        //     console.log('P40 =', _then.calc());
        //     console.log(' -->', _then.args[1].calc());
        // }
        return _then.calc();
    }
    else {
        if (typeof _else === 'undefined') {
            return false;
        } else {
            return _else.calc();
        }
    }
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

function filter(range, condition) {
    let data = range.calc();
    let conditions = condition.calc();
    let cellName = range.formula.name;
    let colAndRow = cellName.match(/([A-Z]+)([0-9]+)/);
    let sheet = range.formula.sheet;
    let colNumber = col_str_2_int(colAndRow[1]);
    let rowNumber = +colAndRow[2];

    let returnValue = sheet[cellName].v;
    for (let i = 0; i < conditions[0].length; i++) {
        if (conditions[0][i]) {
            for (let row = 0; row < data.length; row++) {
                let destinationColumn = colNumber + i;
                let destinationRow = rowNumber + row;
                let destinationCellName = int_2_col_str(destinationColumn) + destinationRow;

                if (sheet[destinationCellName]) {
                    sheet[destinationCellName].v = data[row][i];
                    if (destinationCellName === cellName) {
                        returnValue = data[row][i];
                    }
                } else {
                    sheet[destinationCellName] = { v: data[row][i] };
                }
            }
        }
    }
    return returnValue;
}

module.exports = {
    'OFFSET': raw_offset,
    'IFERROR': iferror,
    'IF': _if,
    'AND': and,
    'OR': _or,
    'TRANSPOSE': transpose,
    'FILTER': filter
};
