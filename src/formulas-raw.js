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
        var cell_name = col_str + row;
        if (height === 1 && width === 1) {
            return new RefValue(cell_name, ref_value.formula).calc();
        }
        else {
            var end_range_col = int_2_col_str(col + width - 1);
            var end_range_row = row + height - 1;
            var end_range = end_range_col + end_range_row;
            var str_expression = parsed_ref.sheet_name + '!' + cell_name + ':' + end_range;
            return new Range(str_expression, ref_value.formula).calc();
        }
    }
}

function resolveOnErrorValue(onerrorvalue, resolve, reject) {
    let v_or_promise = onerrorvalue.calc();
    if (typeof v_or_promise === 'object' && typeof v_or_promise['then'] === 'function') {
        //console.log('resolvendo onerrorvalue');
        v_or_promise.then(r => {
            //console.log('On error value =', r);
            resolve(r);
        }).catch(e=> {
            //console.log('Erro no on error value');
            reject(e);
        });
    } else {
        //console.log('valor no caso de erro =', v_or_promise);
        resolve(v_or_promise);
    }
}

function iferror(cell_ref, onerrorvalue) {
    return new Promise((resolve, reject) => {
        try {
            cell_ref.calc().then(value=>{
                //console.log('tudo ok com o cell_ref...', typeof value);
                if (typeof value === 'undefined') {
                    resolveOnErrorValue(onerrorvalue, resolve, reject);
                } else if (typeof value === 'number' && (isNaN(value) || value === Infinity || value === -Infinity)) {
                    resolveOnErrorValue(onerrorvalue, resolve, reject);
                } else {
                    resolve(value);
                }
            }).catch(e => {
                //console.log('2 error level');
                resolveOnErrorValue(onerrorvalue, resolve, reject);
            });
        } catch(e) {
            return onerrorvalue.calc();
        }
    });
}

module.exports = {
    'OFFSET': raw_offset,
    'IFERROR': iferror
};
