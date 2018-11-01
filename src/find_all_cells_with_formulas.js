"use strict";

module.exports = function find_all_cells_with_formulas(wb, exec_formula) {
    let formula_ref = {};
    let cells = [];
    for (let sheet_name in wb.Sheets) {
        let sheet = wb.Sheets[sheet_name];
        for (let cell_name in sheet) {
            if (sheet[cell_name] && sheet[cell_name].f) {
                let formula = formula_ref[sheet_name + '!' + cell_name] = {
                    formula_ref: formula_ref,
                    wb: wb,
                    sheet: sheet,
                    sheet_name: sheet_name,
                    cell: sheet[cell_name],
                    name: cell_name,
                    status: 'new',
                    exec_formula: exec_formula
                };
                cells.push(formula);
            }
        }
    }
    return cells;
};
