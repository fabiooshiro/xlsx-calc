"use strict";

module.exports = function RefValue(str_expression, formula) {
    var self = this;
    this.name = 'RefValue';
    this.str_expression = str_expression;
    this.formula = formula;

    self.parseRef = function() {
        var sheet, sheet_name, cell_name, cell_full_name;
        if (str_expression.indexOf('!') != -1) {
            var aux = str_expression.split('!');
            sheet = formula.wb.Sheets[aux[0]];
            if (!sheet) {
                var quoted = aux[0].match(/^'(.*)'$/);
                if (quoted) {
                    aux[0] = quoted[1];
                }
                sheet = formula.wb.Sheets[aux[0]];
            }
            sheet_name = aux[0];
            cell_name = aux[1];
        }
        else {
            sheet = formula.sheet;
            sheet_name = formula.sheet_name;
            cell_name = str_expression;
        }
        if (!sheet) {
            throw Error("Sheet " + sheet_name + " not found.");
        }
        cell_full_name = sheet_name + '!' + cell_name;
        return {
            sheet: sheet,
            sheet_name: sheet_name,
            cell_name: cell_name,
            cell_full_name: cell_full_name
        };
    };

    this.calc = function() {
        var resolved_ref = self.parseRef();
        var sheet = resolved_ref.sheet;
        var cell_name = resolved_ref.cell_name;
        var cell_full_name = resolved_ref.cell_full_name;
        var ref_cell = sheet[cell_name];
        if (!ref_cell) {
            return null;
        }
        var formula_ref = formula.formula_ref[cell_full_name];
        if (formula_ref) {
            if (formula_ref.status === 'new') {
                formula.exec_formula(formula_ref);
                return sheet[cell_name].v;
            }
            else if (formula_ref.status === 'working') {
                throw new Error('Circular ref');
            }
            else if (formula_ref.status === 'done') {
                return sheet[cell_name].v;
            }
        }
        else {
            return sheet[cell_name].v;
        }
    };
};
