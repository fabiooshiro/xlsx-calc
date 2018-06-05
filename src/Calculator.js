"use strict";

const RawValue = require('./RawValue.js');
const find_all_cells_with_formulas = require('./find_all_cells_with_formulas.js');

class Calculator {
    
    constructor(workbook, exec_formula) {
        this.workbook = workbook;
        this.expressions = [];
        this.exec_formula = exec_formula;
        this.variables = {};
        let formulas = find_all_cells_with_formulas(workbook, exec_formula);
        for (let i = formulas.length - 1; i >= 0; i--) {
            let exp = exec_formula.build_expression(formulas[i]);
            this.expressions.push(exp);
        }
    }
    
    setVar(var_name, value) {
        let variable = this.variables[var_name];
        if (variable) {
            variable.setValue(value);
        } else {
            this.expressions.forEach(exp => {
                this.setVarOfExpression(exp, var_name, value);
            });
        }
    }
    
    setVarOfExpression(exp, var_name, value) {
        for (let i = 0; i < exp.args.length; i++) {
            let arg = exp.args[i];
            if (arg === var_name) {
                exp.args[i] = this.variables[var_name] || (this.variables[var_name] = new RawValue(value));
            } else if (typeof arg === 'object' && (arg.name === 'Expression' || arg.name === 'UserFn')) {
                this.setVarOfExpression(arg, var_name, value);
            }
        }
    }
    
    execute() {
        this.expressions.forEach(exp => {
            exp.update_cell_value();
        });
    }
}

module.exports = Calculator;