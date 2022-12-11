export { int_2_col_str } from './int_2_col_str';
export { col_str_2_int } from './col_str_2_int';
import { exec_formula } from './exec_formula';
import { find_all_cells_with_formulas } from './find_all_cells_with_formulas';
import { Calculator } from './Calculator';
import { int_2_col_str } from './int_2_col_str';

export default function mymodule(workbook, options?: any) {
    var formulas = find_all_cells_with_formulas(workbook, exec_formula);
    for (var i = formulas.length - 1; i >= 0; i--) {
        try {
            exec_formula(formulas[i]);
        } catch (error) {
            if (!options || !options.continue_after_error) {
                throw error
            }
            if (options.log_error) {
                console.log('error executing formula', 'sheet', formulas[i].sheet_name, 'cell', formulas[i].name, error)
            }
        }
    }
};

export function calculator(workbook) {
    return new Calculator(workbook, exec_formula);
};
mymodule.int_2_col_str = int_2_col_str;
mymodule.calculator = calculator;
export const setFx = mymodule.set_fx = exec_formula.set_fx;
export const execFx = mymodule.exec_fx = exec_formula.exec_fx;
export const importFunctions = mymodule.import_functions = exec_formula.import_functions;
export const importRawFunctions = mymodule.import_raw_functions = exec_formula.import_raw_functions;
export const xlsx_Fx = mymodule.xlsx_Fx = exec_formula.xlsx_Fx;
export const localizeFunctions = mymodule.localizeFunctions = exec_formula.localizeFunctions;


