(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory();
	else if(typeof define === 'function' && define.amd)
		define([], factory);
	else {
		var a = factory();
		for(var i in a) (typeof exports === 'object' ? exports : root)[i] = a[i];
	}
})(this, function() {
return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/index.js");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/Calculator.js":
/*!***************************!*\
  !*** ./src/Calculator.js ***!
  \***************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const RawValue = __webpack_require__(/*! ./RawValue.js */ "./src/RawValue.js");
const str_2_val = __webpack_require__(/*! ./str_2_val.js */ "./src/str_2_val.js");
const find_all_cells_with_formulas = __webpack_require__(/*! ./find_all_cells_with_formulas.js */ "./src/find_all_cells_with_formulas.js");

class Calculator {
    
    constructor(workbook, exec_formula) {
        this.workbook = workbook;
        this.expressions = [];
        this.exec_formula = exec_formula;
        this.variables = {};
        this.formulas = find_all_cells_with_formulas(workbook, exec_formula);
        for (let i = this.formulas.length - 1; i >= 0; i--) {
            let exp = exec_formula.build_expression(this.formulas[i]);
            this.expressions.push(exp);
        }
        this.calcNames();
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
    
    getVars() {
        let vars = {};
        for (let k in this.variables) {
            vars[k] = this.variables[k].calc();
        }
        return vars;
    }
    
    calcNames() {
        if (!this.workbook || !this.workbook.Workbook || !this.workbook.Workbook.Names) {
            return;
        }
        this.workbook.Workbook.Names.forEach(item => {
            let val = this.getRef(item.Ref);
            this.variables[item.Name] = val;
            this.expressions.forEach(exp => {
                this.setVarOfExpression(exp, item.Name);
            });
        });
    }
    
    getRef(ref_name) {
        if (!this.formulas.length) {
            throw new Error("No formula found.");
        }
        let first_formula = this.formulas[0];
        let formula_ref = first_formula.formula_ref;
        let formula = {
            formula_ref: formula_ref,
            wb: this.workbook,
            exec_formula: this.exec_formula
        };
        return str_2_val(ref_name, formula);
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

/***/ }),

/***/ "./src/Exp.js":
/*!********************!*\
  !*** ./src/Exp.js ***!
  \********************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const RawValue = __webpack_require__(/*! ./RawValue.js */ "./src/RawValue.js");
const RefValue = __webpack_require__(/*! ./RefValue.js */ "./src/RefValue.js");
const Range = __webpack_require__(/*! ./Range.js */ "./src/Range.js");
const str_2_val = __webpack_require__(/*! ./str_2_val.js */ "./src/str_2_val.js");

var exp_id = 0;

module.exports = function Exp(formula) {
    var self = this;
    self.id = ++exp_id;
    self.args = [];
    self.name = 'Expression';
    self.update_cell_value = update_cell_value;
    self.formula = formula;
    
    function update_cell_value() {
        try {
            formula.cell.v = self.calc();
            if (typeof(formula.cell.v) === 'string') {
                formula.cell.t = 's';
            }
            else if (typeof(formula.cell.v) === 'number') {
                formula.cell.t = 'n';
            }
        }
        catch (e) {
            var errorValues = {
                '#NULL!': 0x00,
                '#DIV/0!': 0x07,
                '#VALUE!': 0x0F,
                '#REF!': 0x17,
                '#NAME?': 0x1D,
                '#NUM!': 0x24,
                '#N/A': 0x2A,
                '#GETTING_DATA': 0x2B
            };
            if (errorValues[e.message] !== undefined) {
                formula.cell.t = 'e';
                formula.cell.w = e.message;
                formula.cell.v = errorValues[e.message];
            }
            else {
                throw e;
            }
        }
        finally {
            formula.status = 'done';
        }
    }
    
    function checkVariable(obj) {
        if (typeof obj.calc !== 'function') {
            throw new Error('Undefined ' + obj);
        }
    }
    
    function exec(op, args, fn) {
        for (var i = 0; i < args.length; i++) {
            if (args[i] === op) {
                try {
                    if (i===0 && op==='+') {
                        checkVariable(args[i + 1]);
                        var r = args[i + 1].calc();
                        args.splice(i, 2, new RawValue(r));
                    } else {
                        checkVariable(args[i - 1]);
                        checkVariable(args[i + 1]);
                        var r = fn(args[i - 1].calc(), args[i + 1].calc());
                        args.splice(i - 1, 3, new RawValue(r));
                        i--;
                    }
                }
                catch (e) {
                    // console.log('[Exp.js] - ' + formula.name + ': evaluating ' + formula.cell.f + '\n' + e.message);
                    throw e;
                }
            }
        }
    }

    function exec_minus(args) {
        for (var i = args.length; i--;) {
            if (args[i] === '-') {
                checkVariable(args[i + 1]);
                var r = -args[i + 1].calc();
                if (typeof args[i - 1] !== 'string' && i > 0) {
                    args.splice(i, 1, '+');
                    args.splice(i + 1, 1, new RawValue(r));
                }
                else {
                    args.splice(i, 2, new RawValue(r));
                }
            }
        }
    }

    self.calc = function() {
        let args = self.args.concat();
        exec_minus(args);
        exec('^', args, function(a, b) {
            return Math.pow(+a, +b);
        });
        exec('/', args, function(a, b) {
            if (b == 0) {
                throw Error('#DIV/0!');
            }
            return (+a) / (+b);
        });
        exec('*', args, function(a, b) {
            return (+a) * (+b);
        });
        exec('+', args, function(a, b) {
            return (+a) + (+b);
        });
        exec('&', args, function(a, b) {
            return '' + a + b;
        });
        exec('<', args, function(a, b) {
            return a < b;
        });
        exec('>', args, function(a, b) {
            return a > b;
        });
        exec('>=', args, function(a, b) {
            return a >= b;
        });
        exec('<=', args, function(a, b) {
            return a <= b;
        });
        exec('<>', args, function(a, b) {
            return a != b;
        });
        exec('=', args, function(a, b) {
            return a == b;
        });
        if (args.length == 1) {
            if (typeof(args[0].calc) !== 'function') {
                return args[0];
            }
            else {
                return args[0].calc();
            }
        }
    };

    var last_arg;
    self.push = function(buffer) {
        if (buffer) {
            var v = str_2_val(buffer, formula);
            if (((v === '=') && (last_arg == '>' || last_arg == '<')) || (last_arg == '<' && v === '>')) {
                self.args[self.args.length - 1] += v;
            }
            else {
                self.args.push(v);
            }
            last_arg = v;
            //console.log(self.id, '-->', v);
        }
    };
};

/***/ }),

/***/ "./src/Range.js":
/*!**********************!*\
  !*** ./src/Range.js ***!
  \**********************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const col_str_2_int = __webpack_require__(/*! ./col_str_2_int.js */ "./src/col_str_2_int.js");
const int_2_col_str = __webpack_require__(/*! ./int_2_col_str.js */ "./src/int_2_col_str.js");
const getSanitizedSheetName = __webpack_require__(/*! ./getSanitizedSheetName.js */ "./src/getSanitizedSheetName.js");

module.exports = function Range(str_expression, formula) {
    this.calc = function() {
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
        for (var i = min_row; i <= max_row; i++) {
            var row = [];
            matrix.push(row);
            for (var j = min_col; j <= max_col; j++) {
                var cell_name = int_2_col_str(j) + i;
                var cell_full_name = sheet_name + '!' + cell_name;
                if (formula.formula_ref[cell_full_name]) {
                    if (formula.formula_ref[cell_full_name].status === 'new') {
                        formula.exec_formula(formula.formula_ref[cell_full_name]);
                    }
                    else if (formula.formula_ref[cell_full_name].status === 'working') {
                        throw new Error('Circular ref');
                    }
                    row.push(sheet[cell_name].v);
                }
                else if (sheet[cell_name]) {
                    row.push(sheet[cell_name].v);
                }
                else {
                    row.push(null);
                }
            }
        }
        return matrix;
    };
};


/***/ }),

/***/ "./src/RawValue.js":
/*!*************************!*\
  !*** ./src/RawValue.js ***!
  \*************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


module.exports = function RawValue(value) {
    this.setValue = function(v) {
        value = v;
    };
    this.calc = function() {
        return value;
    };
};


/***/ }),

/***/ "./src/RefValue.js":
/*!*************************!*\
  !*** ./src/RefValue.js ***!
  \*************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const getSanitizedSheetName = __webpack_require__(/*! ./getSanitizedSheetName.js */ "./src/getSanitizedSheetName.js");

module.exports = function RefValue(str_expression, formula) {
    var self = this;
    this.name = 'RefValue';
    this.str_expression = str_expression;
    this.formula = formula;

    self.parseRef = function() {
        var sheet, sheet_name, cell_name, cell_full_name;
        if (str_expression.indexOf('!') != -1) {
            var aux = str_expression.split('!');
            sheet_name = getSanitizedSheetName(aux[0]);
            sheet = formula.wb.Sheets[sheet_name];
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
                if (ref_cell.t === 'e') {
                    console.log('ref is an error with new formula', cell_name);
                    throw new Error(ref_cell.w);
                }
                return ref_cell.v;
            }
            else if (formula_ref.status === 'working') {
                throw new Error('Circular ref');
            }
            else if (formula_ref.status === 'done') {
                if (ref_cell.t === 'e') {
                    console.log('ref is an error after formula eval');
                    throw new Error(ref_cell.w);
                }
                return ref_cell.v;
            }
        }
        else {
            if (ref_cell.t === 'e') {
                console.log('ref is an error with no formula', cell_name);
                throw new Error(ref_cell.w);
            }
            return ref_cell.v;
        }
    };
};


/***/ }),

/***/ "./src/UserFnExecutor.js":
/*!*******************************!*\
  !*** ./src/UserFnExecutor.js ***!
  \*******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


module.exports = function UserFnExecutor(user_function) {
    var self = this;
    self.name = 'UserFn';
    self.args = [];
    self.calc = function() {
        var errorValues = {
            '#NULL!': 0x00,
            '#DIV/0!': 0x07,
            '#VALUE!': 0x0F,
            '#REF!': 0x17,
            '#NAME?': 0x1D,
            '#NUM!': 0x24,
            '#N/A': 0x2A,
            '#GETTING_DATA': 0x2B
        }, result;
        try {
            result = user_function.apply(self, self.args.map(f=>f.calc()));
        } catch (e) {
            if (user_function.name === 'is_blank'
                && errorValues[e.message] !== undefined) {
                // is_blank applied to an error cell doesn't propagate the error
                result = 0;
            } else {
                throw e;
            }
        }
        return result;
    };
    self.push = function(buffer) {
        self.args.push(buffer);
    };
};

/***/ }),

/***/ "./src/UserRawFnExecutor.js":
/*!**********************************!*\
  !*** ./src/UserRawFnExecutor.js ***!
  \**********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


module.exports = function UserRawFnExecutor(user_function) {
    var self = this;
    self.name = 'UserRawFn';
    self.args = [];
    self.calc = function() {
        return user_function.apply(self, self.args);
    };
    self.push = function(buffer) {
        self.args.push(buffer);
    };
};


/***/ }),

/***/ "./src/col_str_2_int.js":
/*!******************************!*\
  !*** ./src/col_str_2_int.js ***!
  \******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


module.exports = function col_str_2_int(col_str) {
    var r = 0;
    var colstr = col_str.replace(/[0-9]+$/, '');
    for (var i = colstr.length; i--;) {
        r += Math.pow(26, colstr.length - i - 1) * (colstr.charCodeAt(i) - 64);
    }
    return r - 1;
};

/***/ }),

/***/ "./src/exec_formula.js":
/*!*****************************!*\
  !*** ./src/exec_formula.js ***!
  \*****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const Exp = __webpack_require__(/*! ./Exp.js */ "./src/Exp.js");
const RawValue = __webpack_require__(/*! ./RawValue.js */ "./src/RawValue.js");
const UserFnExecutor = __webpack_require__(/*! ./UserFnExecutor.js */ "./src/UserFnExecutor.js");
const UserRawFnExecutor = __webpack_require__(/*! ./UserRawFnExecutor.js */ "./src/UserRawFnExecutor.js");

var xlsx_Fx = {};
var xlsx_raw_Fx = {};

import_functions(__webpack_require__(/*! ./formulas.js */ "./src/formulas.js"));
import_raw_functions(__webpack_require__(/*! ./formulas-raw.js */ "./src/formulas-raw.js"));

const common_operations = {
    '*': 'multiply',
    '+': 'plus',
    '-': 'minus',
    '/': 'divide',
    '^': 'power',
    '&': 'concat',
    '<': 'lt',
    '>': 'gt',
    '=': 'eq'
};

function import_raw_functions(functions, opts) {
    for (var key in functions) {
        xlsx_raw_Fx[key] = functions[key];
    }
}

function import_functions(formulajs, opts) {
    opts = opts || {};
    var prefix = opts.prefix || '';
    for (var key in formulajs) {
        var obj = formulajs[key];
        if (typeof(obj) === 'function') {
            if (opts.override || !xlsx_Fx[prefix + key]) {
                xlsx_Fx[prefix + key] = obj;
            }
            // else {
            //     console.log(prefix + key, 'already exists.');
            //     console.log('  to override:');
            //     console.log('    XLSX_CALC.import_functions(yourlib, {override: true})');
            // }
        }
        else if (typeof(obj) === 'object') {
            import_functions(obj, my_assign(opts, { prefix: key + '.' }));
        }
    }
}

function my_assign(dest, source) {
    var obj = JSON.parse(JSON.stringify(dest));
    for (var k in source) {
        obj[k] = source[k];
    }
    return obj;
}

function build_expression(formula) {
    formula.status = 'working';
    var root_exp;
    var str_formula = formula.cell.f;
    if (str_formula[0] == '=') {
        str_formula = str_formula.substr(1);
    }
    var exp_obj = root_exp = new Exp(formula);
    var buffer = '',
        is_string = false,
        was_string = false;
    var fn_stack = [{
        exp: exp_obj
    }];
    for (var i = 0; i < str_formula.length; i++) {
        if (str_formula[i] == '"') {
            if (is_string) {
                exp_obj.push(new RawValue(buffer));
                is_string = false;
                was_string = true;
            }
            else {
                is_string = true;
            }
            buffer = '';
        }
        else if (is_string) {
            buffer += str_formula[i];
        }
        else if (str_formula[i] == '(') {
            var o, trim_buffer = buffer.trim(),
                special = xlsx_Fx[trim_buffer];
            var special_raw = xlsx_raw_Fx[trim_buffer];
            if (special_raw) {
                special = new UserRawFnExecutor(special_raw, formula);
            }
            else if (special) {
                special = new UserFnExecutor(special, formula);
            }
            else if (trim_buffer) {
                //Error: "Worksheet 1"!D145: Function INDEX not found
                throw new Error('"' + formula.sheet_name + '"!' + formula.name + ': Function ' + buffer + ' not found');
            }
            o = new Exp(formula);
            fn_stack.push({
                exp: o,
                special: special
            });
            exp_obj = o;
            buffer = '';
        }
        else if (common_operations[str_formula[i]]) {
            if (!was_string) {
                exp_obj.push(buffer);
            }
            was_string = false;
            exp_obj.push(str_formula[i]);
            buffer = '';
        }
        else if (str_formula[i] === ',' && fn_stack[fn_stack.length - 1].special) {
            was_string = false;
            fn_stack[fn_stack.length - 1].exp.push(buffer);
            fn_stack[fn_stack.length - 1].special.push(fn_stack[fn_stack.length - 1].exp);
            fn_stack[fn_stack.length - 1].exp = exp_obj = new Exp(formula);
            buffer = '';
        }
        else if (str_formula[i] == ')') {
            var v, stack = fn_stack.pop();
            exp_obj = stack.exp;
            exp_obj.push(buffer);
            v = exp_obj;
            buffer = '';
            exp_obj = fn_stack[fn_stack.length - 1].exp;
            if (stack.special) {
                stack.special.push(v);
                exp_obj.push(stack.special);
            }
            else {
                exp_obj.push(v);
            }
        }
        else {
            buffer += str_formula[i];
        }
    }
    root_exp.push(buffer);
    return root_exp;
}

function exec_formula(formula) {
    let root_exp = build_expression(formula);
    root_exp.update_cell_value();
}

exec_formula.set_fx = function set_fx(name, fn) {
    xlsx_Fx[name] = fn;
};

exec_formula.exec_fx = function exec_fx(name, args) {
    return xlsx_Fx[name].apply(this, args);
};

exec_formula.import_functions = import_functions;
exec_formula.import_raw_functions = import_raw_functions;
exec_formula.build_expression = build_expression;
exec_formula.xlsx_Fx = xlsx_Fx;
module.exports = exec_formula;


/***/ }),

/***/ "./src/find_all_cells_with_formulas.js":
/*!*********************************************!*\
  !*** ./src/find_all_cells_with_formulas.js ***!
  \*********************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

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


/***/ }),

/***/ "./src/formulas-raw.js":
/*!*****************************!*\
  !*** ./src/formulas-raw.js ***!
  \*****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const int_2_col_str = __webpack_require__(/*! ./int_2_col_str.js */ "./src/int_2_col_str.js");
const col_str_2_int = __webpack_require__(/*! ./col_str_2_int.js */ "./src/col_str_2_int.js");
const RawValue = __webpack_require__(/*! ./RawValue.js */ "./src/RawValue.js");
const Range = __webpack_require__(/*! ./Range.js */ "./src/Range.js");
const RefValue = __webpack_require__(/*! ./RefValue.js */ "./src/RefValue.js");

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

function iferror(cell_ref, onerrorvalue) {
    try {
        var value = cell_ref.calc();
        if (typeof value === 'number' && (isNaN(value) || value === Infinity || value === -Infinity)) {
            return onerrorvalue.calc();
        }
        return value;
    } catch(e) {
        return onerrorvalue.calc();
    }
}

function _if(condition, _then, _else) {
    if (condition.calc()) {
        return _then.calc();
    }
    else {
        return _else.calc();
    }
}

function and() {
    for (var i = 0; i < arguments.length; i++) {
        if(!arguments[i].calc()) return false;
    }
    return true;
}

module.exports = {
    'OFFSET': raw_offset,
    'IFERROR': iferror,
    'IF': _if,
    'AND': and
};


/***/ }),

/***/ "./src/formulas.js":
/*!*************************!*\
  !*** ./src/formulas.js ***!
  \*************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


// +---------------------+
// | FORMULAS REGISTERED |
// +---------------------+
let formulas = {
    'FLOOR': Math.floor,
    '_xlfn.FLOOR.MATH': Math.floor,
    'ABS': Math.abs,
    'SQRT': Math.sqrt,
    'VLOOKUP': vlookup,
    'MAX': max,
    'SUM': sum,
    'MIN': min,
    'CONCATENATE': concatenate,
    'PMT': pmt,
    'COUNTA': counta,
    'IRR': irr,
    'NORM.INV': normsInv,
    '_xlfn.NORM.INV': normsInv,
    'STDEV': stDeviation,
    'AVERAGE': avg,
    'EXP': EXP,
    'LN': Math.log,
    '_xlfn.VAR.P': var_p,
    'VAR.P': var_p,
    '_xlfn.COVARIANCE.P': covariance_p,
    'COVARIANCE.P': covariance_p,
    'TRIM': trim,
    'LEN': len,
    'ISBLANK': is_blank,
    'HLOOKUP': hlookup,
    'INDEX': index,
    'MATCH': match,
    'SUMPRODUCT': sumproduct,
    'ISNUMBER': isnumber
};

function isnumber(x) {
    return !isNaN(x);
}

function sumproduct() {
    var parseNumber = function (string) {
        if (string === undefined || string === '' || string === null) {
            return 0;
        }
        if (!isNaN(string)) {
            return parseFloat(string);
        }
        return 0;
    },
    consistentSizeRanges = function (matrixArray) {
        var getRowCount = function(matrix) {
                return matrix.length;
            },
            getColCount = function(matrix) {
                return matrix[0].length;
            },
            rowCount = getRowCount(matrixArray[0]),
            colCount = getColCount(matrixArray[0]);

        for (var i = 1; i < matrixArray.length; i++) {
            if (getRowCount(matrixArray[i]) !== rowCount
                || getColCount(matrixArray[i]) !== colCount) {
                return false;
            }
        }
        return true;
    };

    if (!arguments || arguments.length === 0) {
        throw Error('#VALUE!');
    }
    if (!consistentSizeRanges(arguments)) {
        throw Error('#VALUE!');
    }

    var arrays = arguments.length + 1;
    var result = 0;
    var product;
    var k;
    var _i;
    var _ij;
    for (var i = 0; i < arguments[0].length; i++) {
        if (!(arguments[0][i] instanceof Array)) {
            product = 1;
            for (k = 1; k < arrays; k++) {
                _i = parseNumber(arguments[k - 1][i]);
                
                product *= _i;
            }
            result += product;
        } else {
            for (var j = 0; j < arguments[0][i].length; j++) {
                product = 1;
                for (k = 1; k < arrays; k++) {
                    _ij = parseNumber(arguments[k - 1][i][j]);
                    
                    product *= _ij;
                }
                result += product;
            }
        }
    }
    return result;
}

function match(lookupValue, matrix, matchType) {
    if (Array.isArray(matrix) 
        && matrix.length === 1
        && Array.isArray(matrix[0])) {
        matrix = matrix[0];
    }
    if (!lookupValue && !matrix) {
        throw Error('#N/A');
    }
    
    if (arguments.length === 2) {
        matchType = 1;
    }
    if (!(matrix instanceof Array)) {
        throw Error('#N/A');
    }

    if (matchType !== -1 && matchType !== 0 && matchType !== 1) {
        throw Error('#N/A');
    }
    var index;
    var indexValue;
    for (var idx = 0; idx < matrix.length; idx++) {
        if (matchType === 1) {
            if (matrix[idx] === lookupValue) {
                return idx + 1;
            } else if (matrix[idx] < lookupValue) {
                if (!indexValue) {
                    index = idx + 1;
                    indexValue = matrix[idx];
                } else if (matrix[idx] > indexValue) {
                    index = idx + 1;
                    indexValue = matrix[idx];
                }
            }
        } else if (matchType === 0) {
            if (typeof lookupValue === 'string') {
                lookupValue = lookupValue.replace(/\?/g, '.');

                if (Array.isArray(matrix[idx])) {
                    if (matrix[idx].length === 1
                        && typeof matrix[idx][0] === 'string') {
                            if (matrix[idx][0].toLowerCase() === lookupValue.toLowerCase()) {
                                return idx + 1;
                            }
                        } 
                } else if (typeof matrix[idx] === 'string') {
                    if (matrix[idx].toLowerCase() === lookupValue.toLowerCase()) {
                        return idx + 1;
                    }
                }
            } else {
                if (matrix[idx] === lookupValue) {
                    return idx + 1;
                }
            }
        } else if (matchType === -1) {
            if (matrix[idx] === lookupValue) {
                return idx + 1;
            } else if (matrix[idx] > lookupValue) {
                if (!indexValue) {
                    index = idx + 1;
                    indexValue = matrix[idx];
                } else if (matrix[idx] < indexValue) {
                    index = idx + 1;
                    indexValue = matrix[idx];
                }
            }
        }
    }
    if (!index ) {
        throw Error('#N/A');
    }
    return index;
}

function index(matrix, row_num, column_num) {
    if (row_num <= matrix.length) {
        var row = matrix[row_num - 1];
        if (Array.isArray(row)) {
            if (!column_num) {
                return row;
            } else if (column_num <= row.length) {
                return row[column_num - 1];
            }
        } else {
            return matrix[row_num];
        }
    }
    throw Error('#REF!');
}

// impl ported from https://github.com/FormulaPages/hlookup
function hlookup(needle, table, index, exactmatch) {
    if (typeof needle === "undefined" || (0, is_blank)(needle)) {
        return null;
    }

    index = index || 0;
    let row = table[0];

    for (let i = 0; i < row.length; i++) {
        if (exactmatch && row[i] === needle || row[i].toLowerCase().indexOf(needle.toLowerCase()) !== -1) {
            return index < table.length + 1 ? table[index - 1][i] : table[0][i];
        }
    }

    throw Error('#N/A');
}

function len(a) {
    return ('' + a).length;
}

function trim(a) {
    return ('' + a).trim();
}

function is_blank(a) {
    return !a;
}

function covariance_p(a, b) {
    a = getArrayOfNumbers(a);
    b = getArrayOfNumbers(b);
    if (a.length != b.length) {
        return 'N/D';
    }
    var inv_n = 1.0 / a.length;
    var avg_a = sum.apply(this, a) / a.length;
    var avg_b = sum.apply(this, b) / b.length;
    var s = 0.0;
    for (var i = 0; i < a.length; i++) {
        s += (a[i] - avg_a) * (b[i] - avg_b);
    }
    return s * inv_n;
}

function getArrayOfNumbers(range) {
    var arr = [];
    for (var i = 0; i < range.length; i++) {
        var arg = range[i];
        if (Array.isArray(arg)) {
            var matrix = arg;
            for (var j = matrix.length; j--;) {
                if (typeof(matrix[j]) == 'number') {
                    arr.push(matrix[j]);
                }
                else if (Array.isArray(matrix[j])) {
                    for (var k = matrix[j].length; k--;) {
                        if (typeof(matrix[j][k]) == 'number') {
                            arr.push(matrix[j][k]);
                        }
                    }
                }
                // else {
                //   wtf is that?
                // }
            }
        }
        else {
            if (typeof(arg) == 'number') {
                arr.push(arg);
            }
        }
    }
    return arr;
}

function var_p() {
    var average = avg.apply(this, arguments);
    var s = 0.0;
    var c = 0;
    for (var i = 0; i < arguments.length; i++) {
        var arg = arguments[i];
        if (Array.isArray(arg)) {
            var matrix = arg;
            for (var j = matrix.length; j--;) {
                for (var k = matrix[j].length; k--;) {
                    if (matrix[j][k] !== null && matrix[j][k] !== undefined) {
                        s += Math.pow(matrix[j][k] - average, 2);
                        c++;
                    }
                }
            }
        }
        else {
            s += Math.pow(arg - average, 2);
            c++;
        }
    }
    return s / c;
}

function EXP(n) {
    return Math.pow(Math.E, n);
}

function avg() {
    return sum.apply(this, arguments) / counta.apply(this, arguments);
}

function stDeviation() {
    var array = getArrayOfNumbers(arguments);

    function _mean(array) {
        return array.reduce(function(a, b) {
            return a + b;
        }) / array.length;
    }
    var mean = _mean(array),
        dev = array.map(function(itm) {
            return (itm - mean) * (itm - mean);
        });
    return Math.sqrt(dev.reduce(function(a, b) {
        return a + b;
    }) / (array.length - 1));
}

/// Original C++ implementation found at http://www.wilmott.com/messageview.cfm?catid=10&threadid=38771
/// C# implementation found at http://weblogs.asp.net/esanchez/archive/2010/07/29/a-quick-and-dirty-implementation-of-excel-norminv-function-in-c.aspx
/*
 *     Compute the quantile function for the normal distribution.
 *
 *     For small to moderate probabilities, algorithm referenced
 *     below is used to obtain an initial approximation which is
 *     polished with a final Newton step.
 *
 *     For very large arguments, an algorithm of Wichura is used.
 *
 *  REFERENCE
 *
 *     Beasley, J. D. and S. G. Springer (1977).
 *     Algorithm AS 111: The percentage points of the normal distribution,
 *     Applied Statistics, 26, 118-121.
 *
 *      Wichura, M.J. (1988).
 *      Algorithm AS 241: The Percentage Points of the Normal Distribution.
 *      Applied Statistics, 37, 477-484.
 */
function normsInv(p, mu, sigma) {
    if (p < 0 || p > 1) {
        throw "The probality p must be bigger than 0 and smaller than 1";
    }
    if (sigma < 0) {
        throw "The standard deviation sigma must be positive";
    }

    if (p == 0) {
        return -Infinity;
    }
    if (p == 1) {
        return Infinity;
    }
    if (sigma == 0) {
        return mu;
    }

    var q, r, val;

    q = p - 0.5;

    /*-- use AS 241 --- */
    /* double ppnd16_(double *p, long *ifault)*/
    /*      ALGORITHM AS241  APPL. STATIST. (1988) VOL. 37, NO. 3
            Produces the normal deviate Z corresponding to a given lower
            tail area of P; Z is accurate to about 1 part in 10**16.
    */
    if (Math.abs(q) <= .425) { /* 0.075 <= p <= 0.925 */
        r = .180625 - q * q;
        val =
            q * (((((((r * 2509.0809287301226727 +
                            33430.575583588128105) * r + 67265.770927008700853) * r +
                        45921.953931549871457) * r + 13731.693765509461125) * r +
                    1971.5909503065514427) * r + 133.14166789178437745) * r +
                3.387132872796366608) / (((((((r * 5226.495278852854561 +
                        28729.085735721942674) * r + 39307.89580009271061) * r +
                    21213.794301586595867) * r + 5394.1960214247511077) * r +
                687.1870074920579083) * r + 42.313330701600911252) * r + 1);
    }
    else { /* closer than 0.075 from {0,1} boundary */

        /* r = min(p, 1-p) < 0.075 */
        if (q > 0)
            r = 1 - p;
        else
            r = p;

        r = Math.sqrt(-Math.log(r));
        /* r = sqrt(-log(r))  <==>  min(p, 1-p) = exp( - r^2 ) */

        if (r <= 5) { /* <==> min(p,1-p) >= exp(-25) ~= 1.3888e-11 */
            r += -1.6;
            val = (((((((r * 7.7454501427834140764e-4 +
                                .0227238449892691845833) * r + .24178072517745061177) *
                            r + 1.27045825245236838258) * r +
                        3.64784832476320460504) * r + 5.7694972214606914055) *
                    r + 4.6303378461565452959) * r +
                1.42343711074968357734) / (((((((r *
                                1.05075007164441684324e-9 + 5.475938084995344946e-4) *
                            r + .0151986665636164571966) * r +
                        .14810397642748007459) * r + .68976733498510000455) *
                    r + 1.6763848301838038494) * r +
                2.05319162663775882187) * r + 1);
        }
        else { /* very close to  0 or 1 */
            r += -5;
            val = (((((((r * 2.01033439929228813265e-7 +
                                2.71155556874348757815e-5) * r +
                            .0012426609473880784386) * r + .026532189526576123093) *
                        r + .29656057182850489123) * r +
                    1.7848265399172913358) * r + 5.4637849111641143699) *
                r + 6.6579046435011037772) / (((((((r *
                            2.04426310338993978564e-15 + 1.4215117583164458887e-7) *
                        r + 1.8463183175100546818e-5) * r +
                    7.868691311456132591e-4) * r + .0148753612908506148525) * r + .13692988092273580531) * r +
                .59983220655588793769) * r + 1);
        }

        if (q < 0.0) {
            val = -val;
        }
    }

    return mu + sigma * val;
}

function irr(range, guess) {
    var min = -2.0;
    var max = 1.0;
    var n = 0;
    do {
        var guest = (min + max) / 2;
        var NPV = 0;
        for (var i = 0; i < range.length; i++) {
            var arg = range[i];
            NPV += arg[0] / Math.pow((1 + guest), i);
        }
        if (NPV > 0) {
            if (min === max) {
                max += Math.abs(guest);
            }
            min = guest;
        }
        else {
            max = guest;
        }
        n++;
    } while (Math.abs(NPV) > 0.000001 && n < 100000);
    //console.log(n);
    return guest;
}

function counta() {
    var r = 0;
    for (var i = arguments.length; i--;) {
        var arg = arguments[i];
        if (Array.isArray(arg)) {
            var matrix = arg;
            for (var j = matrix.length; j--;) {
                for (var k = matrix[j].length; k--;) {
                    if (matrix[j][k] !== null && matrix[j][k] !== undefined) {
                        r++;
                    }
                }
            }
        }
        else {
            if (arg !== null && arg !== undefined) {
                r++;
            }
        }
    }
    return r;
}

function pmt(rate_per_period, number_of_payments, present_value, future_value, type) {
    type = type || 0;
    future_value = future_value || 0;
    if (rate_per_period != 0.0) {
        // Interest rate exists
        var q = Math.pow(1 + rate_per_period, number_of_payments);
        return -(rate_per_period * (future_value + (q * present_value))) / ((-1 + q) * (1 + rate_per_period * (type)));

    }
    else if (number_of_payments != 0.0) {
        // No interest rate, but number of payments exists
        return -(future_value + present_value) / number_of_payments;
    }
    return 0;
}

function concatenate() {
    var r = '';
    for (var i = 0; i < arguments.length; i++) {
        var arg = arguments[i];
        if (arg === null || arg === undefined) continue;
        r += arg;
    }
    return r;
}

function sum() {
    var r = 0;
    for (var i = arguments.length; i--;) {
        var arg = arguments[i];
        if (Array.isArray(arg)) {
            var matrix = arg;
            for (var j = matrix.length; j--;) {
                for (var k = matrix[j].length; k--;) {
                    if (!isNaN(matrix[j][k])) {
                        r += +matrix[j][k];
                    }
                }
            }
        }
        else {
            r += +arg;
        }
    }
    return r;
}

function max() {
    var max = null;
    for (var i = arguments.length; i--;) {
        var arg = arguments[i];
        if (Array.isArray(arg)) {
            var arr = arg;
            for (var j = arr.length; j--;) {
                max = max == null || max < arr[j] ? arr[j] : max;
            }
        }
        else if (!isNaN(arg)) {
            max = max == null || max < arg ? arg : max;
        }
        else {
            console.log('WTF??', arg);
        }
    }
    return max;
}

function min() {
    var result = null;
    for (var i = arguments.length; i--;) {
        var arg = arguments[i];
        if (Array.isArray(arg)) {
            var arr = arg;
            for (var j = arr.length; j--;) {
                result = result == null || result > arr[j] ? arr[j] : result;
            }
        }
        else if (!isNaN(arg)) {
            result = result == null || result > arg ? arg : result;
        }
        else {
            console.log('WTF??', arg);
        }
    }
    return result;
}

function vlookup(key, matrix, return_index) {
    for (var i = 0; i < matrix.length; i++) {
        if (matrix[i][0] == key) {
            return matrix[i][return_index - 1];
        }
    }
    throw Error('#N/A');
}

module.exports = formulas;


/***/ }),

/***/ "./src/getSanitizedSheetName.js":
/*!**************************************!*\
  !*** ./src/getSanitizedSheetName.js ***!
  \**************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


module.exports = function getSanitizedSheetName(sheet_name) {
    var quotedMatch = sheet_name.match(/^'(.*)'$/);
    if (quotedMatch) {
        return quotedMatch[1];
    }
    else {
        return sheet_name;
    }
};


/***/ }),

/***/ "./src/index.js":
/*!**********************!*\
  !*** ./src/index.js ***!
  \**********************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


const int_2_col_str = __webpack_require__(/*! ./int_2_col_str.js */ "./src/int_2_col_str.js");
const col_str_2_int = __webpack_require__(/*! ./col_str_2_int.js */ "./src/col_str_2_int.js");
const exec_formula = __webpack_require__(/*! ./exec_formula.js */ "./src/exec_formula.js");
const find_all_cells_with_formulas = __webpack_require__(/*! ./find_all_cells_with_formulas.js */ "./src/find_all_cells_with_formulas.js");
const Calculator = __webpack_require__(/*! ./Calculator.js */ "./src/Calculator.js");

var mymodule = function(workbook) {
    var formulas = find_all_cells_with_formulas(workbook, exec_formula);
    for (var i = formulas.length - 1; i >= 0; i--) {
        exec_formula(formulas[i]);
    }
};

mymodule.calculator = function calculator(workbook) {
    return new Calculator(workbook, exec_formula);
};

mymodule.set_fx = exec_formula.set_fx;
mymodule.exec_fx = exec_formula.exec_fx;
mymodule.col_str_2_int = col_str_2_int;
mymodule.int_2_col_str = int_2_col_str;
mymodule.import_functions = exec_formula.import_functions;
mymodule.import_raw_functions = exec_formula.import_raw_functions;
mymodule.xlsx_Fx = exec_formula.xlsx_Fx;



module.exports = mymodule;

/***/ }),

/***/ "./src/int_2_col_str.js":
/*!******************************!*\
  !*** ./src/int_2_col_str.js ***!
  \******************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

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

/***/ }),

/***/ "./src/str_2_val.js":
/*!**************************!*\
  !*** ./src/str_2_val.js ***!
  \**************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

const RawValue = __webpack_require__(/*! ./RawValue.js */ "./src/RawValue.js");
const RefValue = __webpack_require__(/*! ./RefValue.js */ "./src/RefValue.js");
const Range = __webpack_require__(/*! ./Range.js */ "./src/Range.js");

module.exports = function str_2_val(buffer, formula) {
    var v;
    if (!isNaN(buffer)) {
        v = new RawValue(+buffer);
    }
    else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/)) {
        v = new Range(buffer.trim().replace(/\$/g, ''), formula);
    }
    else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[^!]+![A-Z]+[0-9]+:[A-Z]+[0-9]+$/)) {
        v = new Range(buffer.trim().replace(/\$/g, ''), formula);
    }
    else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[A-Z]+:[A-Z]+$/)) {
        v = new Range(buffer.trim().replace(/\$/g, ''), formula);
    }
    else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[^!]+![A-Z]+:[A-Z]+$/)) {
        v = new Range(buffer.trim().replace(/\$/g, ''), formula);
    }
    else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[A-Z]+[0-9]+$/)) {
        v = new RefValue(buffer.trim().replace(/\$/g, ''), formula);
    }
    else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[^!]+![A-Z]+[0-9]+$/)) {
        v = new RefValue(buffer.trim().replace(/\$/g, ''), formula);
    }
    else if (typeof buffer === 'string' && !isNaN(buffer.trim().replace(/%$/, ''))) {
        v = new RawValue(+(buffer.trim().replace(/%$/, '')) / 100.0);
    }
    else {
        v = buffer;
    }
    return v;
};

/***/ })

/******/ });
});
//# sourceMappingURL=xlsx-calc.js.map