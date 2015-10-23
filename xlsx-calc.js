"use strict";

(function() {

  var xlsx_Fx = {
    'FLOOR': Math.floor,
    '_xlfn.FLOOR.MATH': Math.floor,
    'ABS': Math.abs,
    'SQRT': Math.sqrt,
    'VLOOKUP': vlookup,
    'MAX': max,
    'SUM': sum,
    'MIN': min,
    'CONCATENATE': concatenate,
    'IF': _if
  };
  
  function _if(condition, _then, _else) {
    console.log(eval(condition));
    if(eval(condition)) {
      return _then;
    } else {
      return _else;
    }
  }
  
  function concatenate() {
    var r = '';
    for (var i = 0; i < arguments.length; i++) {
      r += arguments[i];
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
            r += +matrix[j][k];
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

  var mymodule = function(workbook) {
    var formulas = find_all_cells_with_formulas(workbook);
    for (var i = formulas.length - 1; i >= 0; i--) {
      exec_formula(formulas[i]);
    }
  };
  
  mymodule.set_fx = function(name, fn) {
    xlsx_Fx[name] = fn;
  };

  function UserFnExecutor(user_function, formula) {
    var self = this;
    self.name = 'UserFn';
    self.args = [];
    self.calc = function() {
      return user_function.apply(self, self.args);
    };
    self.push = function(buffer) {
      if (buffer) {
        //console.log('pushing', buffer, 'into', self.name);
        self.args.push(buffer.calc());
      }
    };
  }

  function RawValue(value) {
    this.calc = function() {
      return value;
    };
  }

  function RefValue(str_expression, formula) {
    this.calc = function() {
      var ref_cell = formula.sheet[str_expression];
      if (!ref_cell) {
        throw Error("Cell " + str_expression + " not found.");
      }
      var formula_ref = formula.formula_ref[str_expression];
      if (formula_ref) {
        if (formula_ref.status === 'new') {
          exec_formula(formula_ref);
          return formula.sheet[str_expression].v;
        }
        else if (formula_ref.status === 'working') {
          throw new Error('Circular ref');
        }
        else if (formula_ref.status === 'done') {
          return formula.sheet[str_expression].v;
        }
      }
      else {
        return formula.sheet[str_expression].v;
      }
    };
  }

  function col_str_2_int(col_str) {
    var r = 0;
    var colstr = col_str.replace(/[0-9]+$/, '');
    for (var i = colstr.length; i--;) {
      r += Math.pow(26, colstr.length - i - 1) * (colstr.charCodeAt(i) - 64);
    }
    return r - 1;
  }

  function int_2_col_str(n) {
    var r = '';
    while (n > 25) {
      n = n - 26;
      r += 'A';
    }
    return r + String.fromCharCode(n + 65);
  }
  
  mymodule.col_str_2_int = col_str_2_int;
  mymodule.int_2_col_str = int_2_col_str;

  function Range(str_expression, formula) {
    this.calc = this.values = function() {
      var arr = str_expression.split(':');
      var min_row = parseInt(arr[0].replace(/^[A-Z]+/, ''), 10);
      var max_row = parseInt(arr[1].replace(/^[A-Z]+/, ''), 10);
      var min_col = col_str_2_int(arr[0]);
      var max_col = col_str_2_int(arr[1]);
      var matrix = [];
      for (var i = min_row; i <= max_row; i++) {
        var row = [];
        matrix.push(row);
        for (var j = min_col; j <= max_col; j++) {
          var cell_name = int_2_col_str(j) + i;
          if (formula.formula_ref[cell_name]) {
            if (formula.formula_ref[cell_name].status === 'new') {
              exec_formula(formula.formula_ref[cell_name]);
            }
            else if (formula.formula_ref[cell_name].status === 'working') {
              throw new Error('Circular ref');
            }
            row.push(formula.sheet[cell_name].v);
          }
          else if (formula.sheet[cell_name]) {
            row.push(formula.sheet[cell_name].v);
          }
          else {
            row.push(null);
          }
        }
      }
      return matrix;
    };
  }

  function Exp(formula) {
    var self = this;
    self.args = [];
    self.name = 'Expression';

    function exec(op, fn) {
      for (var i = 0; i < self.args.length; i++) {
        if (self.args[i] === op) {
          try {
            var r = fn(self.args[i - 1].calc(), self.args[i + 1].calc());
            self.args.splice(i - 1, 3, new RawValue(r));
            i--;
          }
          catch(e) {
            throw Error(formula.name + ': evaluating ' + formula.cell.f);
          }
        }
      }
    }

    function exec_minus() {
      for (var i = self.args.length; i--;) {
        if (self.args[i] === '-') {
          var r = -self.args[i + 1].calc();
          if (typeof self.args[i - 1] !== 'string' && i > 0) {
            self.args.splice(i, 1, '+');
            self.args.splice(i + 1, 1, new RawValue(r));
          }
          else {
            self.args.splice(i, 2, new RawValue(r));
          }
        }
      }
    }

    self.calc = function() {
      exec_minus();
      //console.log('ending of exp...');
      exec('^', function(a, b) {
        //console.log(a, '^', b);
        return Math.pow(+a, +b);
      });
      exec('*', function(a, b) {
        //console.log(a, '*', b);
        return (+a) * (+b);
      });
      exec('/', function(a, b) {
        //console.log(a, '/', b);
        return (+a) / (+b);
      });
      exec('+', function(a, b) {
        //console.log(a, '+', b);
        return (+a) + (+b);
      });
      exec('&', function(a, b) {
        //console.log(a, '&', b);
        return '' + a + b;
      });
      exec('<', function(a, b) {
        return a < b; 
      });
      exec('>', function(a, b) {
        return a > b; 
      });
      exec('>=', function(a, b) {
        return a >= b; 
      });
      exec('<=', function(a, b) {
        return a <= b; 
      });
      exec('<>', function(a, b) {
        return a != b; 
      });
      exec('=', function(a, b) {
        return a == b; 
      });
      if (self.args.length == 1) {
        return self.args[0].calc();
      }
    };

    var last_arg;
    self.push = function(buffer) {
      if (buffer) {
        var v;
        if (!isNaN(buffer)) {
          v = new RawValue(+buffer);
        }
        else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/)) {
          v = new Range(buffer.trim().replace(/\$/g, ''), formula);
        }
        else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[A-Z]+[0-9]+$/)) {
          v = new RefValue(buffer.trim().replace(/\$/g, ''), formula);
        }
        else {
          v = buffer;
        }
        if(((v === '=') && (last_arg == '>' || last_arg == '<')) || (last_arg == '<' && v === '>')) {
          self.args[self.args.length-1] += v;
        } else {
          self.args.push(v);
        }
        last_arg = v;
      }
    };
  }

  var common_operations = {
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

  function exec_formula(formula) {
    formula.status = 'working';
    var root_exp;
    var str_formula = formula.cell.f;
    var exp_obj = root_exp = new Exp(formula);
    var buffer = '', is_string = false, was_string = false;
    var fn_stack = [{
      exp: exp_obj
    }];
    for (var i = 0; i < str_formula.length; i++) {
      if (str_formula[i] == '"') {
        if (is_string) {
          exp_obj.push(new RawValue(buffer));
          is_string = false;
          was_string = true;
        } else {
          is_string = true;
        }
        buffer = '';
      }
      else if (is_string) {
        buffer += str_formula[i];
      }
      else if (str_formula[i] == '(') {
        var o, special = xlsx_Fx[buffer];
        if (special) {
          special = new UserFnExecutor(special, formula);
        }
        else if (buffer) {
          throw new Error(formula.name + ': Function ' + buffer + ' not found');
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
        if(!was_string) {
          exp_obj.push(buffer);
        }
        was_string = false;
        exp_obj.push(str_formula[i]);
        buffer = '';
      }
      else if (str_formula[i] === ',' && fn_stack[fn_stack.length - 1].special) {
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
        if(stack.special) {
          stack.special.push(v);
          exp_obj.push(stack.special);
        } else {
          exp_obj.push(v);
        }
      }
      else {
        buffer += str_formula[i];
      }
    }
    root_exp.push(buffer);
    try {
      formula.cell.v = root_exp.calc();
    }
    catch (e) {
      if (e.message == '#N/A') {
        formula.cell.v = 42;
        formula.cell.t = 'e';
        formula.cell.w = e.message;
      }
      else {
        throw e;
      }
    }
    finally {
      formula.status = 'done';
    }
  }

  function find_all_cells_with_formulas(wb) {
    var formula_ref = {};
    var cells = [];
    for (var sheet_name in wb.Sheets) {
      var sheet = wb.Sheets[sheet_name];
      for (var cell_name in sheet) {
        if (sheet[cell_name].f) {
          var formula = formula_ref[cell_name] = {
            formula_ref: formula_ref,
            wb: wb,
            sheet: sheet,
            cell: sheet[cell_name],
            name: cell_name,
            status: 'new'
          };
          cells.push(formula);
        }
      }
    }
    return cells;
  }

  uexp(this, 'XLSX_CALC', mymodule);

  function uexp(root, MODULENAME, mymodule) {
    /**
     * Generic code to export the module
     */
    var previous_mymodule = root[MODULENAME];
    mymodule.noConflict = function() {
      root[MODULENAME] = previous_mymodule;
      return mymodule;
    };

    // backwards-compatibility for their old module API. If we're in
    // the browser, add the module as a global object.
    if (typeof exports != 'undefined') {
      if (typeof module != 'undefined' && module.exports) {
        exports = module.exports = mymodule;
      }
      exports[MODULENAME] = mymodule;
    }
    else {
      root[MODULENAME] = mymodule;
    }

    // AMD registration happens at the end for compatibility with AMD loaders
    // that may not enforce next-turn semantics on modules. Even though general
    // practice for AMD registration is to be anonymous, underscore registers
    // as a named module because, like jQuery, it is a base library that is
    // popular enough to be bundled in a third party lib, but not be part of
    // an AMD load request. Those cases could generate an error when an
    // anonymous define() is called outside of a loader request.
    if (typeof define == 'function' && define.amd) {
      define(MODULENAME, [], function() {
        return mymodule;
      });
    }
  }

}).call(this);