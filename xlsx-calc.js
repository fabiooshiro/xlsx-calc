"use strict";

(function() {

  var mymodule = function(workbook) {
    var formulas = find_all_cells_with_formulas(workbook);
    for (var i = formulas.length - 1; i >= 0; i--) {
      expression(formulas[i]);
    }
  };

  var functions = {
    SUM: SumArgs
  };
  
  function PushDecorator(impl, formula) {
    var self = this;
    self.calc = impl.calc;
    self.args = impl.args;
    self.push = function(buffer) {
      if (buffer) {
        if (!isNaN(buffer)) {
          impl.args.push(+buffer);
        }
        else if (buffer.match(/^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/)) {
          var range = new Range(buffer, formula).values();
          for (var i = 0; i < range.length; i++) {
            self.args.push(range[i]);
          }
        }
        else if(buffer.match(/^[A-Z]+[0-9]+$/)) {
          impl.args.push(new RefValue(buffer, formula).calc());
        }
        else {
          impl.args.push(buffer);
        }
      }
    };
  }

  function RawValue(value) {
    this.calc = function() {
      return value;
    };
  }

  function SumArgs(formula) {
    var self = this;
    self.args = [];
    this.calc = function() {
      var r = 0;
      for (var i = self.args.length; i--;) {
        r += self.args[i];
      }
      return r;
    };
  }

  function RefValue(str_expression, formula) {
    this.calc = function() {
      var ref_cell = formula.sheet[str_expression];
      if (!ref_cell) {
        throw Error("Cell " + str_expression + " not found.");
      }
      var formula_ref = formula.formula_ref[str_expression];
      if (formula_ref && formula_ref.exp_obj) {
        return formula_ref.exp_obj.calc();
      }
      else {
        if (formula.sheet[str_expression].f) {
          expression(formula.formula_ref[str_expression]);
        }
        return formula.sheet[str_expression].v;
      }
    };
  }

  function Range(str_expression, formula) {
    this.is_range = true;
    this.values = function() {
      var arr = str_expression.split(':');
      var min_n = parseInt(arr[0].replace(/^[A-Z]+/, ''));
      var max_n = parseInt(arr[1].replace(/^[A-Z]+/, ''));
      var min_L = arr[0].replace(/[0-9]$/, '');
      var max_L = arr[1].replace(/[0-9]$/, '');
      var r = [];
      var L = min_L;
      for (var i = min_n; i <= max_n; i++) {
        var cell_name = L + i;
        r.push(formula.sheet[cell_name].v);
      }
      return r;
    };
  }

  function Exp(formula) {
    var self = this;
    self.args = [];

    function exec(op, fn) {
      for (var i = 0; i < self.args.length; i++) {
        if (self.args[i] === op) {
          var r = fn(self.args[i - 1].calc(), self.args[i + 1].calc());
          self.args.splice(i - 1, 3, new RawValue(r));
          i--;
        }
      }
    }

    function exec_minus() {
      for (var i = self.args.length - 1; i--;) {
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
      exec('^', function(a, b) {
        console.log(a, '^', b);
        return Math.pow(a, b);
      });
      exec('*', function(a, b) {
        console.log(a, '*', b);
        return a * b;
      });
      exec('/', function(a, b) {
        console.log(a, '/', b);
        return a / b;
      });
      exec('+', function(a, b) {
        console.log(a, '+', b);
        return a + b;
      });
      if (self.args.length == 1) {
        return self.args[0].calc();
      }
    };
    
    self.push = function(buffer) {
      if (buffer) {
        if (!isNaN(buffer)) {
          self.args.push(new RawValue(+buffer));
        }
        else if(buffer.match(/^[A-Z]+[0-9]+$/)) {
          self.args.push(new RefValue(buffer, formula));
        }
        else {
          self.args.push(buffer);
        }
      }
    };
  }

  var common_operations = {
    '*': 'multiply',
    '+': 'plus',
    '-': 'minus',
    '/': 'divide',
    '^': 'power'
  };

  function expression(formula) {
    var root_exp;
    var str_formula = formula.cell.f;
    var exp_obj = root_exp = new Exp(formula);
    var buffer = '';
    var fn_stack = [{
      exp: exp_obj
    }];
    for (var i = 0; i < str_formula.length; i++) {
      if (str_formula[i] == '(') {
        var special = functions[buffer];
        if (special) {
          var o = new PushDecorator(new special(), formula);
          fn_stack.push({
            exp: o,
            special: true
          });
          exp_obj.args.push(o);
        }
        else if (common_operations[buffer]) {
          exp_obj.args.push(buffer);
        }
        else {
          o = new Exp(formula);
          fn_stack.push({
            exp: o
          });
          exp_obj.args.push(o);
          exp_obj = o;
        }
        buffer = '';
      }
      else if (common_operations[str_formula[i]]) {
        exp_obj.push(buffer);
        exp_obj.args.push(str_formula[i]);
        buffer = '';
      }
      else if (str_formula[i] == ')') {
        var stack = fn_stack.pop();
        exp_obj = stack.exp;
        exp_obj.push(buffer);
        buffer = '';
        exp_obj = fn_stack[fn_stack.length - 1].exp;
      }
      else {
        buffer += str_formula[i];
      }
    }
    root_exp.push(buffer);
    // for (var i = 0; i < root_exp.args.length; i++) {
    //   console.log('->', root_exp.args[i]);
    // }
    formula.cell.v = root_exp.calc();
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
            name: cell_name
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