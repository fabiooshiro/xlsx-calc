"use strict";

(function() {
  
  var mymodule = function(workbook) {
    var formulas = find_all_cells_with_formulas(workbook);
    for (var i = formulas.length - 1; i >= 0; i--) {
      expression(formulas[i]);
    }
  };
  
  var instancecount = 0;
  
  function FormulaWrapper(impl, formula) {
    if(!formula) {
      throw Error("Formula is undefined.");
    }
    var id = instancecount++;
    var state = new StateIdle();
    this.args = impl.args || [];
    this.calc = function() {
      return state.calc();
    };
    
    return;
    
    function StateIdle() {
      this.calc = function() {
        console.log('eval', id, formula.name, impl.constructor.name);
        state = new StateBusy();
        var v = impl.calc();
        formula.cell.v = v;
        state = new StateDone(v);
        return state.calc();
      };
    }
    
    function StateDone(value) {
      this.calc = function() {
        return value;
      };
    }
    
    function StateBusy() {
      this.calc = function() {
        throw Error("Circular ref.");
      };
    }
    
  }
  
  var functions = {
    SUM: SumArgs
  };
  
  function RawValue(value) {
    this.calc = function() {
      return value;
    };
  }
  
  function SumArgs() {
    var self = this;
    self.args = [];
    this.calc = function() {
      var r = 0;
      for (var i = self.args.length; i--; ) {
        if(self.args[i].is_range) {
          var range = self.args[i];
          var formulas = range.formulas();
          for(var j = formulas.length; j--; ) {
            r += formulas[j].calc();
          }
        } else {
          r += self.args[i].calc();
        }
      }
      return r;
    };
  }
  
  function SubArgs() {
    var self = this;
    self.args = [];
    this.calc = function() {
      var r = self.args[self.args.length - 1].calc();
      for (var i = self.args.length - 1; i--; ) {
        r -= self.args[i].calc();
      }
      return r;
    };
  }
  
  function Nope() {
    var self = this;
    self.args = [];
    this.calc = function() {
      var v = self.args[0].calc();
      //console.log('Nope result is', v, 'and args length is ', self.args.length);
      return v;
    };
  }
  
  function MulArgs(str_expression, formula) {
    var self = this;
    self.args = [];
    this.calc = function() {
      var r = self.args[self.args.length - 1].calc();
      for (var i = self.args.length - 1; i--; ) {
        r *= self.args[i].calc();
      }
      return r;
    };
  }
  
  function DivArgs(str_expression, formula) {
    var self = this;
    self.args = [];
    this.calc = function() {
      var r = self.args[self.args.length - 1].calc();
      for (var i = self.args.length - 1; i--; ) {
        r /= self.args[i].calc();
      }
      return r;
    };
  }
  
  function RefValue(str_expression, formula) {
    this.calc = function() {
      var ref_cell = formula.sheet[str_expression];
      if(!ref_cell) {
        throw Error("Cell " + str_expression + " not found.");
      }
      var formula_ref = formula.formula_ref[str_expression];
      if(formula_ref && formula_ref.exp_obj) {
        return formula_ref.exp_obj.calc();
      } else {
        if(formula.sheet[str_expression].f) {
          expression(formula.formula_ref[str_expression]);
        }
        return formula.sheet[str_expression].v;
      }
    };
  }
  
  function Range(str_expression, formula) {
    this.is_range = true;
    this.formulas = function() {
      var arr = str_expression.split(':');
      var min_n = parseInt(arr[0].replace(/^[A-Z]+/,''));
      var max_n = parseInt(arr[1].replace(/^[A-Z]+/,''));
      var min_L = arr[0].replace(/[0-9]$/,'');
      var max_L = arr[1].replace(/[0-9]$/,'');
      var r = [];
      var L = min_L;
      for (var i = min_n; i <= max_n; i++) {
        var cell_name = L+i;
        if(!formula.sheet[cell_name].f) {
          r.push(new RawValue(formula.sheet[cell_name].v));
        } else {
          if(!formula.formula_ref[cell_name].exp_obj) {
            expression(formula.formula_ref[cell_name]);
          }
          r.push(formula.formula_ref[cell_name].exp_obj);
        }
      }
      return r;
    };
  }
  function ident(x) {
    var spaces = '';
    for(;x--;) spaces += '  ';
    return spaces;
  }
  function create_expression_obj(str_expression, formula, tabs) {
    tabs = tabs || 0;
    //console.log(ident(tabs), 'create_expression_obj', str_expression);
    var impl_class = functions[str_expression];
    if(impl_class) {
      return new FormulaWrapper(new impl_class(formula), formula);
    } else {
      if(str_expression.match(/^[0-9]+$/)) {
        return new RawValue(parseInt(str_expression, 10));
      } else if(str_expression.match(/^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/)) {
        return new Range(str_expression, formula);
      } else if (str_expression.match(/^[A-Z]+[0-9]+$/)) {
        return new FormulaWrapper(new RefValue(str_expression, formula), formula);
      } else {
        var operation, args;
        if(str_expression.indexOf('+') !== -1) {
          operation = new SumArgs();
          args = str_expression.split('+');
        } else if(str_expression.indexOf('-') !== -1) {
          operation = new SubArgs();
          args = str_expression.split('-');
        } else if(str_expression.indexOf('/') !== -1) {
          operation = new DivArgs();
          args = str_expression.split('/');
        } else if(str_expression.indexOf('*') !== -1) {
          operation = new MulArgs();
          args = str_expression.split('*');
        } else if(str_expression == '') {
          operation = new Nope();
          args = [];
        }
        for (var i = args.length; i--; ) {
          if(args[i]) {
            operation.args.push(create_expression_obj(args[i], formula, tabs + 1));
          }
        }
        return new FormulaWrapper(operation, formula);
      }
    }
  }
  
  function Exp() {
    var self = this;
    self.args = [];
    function exec(op, fn) {
      var found = true;
      while(found) {
        found = false;
        for (var i = 0; i < self.args.length; i++) {
          if(self.args[i] === op) {
            var r = fn(self.args[i-1].calc(), self.args[i+1].calc());
            self.args.splice(i-1, 3, new RawValue(r));
            found = true;
            break;
          }
        }
      }
    }
    function exec_minus() {
      var found = true;
      while(found) {
        found = false;
        for (var i = 0; i < self.args.length; i++) {
          if(self.args[i] === '-' && (self.args[i-1] == undefined || typeof self.args[i-1] == 'string')) {
            var r = -self.args[i+1].calc();
            self.args.splice(i, 2, new RawValue(r));
            found = true;
            break;
          }
        }
      }
    }
    self.calc = function() {
      exec_minus();
      exec('*', function(a, b) { 
        console.log(a, '*', b);
        return a * b;
      });
      exec('/', function(a, b) { return a / b });
      exec('+', function(a, b) { return a + b });
      exec('-', function(a, b) { 
        console.log(a, '-', b);
        return a - b;
      });
      if(self.args.length == 1) {
        return self.args[0].calc();
      }
    };
  }
  
  var common_operations = {
    '*': 'multiply',
    '+': 'plus',
    '-': 'minus',
    '/': 'divide'
  };
  
  function expression(formula) {
    var root_exp;
    var str_formula = formula.cell.f;
    var exp_obj = root_exp = new Exp();
    var buffer = '';
    var fn_stack = [{exp: exp_obj}];
    for (var i = 0; i < str_formula.length; i++) {
      if(str_formula[i] == '(') {
        var special = functions[buffer] ? true : false;
        if(special) {
          var o = create_expression_obj(buffer, formula);
          fn_stack.push({exp: o, special: special});
          exp_obj.args.push(o);
        } else if(common_operations[buffer]) {
          exp_obj.args.push(buffer);
        } else {
          o = new Exp();
          fn_stack.push({exp: o});
          exp_obj.args.push(o);
          exp_obj = o;
        }
        buffer = '';
      } else if(common_operations[str_formula[i]]) {
        console.log('buff', buffer, str_formula[i]);
        if(buffer) {
          if(!isNaN(buffer)) {
            exp_obj.args.push(new RawValue(+buffer));
          } else {
            exp_obj.args.push(buffer);
          }
        }
        exp_obj.args.push(str_formula[i]);
        buffer = '';
      } else if(str_formula[i] == ')') {
        var stack = fn_stack.pop();
        exp_obj = stack.exp;
        console.log('close buff', buffer, 'stack', fn_stack.length);
        if(buffer) {
          exp_obj.args.push(create_expression_obj(buffer, formula));
        }
        buffer = '';
        exp_obj = fn_stack[fn_stack.length-1].exp;
      } else {
        buffer += str_formula[i];
      }
    }
    if(buffer) {
      console.log('buffer', buffer);
      if(!isNaN(buffer)) {
        root_exp.args.push(new RawValue(+buffer));
      } else {
        root_exp.args.push(buffer);
      }
    }
    for (var i = 0; i < root_exp.args.length; i++) {
      console.log('->', root_exp.args[i]);
    }
    formula.cell.v = root_exp.calc();
  } 
  
  function find_all_cells_with_formulas(wb) {
    var formula_ref = {};
    var cells = [];
    for (var sheet_name in wb.Sheets) {
      var sheet = wb.Sheets[sheet_name];
      for (var cell_name in sheet) {
        if(sheet[cell_name].f) {
          var formula = formula_ref[cell_name] = { formula_ref: formula_ref, wb: wb, sheet: sheet, cell: sheet[cell_name], name: cell_name };
          cells.push(formula);
        }
      }
    }
    return cells;
  }

  /**
   * Generic code to export the module
   */
  var MODULENAME = 'XLSX_CALC';
  var root = this;
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

}).call(this);