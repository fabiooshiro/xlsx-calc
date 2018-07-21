"use strict";

const RawValue = require('./RawValue.js');
const RefValue = require('./RefValue.js');
const Range = require('./Range.js');
const resolve_promises = require('./resolve_promises.js');
let exp_id = 0;
let exec_id = 0;

module.exports = function Exp(formula) {
    var self = this;
    self.id = ++exp_id;
    self.args = [];
    self.name = 'Expression';
    self.update_cell_value = update_cell_value;
    self.formula = formula;
    
    function handleException(e, formula, resolve, reject) {
        if (e.message == '#N/A') {
            formula.cell.v = 42;
            formula.cell.t = 'e';
            formula.cell.w = e.message;
            resolve();
        }
        else {
            //console.error('Error', current_execution, e);
            reject(e);
            //throw e;
        }
    }
    
    function update_cell_value() {
        let hasPromise = false;
        return new Promise((resolve, reject) => {
            //let current_execution = exec_id++;
            try {
                //console.log('Exec', current_execution, formula.name, formula.cell.f);
                var val_or_promise = self.calc();
                if (typeof val_or_promise === 'object' && typeof val_or_promise['then'] === 'function') {
                    hasPromise = true;
                    val_or_promise.then(res => {
                        formula.cell.v = res;
                        if (typeof(formula.cell.v) === 'string') {
                            formula.cell.t = 's';
                        }
                        else if (typeof(formula.cell.v) === 'number') {
                            formula.cell.t = 'n';
                        }
                        formula.status = 'done';
                        resolve(formula.cell.v);
                    }).catch(e => {
                        //console.log('Exp', self.id, 'error:', e);
                        //reject(e);
                        formula.status = 'done';
                        handleException(e, formula, resolve, reject);
                    });
                }
                else {
                    formula.cell.v = val_or_promise;
                    if (typeof(formula.cell.v) === 'string') {
                        formula.cell.t = 's';
                    }
                    else if (typeof(formula.cell.v) === 'number') {
                        formula.cell.t = 'n';
                    }
                    resolve(formula.cell.v);
                }
            }
            catch (e) {
                handleException(e, formula, resolve, reject);
            }
            finally {
                if (!hasPromise) {
                    formula.status = 'done';
                }
            }
        });
    }
    
    function exec(op, args, fn) {
        for (var i = 0; i < args.length; i++) {
            if (args[i] === op) {
                try {
                    var r = fn(args[i - 1].calc(), args[i + 1].calc());
                    args.splice(i - 1, 3, new RawValue(r));
                    i--;
                }
                catch (e) {
                    console.error(e);
                    throw Error(formula.name + ': evaluating ' + formula.cell.f + '\n' + e.message);
                    //throw e;
                }
            }
        }
    }

    function exec_minus(args) {
        for (var i = args.length; i--;) {
            if (args[i] === '-') {
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
        return new Promise((resolve, reject) => {
            resolve_promises(self.args.concat()).then(args => {
                try {
                    exec_minus(args);
                    exec('^', args, function(a, b) {
                        return Math.pow(+a, +b);
                    });
                    exec('*', args, function(a, b) {
                        return (+a) * (+b);
                    });
                    exec('/', args, function(a, b) {
                        return (+a) / (+b);
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
                        if (typeof args[0] === 'object' && typeof args[0]['then'] === 'function') {
                            args[0].then(resolve).catch(reject);
                            return;
                        }
                        if (typeof(args[0].calc) != 'function') {
                            return resolve(args[0]);
                        }
                        else {
                            return resolve(args[0].calc());
                        }
                    }
                    else {
                        console.log('something is not right');
                    }
                } catch(e) {
                    reject(e);
                }
            }).catch(reject);
        });
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