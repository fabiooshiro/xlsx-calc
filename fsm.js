function FSM(formula) {
    var str_formula = formula.cell.f;
    var state = new Start();
    var root_exp = state.exp;
    var stack = [state];
    for (var i = 0; i < str_formula.length; i++) {
        state.read(str_formula[i]);
    }
    state.end();
    console.log('root exp', root_exp.args);
    return;

    function Num() {
        var self = this;
        var buffer = '';
        self.read = function(char) {
            if (char == ')') {
                stack.pop();
                state = stack.pop();
            }
            else if (char == '+') {
                state = new Add();
            }
            else {
                buffer += char;
            }
        };
        self.calc = function() {

        }
        self.end = function() {

        }
    }

    function Start() {
        var self = this;
        self.exp = new BasicExp();
        var buffer = '';
        self.read = function(char) {
            if (char == '(') {
                state = new Start();
                self.exp.args.push(state.exp);
                stack.push(state);
            }
            else if ('0123456789'.indexOf(char) !== -1) {
                state = new Num();
                self.exp.args.push(state);
                stack.push(state);
            }
            else {
                buffer += char;
            }
        };
    }

    function Add() {
        var self = this;
        self.args = [];
        var buffer = '';
        self.read = function(char) {
            buffer += char;
        };
        self.end = function() {
            console.log('bf', buffer);
            var r = +buffer;
            for (var i = self.args.length - 1; i >= 0; i--) {
                r += self.args[i];
            }
            console.log('r', r);
        };
    }

    function BasicExp() {
        this.args = [];
    }
}