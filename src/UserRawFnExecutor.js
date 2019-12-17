"use strict";

module.exports = function UserRawFnExecutor(user_function, formula) {
    var self = this;
    self.name = 'UserRawFn';
    self.args = [];
    self.calc = function() {
        try {
            return user_function.apply(self, self.args);
        } catch(e) {
            // debug
            // console.log('----------------', user_function);
            // console.log(formula.name);
            // console.log(self);
            throw e;
        }
    };
    self.push = function(buffer) {
        self.args.push(buffer);
    };
};
