"use strict";

module.exports = function UserFnExecutor(user_function) {
    var self = this;
    self.name = 'UserFn';
    self.args = [];
    self.calc = function() {
        return user_function.apply(self, self.args);
    };
    self.push = function(buffer) {
        self.args.push(buffer.calc());
    };
};