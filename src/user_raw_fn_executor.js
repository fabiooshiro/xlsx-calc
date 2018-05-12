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
