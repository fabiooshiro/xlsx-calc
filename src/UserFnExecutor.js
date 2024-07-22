"use strict";

const { getErrorValueByMessage } = require('./errors')

module.exports = function UserFnExecutor(user_function) {
    var self = this;
    self.name = 'UserFn';
    self.args = [];
    self.calc = function () {
        var result;
        try {
            result = user_function.apply(self, self.args.map(f => f.calc()));
            if (result instanceof Error) {
                throw result;
            }
        } catch (e) {
            const errorValue = getErrorValueByMessage(e.message)
            if (user_function.name === 'is_blank'
                && errorValue !== undefined) {
                // is_blank applied to an error cell doesn't propagate the error
                result = 0;
            }
            else if (user_function.name === 'iserror'
                && errorValue !== undefined) {
                // iserror applied to an error doesn't propagate the error and returns true
                result = true;
            }
            else {
                throw e;
            }
        }
        return result;
    };
    self.push = function (buffer) {
        self.args.push(buffer);
    };
};