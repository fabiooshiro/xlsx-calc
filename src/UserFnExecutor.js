"use strict";

const resolve_promises = require('./resolve_promises.js');

module.exports = function UserFnExecutor(user_function) {
    var self = this;
    self.name = 'UserFn';
    self.args = [];
    
    self.calc = function() {
        return new Promise((resolve, reject) => {
            resolve_promises(self.args).then(args => {
                try {
                    resolve(user_function.apply(self, args.map(f=>f.calc())));
                } catch(e) {
                    reject(e);
                }
            }).catch(reject);
        });
    };
    
    self.push = function(buffer) {
        self.args.push(buffer);
    };
    
};