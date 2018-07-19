"use strict";

const RawValue = require('./RawValue.js');

module.exports = function resolve_promises(args) {
    //console.log('resolvendo promises...');
    return new Promise((resolve, reject) => {
        let promises = [],
            new_args = [];
        for (let i = 0; i < args.length; i++) {
            let arg = args[i];
            if (typeof arg === 'object' && typeof arg['calc'] === 'function') {
                let val_or_promise = arg.calc();
                if (typeof val_or_promise === 'object' && typeof val_or_promise['then'] === 'function') {
                    promises.push(val_or_promise);
                    val_or_promise.then(r => {
                        new_args[i] = new RawValue(r);
                    }).catch(reject);
                }
                else {
                    new_args[i] = new RawValue(val_or_promise);
                }
            }
            else {
                new_args[i] = arg;
            }
        }
        Promise.all(promises).then(() => {
            //console.log('new_args =', new_args.map(a => a.calc()));
            resolve(new_args);
        }).catch(reject);
    });
};