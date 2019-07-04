"use strict";

module.exports = function LazyValue(fn) {
    this.calc = function() {
        return fn();
    };
};
