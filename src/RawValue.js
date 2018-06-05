"use strict";

module.exports = function RawValue(value) {
    this.setValue = function(v) {
        value = v;
    };
    this.calc = function() {
        return value;
    };
};
