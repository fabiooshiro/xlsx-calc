export function LazyValue(fn) {
    this.calc = function() {
        return fn();
    };
};
