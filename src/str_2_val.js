const RawValue = require('./RawValue.js');
const RefValue = require('./RefValue.js');
const LazyValue = require('./LazyValue.js');
const Range = require('./Range.js');

module.exports = function str_2_val(buffer, formula) {
    if (!isNaN(buffer)) {
        return new RawValue(+buffer);
    }
    if (buffer === 'TRUE') {
        return new RawValue(1);
    }
    if (typeof buffer !== 'string') {
        return buffer;
    }

    buffer = buffer.trim().replace(/\$/g, '')

    if (buffer.match(/^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/)) {
        return new Range(buffer, formula);
    }
    if (buffer.match(/^[^!]+![A-Z]+[0-9]+:[A-Z]+[0-9]+$/)) {
        return new Range(buffer, formula);
    }
    if (buffer.match(/^[A-Z]+:[A-Z]+$/)) {
        return new Range(buffer, formula);
    }
    if (buffer.match(/^[^!]+![A-Z]+:[A-Z]+$/)) {
        return new Range(buffer, formula);
    }
    if (buffer.match(/^[A-Z]+[0-9]+$/)) {
        return new RefValue(buffer, formula);
    }
    if (buffer.match(/^[^!]+![A-Z]+[0-9]+$/)) {
        return new RefValue(buffer, formula);
    }
    if (buffer.match(/%$/)) {
        var inner = str_2_val(buffer.substr(0, buffer.length-1), formula)
        return new LazyValue(() => inner.calc() / 100)
    }
    return buffer;
};
