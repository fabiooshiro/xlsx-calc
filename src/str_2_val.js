const RawValue = require('./RawValue.js');
const RefValue = require('./RefValue.js');
const Range = require('./Range.js');

module.exports = function str_2_val(buffer, formula) {
    var v;
    if (!isNaN(buffer)) {
        v = new RawValue(+buffer);
    }
    else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/)) {
        v = new Range(buffer.trim().replace(/\$/g, ''), formula);
    }
    else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[^!]+![A-Z]+[0-9]+:[A-Z]+[0-9]+$/)) {
        v = new Range(buffer.trim().replace(/\$/g, ''), formula);
    }
    else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[A-Z]+:[A-Z]+$/)) {
        v = new Range(buffer.trim().replace(/\$/g, ''), formula);
    }
    else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[^!]+![A-Z]+:[A-Z]+$/)) {
        v = new Range(buffer.trim().replace(/\$/g, ''), formula);
    }
    else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[A-Z]+[0-9]+$/)) {
        v = new RefValue(buffer.trim().replace(/\$/g, ''), formula);
    }
    else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[^!]+![A-Z]+[0-9]+$/)) {
        v = new RefValue(buffer.trim().replace(/\$/g, ''), formula);
    }
    else if (typeof buffer === 'string' && !isNaN(buffer.trim().replace(/%$/, ''))) {
        v = new RawValue(+(buffer.trim().replace(/%$/, '')) / 100.0);
    }
    else {
        v = buffer;
    }
    return v;
};