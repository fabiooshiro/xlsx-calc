import { RawValue } from './RawValue';
import { RefValue } from './RefValue';
import { LazyValue } from './LazyValue';
import { Range } from './Range';

// this is used to _cache_ range names so that it doesn't need to be queried
// every time a range is used
let definedNames, wb;
function getDefinedName(buffer, formula) {
    if (!(formula.wb.Workbook && formula.wb.Workbook.Names)) {
        return null;
    }
    if (wb !== formula.wb) {
        wb = formula.wb;
        definedNames = null;
        return getDefinedName(buffer, formula);
    }
    if (definedNames) {
        return definedNames[buffer];
    }
    const keys: any[] = Object.values(formula.wb.Workbook.Names);
    if (keys.length === 0) {
        return;
    }
    definedNames = {};
    keys.forEach(({ Name, Ref }) => {
        if (!Name.includes('.')) {
            definedNames[Name] = Ref;
        }
    });

    return getDefinedName(buffer, formula);
}

export function str_2_val(buffer, formula) {
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
    if (getDefinedName(buffer, formula)) {
        return str_2_val(getDefinedName(buffer, formula), formula);
    }
    return buffer;
};
