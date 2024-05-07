const error = require('./errors');

module.exports = function dynamic_array_compatible(fn) {
    return function () {
        var hasArray = false;
        var parameters = [];
        var numberOfRows = [];
        var numberOfColumns = [];
        for (var i = 0, length = arguments.length; i < length; i++) {
            var argument = arguments[i];
            var isArray = Array.isArray(argument);
            if (isArray) {
                hasArray = true;
            }
            var parameter = isArray ? argument.map(value => Array.isArray(value) ? value : [value]) : [[argument]];
            parameters[i] = parameter;
            numberOfRows[i] = parameter.length;
            numberOfColumns[i] = parameter[0].length;
        }
        var maxNumberOfRows = Math.max.apply(null, numberOfRows);
        var maxNumberOfColumns = Math.max.apply(null, numberOfColumns);
        var results = [];
        for (var x = 0; x < maxNumberOfRows; x++) {
            results[x] = [];
            for (var y = 0; y < maxNumberOfColumns; y++) {
                var hasError = false;
                var elements = [];
                for (var i = 0, length = parameters.length; i < length; i++) {
                    var element = parameters[i];
                    var xKey = numberOfRows[i] === 1 ? 0 : x;
                    var yKey = numberOfColumns[i] === 1 ? 0 : y;
                    if (xKey in element) {
                        element = element[xKey];
                        if (yKey in element) {
                            element = element[yKey];
                        } else {
                            hasError = true;
                            break;
                        }
                    } else {
                        hasError = true;
                        break;
                    }
                    elements[i] = element;
                }
                results[x][y] = hasError ? error.na : fn.apply(null, elements);
            }
        }
        return hasArray ? results : results[0][0];
    }
};
