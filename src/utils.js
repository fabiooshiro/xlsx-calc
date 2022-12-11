const error = require('./errors')

function parseBool(bool) {
    if (typeof bool === 'boolean') {
        return bool
    }

    if (bool instanceof Error) {
        return bool
    }

    if (typeof bool === 'number') {
        return bool !== 0
    }

    if (typeof bool === 'string') {
        const up = bool.toUpperCase()

        if (up === 'TRUE') {
            return true
        }

        if (up === 'FALSE') {
            return false
        }
    }

    if (bool instanceof Date && !isNaN(bool)) {
        return true
    }

    return error.value
}

// E.g. addEmptyValuesToArray([[1]], 2, 2) => [[1, ""], ["", ""]]
function addEmptyValuesToArray(array, requiredLength, requiredHeight) {
    if (!array || !requiredLength || !requiredHeight) {
        return array
    }

    if (requiredLength < 0 || requiredHeight < 0) {
        return array
    }

    // array must be a square matrix
    if (!Array.isArray(array) || !array.length) return array;
    for (let i = 0; i < array.length; i++) {
        if (!(array[i] instanceof Array)) return array
    }

    // add empty values to columns
    for (let i = 0; i < array.length; i++) {
        if (array[i].length < requiredLength) {
            for (let j = array[i].length; j < requiredLength; j++) {
                array[i].push('')
            }
        }
    }

    // add empty values to rows
    if (array.length < requiredHeight) {
        for (let i = array.length; i < requiredHeight; i++) {
            array.push([])
            for (let j = 0; j < requiredLength; j++) {
                array[i].push('')
            }
        }
    }

    return array
}

module.exports = {
    addEmptyValuesToArray,
    parseBool,
}