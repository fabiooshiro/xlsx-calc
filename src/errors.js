const nil = new Error('#NULL!')
const div0 = new Error('#DIV/0!')
const value = new Error('#VALUE!')
const ref = new Error('#REF!')
const name = new Error('#NAME?')
const num = new Error('#NUM!')
const na = new Error('#N/A')
const error = new Error('#ERROR!')
const data = new Error('#GETTING_DATA')
const calc = new Error('#CALC!')

const ERROR_MESSAGE_TO_VALUE = {
    '#NULL!': 0x00,
    '#DIV/0!': 0x07,
    '#VALUE!': 0x0F,
    '#REF!': 0x17,
    '#NAME?': 0x1D,
    '#NUM!': 0x24,
    '#N/A': 0x2A,
    '#GETTING_DATA': 0x2B,
    '#CALC!': 0x00, // todo: set the correct error code
};

function getErrorValueByMessage(errorMessage) {
    return ERROR_MESSAGE_TO_VALUE[errorMessage]
}

module.exports = {
    nil,
    div0,
    value,
    ref,
    name,
    num,
    na,
    error,
    data,
    calc,
    getErrorValueByMessage,
}