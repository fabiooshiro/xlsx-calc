export const nil = new Error('#NULL!')
export const div0 = new Error('#DIV/0!')
export const value = new Error('#VALUE!')
export const ref = new Error('#REF!')
export const name = new Error('#NAME?')
export const num = new Error('#NUM!')
export const na = new Error('#N/A')
export const error = new Error('#ERROR!')
export const data = new Error('#GETTING_DATA')
export const calc = new Error('#CALC!')

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

export function getErrorValueByMessage(errorMessage: string) {
    return ERROR_MESSAGE_TO_VALUE[errorMessage]
}
