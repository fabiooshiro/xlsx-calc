"use strict";

module.exports = function getSanitizedSheetName(sheet_name) {
    var quotedMatch = sheet_name.match(/^'(.*)'$/);
    if (quotedMatch) {
        return quotedMatch[1];
    }
    else {
        return sheet_name;
    }
};
