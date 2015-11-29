var XLSX_CALC = require("xlsx-calc");
var workbook = { 
    "Sheets": { 
        "Sheet1": {
            "A1": {},
            "A2": {}
        }
    }
};
workbook.Sheets['Sheet1'].A1.f = "2+2";
workbook.Sheets['Sheet1'].A2.f = "3+A1";
XLSX_CALC(workbook);
workbook.Sheets['Sheet1'].A2.v;