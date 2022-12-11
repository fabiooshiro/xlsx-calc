import * as assert from 'assert';
import XLSX_CALC from "../src";

xdescribe("verifica condicionais", () => {
  it("valida ifs", () => {
    let workbook: any = {
      Sheets: {
        Sheet1: { A1: {} }
      }
    };
    workbook.Sheets.Sheet1.A1.f = "IF(AND(1<2,1<3),123,3)";
    XLSX_CALC(workbook);
    assert.equal(workbook.Sheets.Sheet1.A1.v, 123);
  });
  it("valida ifs", () => {
    let workbook: any = {
      Sheets: {
        Sheet1: { A1: {} }
      }
    };
    XLSX_CALC.set_fx("IS_IOS", () => {
      return false;
    });
    workbook.Sheets.Sheet1.A1.f = "IF(AND(1<2,IS_IOS()),123,3)";
    XLSX_CALC(workbook);
    assert.equal(workbook.Sheets.Sheet1.A1.v, 123);
  });
});
