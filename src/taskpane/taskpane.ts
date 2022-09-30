import { ExcelHelper } from "./helpers";

let excelHelper = new ExcelHelper();

/* global document, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.querySelector("textarea").oninput = excelHelper.changeCellText.bind(excelHelper, "textarea");
    document.querySelector("textarea").onchange = excelHelper.clearNonCellElementValue;
  }
});
