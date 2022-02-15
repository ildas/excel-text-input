import { ExcelHelper } from "./helpers";

let excelHelper = new ExcelHelper();

/* global document, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = excelHelper.run.bind(excelHelper);
    document.querySelector("textarea").onchange = excelHelper.handleActiveCellChange;
  }
});
