import { ExcelHelper } from "../utils/helpers";

let excelHelper = new ExcelHelper();

/* global document, Office */

//TODO add events orchestrator function
implementTaskpaneEvents();

async function isThisExcel(): Promise<boolean> {
  const info = await Office.onReady();
  return info.host === Office.HostType.Excel
}

//TODO rework not to implement a boolean check, but a class possibly
async function implementTaskpaneEvents(): Promise<void> {
  console.log(await isThisExcel())
  if (await isThisExcel()){
    document.querySelector("textarea").oninput = excelHelper.changeCellValue.bind(excelHelper, "textarea");
    document.querySelector("textarea").onchange = excelHelper.clearNonCellElementValue;
  }
}



