import { ExcelHelper } from "./helpers";

let excelHelper = new ExcelHelper();

/* global console, document, Excel, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = run;
    document.querySelector("textarea").onchange = excelHelper.handleActiveCellChange;
  }
});

export async function run(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      await excelHelper.syncTextboxAndActiveCell(context);
    });
  } catch (error) {
    console.error(error);
  }
}
