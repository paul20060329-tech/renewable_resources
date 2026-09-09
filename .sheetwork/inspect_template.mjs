import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const paths = [
  "../Data/再生资源公司收购合规台账模板.xlsx",
  "../Data/绿循智审_100条虚拟导入数据.xlsx",
];

for (const path of paths) {
  const input = await FileBlob.load(path);
  const wb = await SpreadsheetFile.importXlsx(input);
  console.log(`FILE ${path}`);
  console.log((await wb.inspect({ kind: "workbook,sheet,table", maxChars: 12000, tableMaxRows: 5, tableMaxCols: 12, tableMaxCellChars: 120 })).ndjson);
  for (const sheet of wb.worksheets.items) {
    console.log(`SHEET ${sheet.name}`);
    const used = sheet.getUsedRange();
    if (used) {
      console.log((await wb.inspect({ kind: "table", sheetId: sheet.name, range: used.address, include: "values,formulas", tableMaxRows: 8, tableMaxCols: 70, tableMaxCellChars: 120, maxChars: 24000 })).ndjson);
      const preview = await wb.render({ sheetName: sheet.name, autoCrop: "all", scale: 1, format: "png" });
      const safe = path.includes("100条") ? "data" : "template";
      const name = sheet.name.replace(/[\\/:*?"<>|]/g, "_");
      await fs.writeFile(`./${safe}_${name}.png`, new Uint8Array(await preview.arrayBuffer()));
    }
  }
}
