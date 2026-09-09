import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const input = await FileBlob.load("../Data/excel字段与相关细节要求/过磅单和质检单设计.xlsx");
const wb = await SpreadsheetFile.importXlsx(input);
for (const sheet of wb.worksheets.items) {
  const used = sheet.getUsedRange();
  console.log(JSON.stringify({name:sheet.name,address:used?.address,rows:used?.rowCount,cols:used?.columnCount,values:used?.values?.slice(0,8)}));
}
