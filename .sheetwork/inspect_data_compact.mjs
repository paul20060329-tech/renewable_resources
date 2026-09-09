import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const input = await FileBlob.load("../Data/绿循智审_100条虚拟导入数据.xlsx");
const wb = await SpreadsheetFile.importXlsx(input);
for (const sheet of wb.worksheets.items) {
  const used = sheet.getUsedRange();
  console.log(JSON.stringify({name:sheet.name,address:used?.address,rows:used?.rowCount,cols:used?.columnCount,values:used?.values?.slice(0,3)}));
}
