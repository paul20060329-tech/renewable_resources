import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workspace = "C:/Users/ASUS/OneDrive/桌面/renewable_resources";
const sourcePath = `${workspace}/Data/再生资源公司收购合规台账模板.xlsx`;
const outputDir = `${workspace}/outputs/20260909_demo_data`;
const outputPath = `${outputDir}/绿循智审_100条虚拟导入数据.xlsx`;
const dataPath = `${workspace}/Data/绿循智审_100条虚拟导入数据.xlsx`;

const round = (value, digits = 3) => Number(Number(value).toFixed(digits));
const dateOnly = (date) => new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
const timestamp = (date, hour, minute) => new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), hour, minute));
const dateCode = (date) => date.toISOString().slice(0, 10).replaceAll("-", "");
const pad = (value, size = 3) => String(value).padStart(size, "0");
const amountText = (amount) => `人民币${amount.toFixed(2)}元整`;

function makeRows() {
  const goods = [
    ["废钢", "重型废钢", 2650],
    ["废铜", "紫杂铜", 61000],
    ["废铝", "破碎杂铝", 13200],
    ["废电机", "混合电机", 5500],
  ];
  const start = new Date(Date.UTC(2026, 7, 1));
  const rows = [];
  for (let index = 1; index <= 100; index += 1) {
    const date = new Date(start.getTime() + (index - 1) * 86400000);
    const code = dateCode(date);
    const suffix = pad(index);
    const company = index % 10 === 0;
    const supplier = company ? `虚拟再生企业${suffix}` : `模拟自然人${suffix}`;
    const supplierType = company ? "公司" : "个体散户";
    const businessType = company ? "企业采购" : "自然人收购";
    const identity = company ? `SIMORG${pad(index, 6)}` : `SIMID${19800000 + index}`;
    const phone = company ? `SIM400000${pad(index, 3)}` : `SIM138000${pad(index, 5)}`;
    const [goodsName, spec, price] = goods[(index - 1) % goods.length];
    const gross = round(8 + ((index * 17) % 35) / 10, 3);
    const tare = round(2 + ((index * 11) % 14) / 10, 3);
    const net = round(gross - tare, 3);
    const riskQuantity = index === 17;
    const riskTime = index === 23;
    const riskIdentity = index === 29;
    const riskInvoice = index === 31;
    const riskFrequency = index === 37;
    const riskSignature = index === 41;
    const deduct = round(0.05 + (index % 5) * 0.03, 3);
    const inbound = riskQuantity ? 0 : round(net - deduct, 3);
    const settlement = riskQuantity ? round(net - deduct, 3) : inbound;
    const amount = round(settlement * price, 2);
    const weighingNo = `GB${code}${suffix}`;
    const batchNo = `SG${code}${suffix}`;
    const inspectNo = `ZJ${code}${suffix}`;
    const inboundNo = `RK${code}${suffix}`;
    const settlementNo = `JS${code}${suffix}`;
    const invoiceNo = `FX${code}${suffix}`;
    const weighTime = timestamp(date, 9, 0);
    const inspectTime = riskTime ? timestamp(date, 8, 30) : timestamp(date, 9, 15);
    const inboundTime = timestamp(date, 9, 30);
    const paymentTime = timestamp(new Date(date.getTime() + 86400000), 10, 0);
    const annual = riskFrequency ? 5200000 : 320000 + index * 17000;
    const frequency = riskFrequency ? 8 : (index % 5) + 1;
    const risk = riskQuantity || riskTime || riskIdentity || riskInvoice || riskFrequency || riskSignature;
    const paid = !risk && (index % 3 === 0 || index === 1) ? amount : 0;
    const invoiceMode = riskInvoice ? "none" : company || index % 2 === 0 ? "reverse" : "self";
    const invoiceText = invoiceMode === "reverse" ? "反向开票授权" : invoiceMode === "self" ? "自然人自行代开发票" : "未提供合规票据";
    const exception = riskQuantity ? "入库数量缺失" : riskTime ? "质检时间早于过磅时间" : riskIdentity ? "自然人身份资料缺失" : riskInvoice ? "未提供合规票据" : riskFrequency ? "近30日交易频次偏高" : riskSignature ? "财务复核签字缺失" : "";
    rows.push({
      "台账编号": `HG${code}${suffix}`,
      "业务日期": dateOnly(date),
      "业务类型": businessType,
      "收购工单号": batchNo,
      "供货人姓名": supplier,
      "身份证号": riskIdentity ? "" : identity,
      "联系电话": phone,
      "供货人类别": supplierType,
      "过磅单号": weighingNo,
      "过磅时间": weighTime,
      "质检单号": inspectNo,
      "质检时间": inspectTime,
      "入库单号": inboundNo,
      "入库时间": inboundTime,
      "结算单号": settlementNo,
      "发票/收购凭证号": invoiceNo,
      "物资名称": goodsName,
      "品类规格": spec,
      "计量单位": "吨",
      "过磅毛重(吨)": gross,
      "过磅皮重(吨)": tare,
      "过磅净重(吨)": net,
      "质检扣杂率(%)": round(deduct / Math.max(net, 1) * 100, 2),
      "扣杂后结算数量(吨)": settlement,
      "入库实收数量(吨)": inbound,
      "数量差异(吨)": round(inbound - settlement, 3),
      "差异率(%)": settlement ? round((inbound - settlement) / settlement * 100, 2) : 0,
      "收购单价(元/吨)": price,
      "结算金额(元)": amount,
      "税额(元)": 0,
      "实付金额(元)": paid,
      "金额大写": amountText(amount),
      "付款日期": paid ? paymentTime : null,
      "付款方式": "银行转账",
      "付款账户": "虚拟结算账户",
      "收款户名": supplier,
      "收款账号": `621700000000${pad(index, 7)}`,
      "收款开户行": `虚拟银行支行-${pad((index % 6) + 1, 2)}`,
      "银行回单号": paid ? `FH${code}${suffix}` : "",
      "是否本人收款": "是",
      "三单匹配校验": riskQuantity ? "不通过" : "通过",
      "身份核验": riskIdentity ? "不通过" : "通过",
      "收款账户一致性": "通过",
      "累计交易额监控": riskFrequency ? "预警" : "正常",
      "近12月累计交易额(元)": annual,
      "时间逻辑校验": riskTime ? "不通过" : "通过",
      "发票/凭证合规": invoiceText,
      "整体合规结论": risk ? "异常" : paid ? "合规" : "待付款",
      "风险标记": risk ? "关注" : "无",
      "异常说明": exception,
      "制单人": "模拟制单员",
      "财务审核人": riskSignature ? "" : "模拟复核员",
      "财务审核日期": riskSignature ? null : dateOnly(date),
      "审批人(财务负责人)": !risk && amount >= 50000 ? "模拟财务负责人" : "无需审批",
      "归档状态": paid && !risk ? "已归档" : "待归档",
      "归档日期": paid && !risk ? dateOnly(new Date(date.getTime() + 2 * 86400000)) : null,
      "档案位置/编号": paid && !risk ? `ARCHIVE/${code}/${suffix}` : "",
      "附件数量": paid && !risk ? 6 : 3,
      "备注": risk ? "待风险处置，禁止付款" : paid ? "规则全通过且已全额付款" : "规则全通过，等待付款",
    });
  }
  return rows;
}

function scrubExamples(workbook) {
  const replacements = [
    ["张三", "模拟自然人001"], ["王五", "模拟制单员"], ["周八", "模拟复核员"], ["孙七", "模拟财务负责人"],
    ["成都分行", "虚拟银行总行"], ["成都市", "虚拟区域"], ["成都", "虚拟区域"], ["金牛区", "虚拟片区"],
    ["中国工商银行", "虚拟银行"],
  ];
  for (const sheet of workbook.worksheets.items) {
    const used = sheet.getUsedRange();
    if (!used) continue;
    const values = used.values;
    const cleaned = values.map(row => row.map(value => {
      if (typeof value !== "string") return value;
      return replacements.reduce((text, [from, to]) => text.replaceAll(from, to), value);
    }));
    used.values = cleaned;
  }
}

const input = await FileBlob.load(sourcePath);
const workbook = await SpreadsheetFile.importXlsx(input);
const ledger = workbook.worksheets.getItem("合规台账");
const headers = ledger.getRange("A3:BG3").values[0].map(value => String(value ?? "").trim());
ledger.getRange("A4:BG200").clear({ applyTo: "contents" });
const rows = makeRows();
ledger.getRange("A4:BG103").values = rows.map(row => headers.map(header => row[header] ?? null));

// Keep the source style while giving the generated data readable native types.
ledger.getRange("B4:B103").format.numberFormat = "yyyy-mm-dd";
for (const range of ["J4:J103", "L4:L103", "N4:N103"]) ledger.getRange(range).format.numberFormat = "yyyy-mm-dd hh:mm";
for (const range of ["AG4:AG103", "BA4:BA103", "BD4:BD103"]) ledger.getRange(range).format.numberFormat = "yyyy-mm-dd";
for (const range of ["T4:Z103", "AA4:AA103", "AB4:AE103", "AS4:AS103"]) ledger.getRange(range).format.numberFormat = "#,##0.000";
ledger.getRange("AC4:AE103").format.numberFormat = "#,##0.00";
ledger.getRange("BF4:BF103").format.numberFormat = "0";
ledger.getRange("F4:G103").format.numberFormat = "@";
ledger.getRange("AK4:AK103").format.numberFormat = "@";
ledger.freezePanes.freezeRows(3);
scrubExamples(workbook);

workbook.recalculate();
const inspection = await workbook.inspect({ kind: "table", sheetId: "合规台账", range: "A1:BG8", include: "values,formulas", tableMaxRows: 8, tableMaxCols: 59, tableMaxCellChars: 100, maxChars: 18000 });
console.log(inspection.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "generated data formula error scan" });
console.log(errors.ndjson);
const preview = await workbook.render({ sheetName: "合规台账", range: "A1:BG18", scale: 1, format: "png" });
await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(`${outputDir}/preview.png`, new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
await fs.copyFile(outputPath, dataPath);
console.log(`WROTE ${outputPath}`);
console.log(`COPIED ${dataPath}`);
