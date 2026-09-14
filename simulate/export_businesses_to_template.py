"""把 greenloop.db 中现有的业务数据导出为标准导入模板格式（过磅单/质检单/入库单）。

产出的 Excel 可以直接在「单据导入」页上传，或在「同步本地模板」时被识别，
用于把库里的历史数据迁移/备份出来再重新导入。

用法：
    python simulate/export_businesses_to_template.py [输出路径]
"""
import json
import os
import sqlite3
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(ROOT, "greenloop.db")
DEFAULT_OUTPUT = os.path.join(ROOT, "Data", "绿循智审_现有业务数据_可导入备份.xlsx")

WEIGH_HEADERS = ["过磅单编号", "货物批次号", "过磅时间", "供货方名称", "供货方身份证号", "供货方类别", "个体户登记标记", "车牌号", "货物品名", "毛重 (Kg)", "皮重 (Kg)", "净重 (Kg)", "过磅方式", "过磅员签字", "司机签字", "收款户名", "银行卡号", "开户行"]
INSPECT_HEADERS = ["质检单编号", "关联过磅单编号", "质检时间", "货物品名", "货品规格", "质检等级", "质检扣杂数量 (Kg)", "含水率 / 杂质率", "计价单价 (元 / Kg)", "计价金额 (元)", "质检员签字", "复核人签字", "票据合规方式", "结算数量 (Kg)"]
INBOUND_HEADERS = ["入库单编号", "关联过磅单编号", "入库时间", "入库数量 (Kg)", "仓管员签字", "付款时间"]

INVOICE_TEXT = {"reverse": "反向开票", "self": "自然人自行代开发票", "none": "无票据"}


def load_rows():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    rows = [json.loads(item["payload"]) for item in con.execute("SELECT payload FROM businesses").fetchall()]
    con.close()
    rows.sort(key=lambda row: str(row.get("id", "")))
    return rows


def build_workbook(rows):
    wb = Workbook()
    weigh = wb.active; weigh.title = "过磅单"
    inspect = wb.create_sheet("质检单")
    inbound = wb.create_sheet("入库单")
    weigh.append(WEIGH_HEADERS); inspect.append(INSPECT_HEADERS); inbound.append(INBOUND_HEADERS)
    for row in rows:
        business_id = row.get("id", "")
        weigh.append([
            business_id, row.get("batch", ""), row.get("weigh_time", ""), row.get("supplier", ""),
            row.get("id_no", ""), row.get("supplier_type", "自然人"), row.get("registered", "否"),
            row.get("plate", ""), row.get("goods", ""), row.get("gross", 0), row.get("tare", 0),
            row.get("net", 0), row.get("method", ""), row.get("weigher", ""), row.get("driver", ""),
            row.get("bank_name", ""), row.get("bank_account", ""), row.get("bank_branch", ""),
        ])
        inspect.append([
            row.get("inspect_id", ""), business_id, row.get("inspect_time", ""),
            row.get("inspect_goods") or row.get("goods", ""), row.get("spec", ""), row.get("grade", ""),
            row.get("inspect_qty", 0), row.get("moisture", 0), row.get("price", 0), row.get("amount", 0),
            row.get("inspector", ""), row.get("reviewer", ""), INVOICE_TEXT.get(row.get("invoice", ""), "无票据"),
            row.get("settlement_qty", row.get("inbound_qty", 0)),
        ])
        inbound.append([
            row.get("inbound_id", ""), business_id, row.get("inbound_time", ""),
            row.get("inbound_qty", 0), "模板导入", row.get("payment_time", ""),
        ])
    for ws in (weigh, inspect, inbound):
        ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="167A59")
        for column in ws.columns:
            ws.column_dimensions[column[0].column_letter].width = 20
    return wb


def main():
    output = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_OUTPUT
    rows = load_rows()
    os.makedirs(os.path.dirname(output), exist_ok=True)
    build_workbook(rows).save(output)
    print(f"已导出 {len(rows)} 笔业务 -> {output}")


if __name__ == "__main__":
    main()
