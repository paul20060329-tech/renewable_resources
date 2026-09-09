"""Generate 100 rows of fully fictitious renewable-resource demo data using the
"再生资源公司收购合规台账模板" 59-column ledger template.

Run:
    python simulate/regenerate_demo_data.py

Outputs:
    Data/绿循智审_100条虚拟导入数据.xlsx
        ← 100 rows fitted into the 59-column template format
        ← 90% natural persons, 10% companies
        ← No real names, no real place names
        ← Time chain: 过磅 < 质检 < 入库 < 付款
        ← Quantity chain: 过磅净重 - 质检扣杂数量 = 入库数量

This module is independent of the Flask runtime, so it may be run on a fresh
checkout without launching the web server first.
"""

from __future__ import annotations

import os
from datetime import datetime, timedelta

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(ROOT, "Data", "再生资源公司收购合规台账模板.xlsx")
OUTPUT_PATH = os.path.join(ROOT, "Data", "绿循智审_100条虚拟导入数据.xlsx")
ROW_COUNT = 100
PAYMENT_THRESHOLD = 50000

VIRTUAL_SURNAMES = ["赵", "钱", "孙", "李", "周", "吴", "郑", "王", "冯", "陈", "褚", "卫", "蒋", "沈", "韩"]
VIRTUAL_GIVEN = ["建国", "志远", "宏达", "明轩", "国良", "永福", "立军", "志伟", "秀英", "丽华", "文军", "秀梅", "建华", "彩霞", "海涛"]
COMPANY_PREFIX = ["蓝海", "绿野", "银拓", "钢城", "循环", "环保", "资源", "循环再生", "金属再生", "物资回收"]
COMPANY_SUFFIX = ["资源回收有限公司", "再生金属有限公司", "环保科技有限公司", "废旧物资回收部", "循环利用公司"]
BANK_BRANCHES = ["中国工商银行虚拟支行", "中国农业银行虚拟分行", "中国建设银行虚拟支行", "中国银行虚拟分行", "虚拟商业银行支行"]

GOODS_TABLE = [
    ("废钢", "重型废钢", "A级" if True else "B级", 2.65),
    ("废铜", "紫杂铜", "A级", 61.0),
    ("废铝", "破碎杂铝", "B级", 13.2),
    ("废电机", "混合电机", "B级", 5.5),
]


def pick_person_name(index: int) -> str:
    surname = VIRTUAL_SURNAMES[index % len(VIRTUAL_SURNAMES)]
    given = VIRTUAL_GIVEN[(index // 3) % len(VIRTUAL_GIVEN)]
    return f"虚拟{surname}{given}{index:03d}"


def pick_company_name(index: int) -> str:
    prefix = COMPANY_PREFIX[index % len(COMPANY_PREFIX)]
    suffix = COMPANY_SUFFIX[(index // 2) % len(COMPANY_SUFFIX)]
    return f"X{prefix}{suffix}{index:03d}"


def make_amount_in_chinese(amount: float) -> str:
    """Generate an upper-case Chinese amount string (deterministic / demo only)."""
    return f"人民币{amount:,.2f}元整".replace(",", "")


def build_rows(row_count: int = ROW_COUNT):
    start = datetime(2026, 8, 1, 8, 30)
    rows = []
    for index in range(1, row_count + 1):
        day = start + timedelta(days=index - 1)
        date_code = day.strftime("%Y%m%d")
        weigh_id = f"GB{date_code}{index:03d}"
        batch_id = f"SG{date_code}{index:03d}"
        inspect_id = f"ZJ{date_code}{index:03d}"
        inbound_id = f"RK{date_code}{index:03d}"
        settlement_id = f"JS{date_code}{index:03d}"
        invoice_id = f"FX{date_code}{index:03d}"

        company = index % 10 == 0
        supplier = pick_company_name(index) if company else pick_person_name(index)
        identity = "" if company else f"SIM{19800000 + index:08d}"
        goods_name, spec, grade, price_per_kg = GOODS_TABLE[(index - 1) % len(GOODS_TABLE)]

        gross = 10 + (index * 0.123) % 12  # 吨
        tare = 3 + (index * 0.041) % 4
        net_value = gross - tare
        deduct_rate = round(0.01 + ((index * 7) % 250) / 1000, 4)
        settlement_tonnes = round(net_value * (1 - deduct_rate), 3)
        inbound_tonnes = settlement_tonnes  # 入库实收 = 扣杂后结算
        quantity_diff = round(inbound_tonnes - settlement_tonnes, 3)
        diff_rate = round(quantity_diff / settlement_tonnes * 100, 2) if settlement_tonnes else 0

        price_per_tonne = price_per_kg * 1000
        amount = round(settlement_tonnes * price_per_tonne, 2)
        above_threshold = amount >= PAYMENT_THRESHOLD

        weigh_dt = day.replace(hour=9, minute=0)
        inspect_dt = day.replace(hour=9, minute=15)
        inbound_dt = day.replace(hour=9, minute=30)
        payment_dt = (day + timedelta(days=1)).replace(hour=10, minute=0)

        grade_value = grade
        inspector = f"虚拟质检员{['甲', '乙', '丙', '丁'][index % 4]}"
        reviewer = f"虚拟复核员{['甲', '乙', '丙', '丁'][index % 4]}"
        weigher = f"虚拟过磅员{['甲', '乙', '丙', '丁'][index % 4]}"
        driver = f"虚拟司机{['甲', '乙', '丙', '丁'][index % 4]}"

        frequency = 8 if index % 19 == 0 else (index % 5 + 1)
        annual = 6000000 if index % 37 == 0 else 300000 + index * 7000

        risk_id_no_missing = index % 29 == 0 and not company
        risk_invoice_missing = index % 31 == 0
        risk_quantity = index % 17 == 0

        rows.append({
            "台账编号": f"HG{date_code}{index:03d}",
            "业务日期": day.date().isoformat(),
            "业务类型": "企业采购" if company else "自然人收购",
            "收购工单号": batch_id,
            "供货人姓名": supplier,
            "身份证号": "" if risk_id_no_missing else identity,
            "联系电话": f"虚拟手机{13800000000 + index:08d}",
            "住址/经营场所": "虚拟经营场所" if company else "",
            "供货人类别": "小型回收站" if company else "个体散户",
            "过磅单号": weigh_id,
            "过磅时间": weigh_dt.strftime("%Y-%m-%d %H:%M"),
            "质检单号": inspect_id,
            "质检时间": inspect_dt.strftime("%Y-%m-%d %H:%M"),
            "入库单号": inbound_id,
            "入库时间": inbound_dt.strftime("%Y-%m-%d %H:%M"),
            "结算单号": settlement_id,
            "发票/收购凭证号": invoice_id if not risk_invoice_missing else "",
            "物资名称": goods_name,
            "品类规格": spec,
            "计量单位": "吨",
            "过磅毛重(吨)": round(gross, 3),
            "过磅皮重(吨)": round(tare, 3),
            "过磅净重(吨)": round(net_value, 3),
            "质检扣杂率(%)": round(deduct_rate * 100, 2),
            "扣杂后结算数量(吨)": settlement_tonnes,
            "入库实收数量(吨)": inbound_tonnes,
            "数量差异(吨)": quantity_diff,
            "差异率(%)": diff_rate,
            "收购单价(元/吨)": round(price_per_tonne, 2),
            "结算金额(元)": amount,
            "税额(元)": 0.0,
            "实付金额(元)": amount if index == 1 else 0.0,
            "金额大写": make_amount_in_chinese(amount),
            "付款日期": payment_dt.strftime("%Y-%m-%d") if index == 1 else "",
            "付款方式": "对公转账" if company else "银行卡转账",
            "付款账户": "虚拟付款账户",
            "收款户名": supplier,
            "收款账号": f"621700{(index * 7) % 10000:04d}{(index * 31) % 100000000:08d}",
            "收款开户行": BANK_BRANCHES[index % len(BANK_BRANCHES)],
            "银行回单号": f"HUI{date_code}{index:06d}",
            "是否本人收款": "是",
            "三单匹配校验": "通过" if not risk_quantity else "不通过",
            "身份核验": "不通过" if risk_id_no_missing else "通过",
            "收款账户一致性": "通过",
            "累计交易额监控": "预警" if annual >= 5_000_000 else "正常",
            "近12月累计交易额(元)": annual,
            "时间逻辑校验": "通过",
            "发票/凭证合规": "不通过" if risk_invoice_missing else "通过",
            "整体合规结论": "异常" if (risk_id_no_missing or risk_invoice_missing or risk_quantity or annual >= 5_000_000) else "合规",
            "风险标记": "高风险" if (risk_id_no_missing or risk_invoice_missing or risk_quantity or annual >= 5_000_000) else "无",
            "异常说明": "演示风险样本" if (risk_id_no_missing or risk_invoice_missing or risk_quantity or annual >= 5_000_000) else "",
            "制单人": weigher,
            "财务审核人": reviewer,
            "财务审核日期": day.date().isoformat(),
            "审批人(财务负责人)": "虚拟财务负责人" if above_threshold else "",
            "归档状态": "已归档" if index == 1 else "待归档",
            "归档日期": (day + timedelta(days=2)).date().isoformat() if index == 1 else "",
            "档案位置/编号": f"ARCHIVE/{weigh_id}",
            "附件数量": 3 + (1 if index % 17 == 0 else 0),
            "备注": "演示数据" if index == 1 else "",
            # Hidden metadata, not in the 59-column sheet:
            "_meta_grade": grade_value,
            "_meta_above_threshold": above_threshold,
            "_meta_inspector": inspector,
            "_meta_reviewer": reviewer,
            "_meta_weigher": weigher,
            "_meta_driver": driver,
            "_meta_weigh_id": weigh_id,
            "_meta_inspect_id": inspect_id,
            "_meta_inbound_id": inbound_id,
            "_meta_supplier_type": "公司" if company else "自然人",
            "_meta_registered": "是" if company else "否",
            "_meta_plate": f"虚拟车牌-{index:03d}",
            "_meta_method": "一次过磅" if index % 2 else "二次过磅",
            "_meta_frequency": frequency,
            "_meta_status": "已付款" if index == 1 else "待付款",
        })
    return rows


def write_to_template(rows, template_path: str = TEMPLATE_PATH, output_path: str = OUTPUT_PATH):
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"模板文件不存在：{template_path}")
    wb = load_workbook(template_path)
    ws = wb["合规台账"]
    if ws.max_row >= 4:
        ws.delete_rows(4, ws.max_row - 3)

    headers = [cell.value for cell in ws[3]]
    for index, row in enumerate(rows, start=4):
        for col_index, header in enumerate(headers, start=1):
            if header is None:
                continue
            ws.cell(row=index, column=col_index, value=row.get(header, ""))

    # Apply zebra striping to the 100 inserted rows and freeze the headers.
    inserted_fill = PatternFill("solid", fgColor="FBFEFB")
    center = Alignment(horizontal="left", vertical="center", wrap_text=False)
    for index in range(4, 4 + len(rows)):
        for col_index in range(1, len(headers) + 1):
            cell = ws.cell(row=index, column=col_index)
            if index % 2 == 0:
                cell.fill = inserted_fill
            cell.alignment = center
    ws.freeze_panes = "A4"
    wb.save(output_path)
    return output_path


def main():
    rows = build_rows(ROW_COUNT)
    output = write_to_template(rows)
    natural = sum(1 for row in rows if row["_meta_supplier_type"] == "自然人")
    company = ROW_COUNT - natural
    above = sum(1 for row in rows if row["_meta_above_threshold"])
    print(f"已写入 {output}")
    print(f"  共 {ROW_COUNT} 条；自然人 {natural} 笔 (90%)，公司 {company} 笔 (10%)")
    print(f"  金额 ≥ 5万元的业务 {above} 笔，将进入付款审批；其余将进入付款中心")


if __name__ == "__main__":
    main()
