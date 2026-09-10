"""生成比赛演示用的 10 条业务数据（过磅单/质检单/入库单三个工作表）。

设计目标：一份工作簿同时覆盖
  - 完全合规的小额（待付款）与大额（待审批，≥5万）
  - 低风险：复核人签字缺失、自然人身份证号缺失
  - 中风险：数量小幅不符、时间倒置、票据缺失
  - 高风险：数量大幅偏差叠加时间倒置 / 叠加个体户+反向开票

全部使用"虚拟"/"X"前缀的假数据，避免与真实主体重名。
用法：python simulate/generate_demo_workbook.py
"""
import os
import sys
from copy import copy

from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(BASE_DIR, "local_templates", "绿循智审_本地同步模板.xlsx")
OUTPUT = os.path.join(BASE_DIR, "local_templates", "绿循智审_比赛演示数据_10条.xlsx")

WEIGH_HEADERS = ["过磅单编号", "货物批次号", "过磅时间", "供货方名称", "供货方身份证号", "供货方类别",
                 "个体户登记标记", "车牌号", "货物品名", "毛重 (Kg)", "皮重 (Kg)", "净重 (Kg)",
                 "过磅方式", "过磅员签字", "司机签字", "收款户名", "银行卡号", "开户行"]
INSPECT_HEADERS = ["质检单编号", "关联过磅单编号", "质检时间", "货物品名", "货品规格", "质检等级",
                   "质检扣杂数量 (Kg)", "含水率 / 杂质率", "计价单价 (元 / Kg)", "计价金额 (元)",
                   "质检员签字", "复核人签字", "票据合规方式"]
INBOUND_HEADERS = ["入库单编号", "关联过磅单编号", "入库时间", "入库数量 (Kg)", "仓管员签字", "付款时间"]

SELF_INVOICE = "自然人自行代开发票"   # 含"代开" → 视为 self
REVERSE_INVOICE = "反向开票授权"      # 含"反向" → 视为 reverse


def record(seq, date, supplier, supplier_type, id_no, registered, goods, spec, grade, price,
           gross, tare, deduct, inbound, weigh_time, inspect_time, inbound_time, payment_time,
           invoice, reviewer="虚拟复核员甲", inspect_goods=None, note=""):
    """按业务语义组装一条记录，返回三个工作表各自的行数据。"""
    net = gross - tare
    amount = round(inbound * price, 2)
    code = date.replace("-", "")
    weigh_id = f"GB{code}{seq:03d}"
    return {
        "weigh": [weigh_id, f"PC{code}{seq:03d}", weigh_time, supplier, id_no, supplier_type,
                  registered, f"虚拟车牌-{seq:03d}", goods, gross, tare, net,
                  "一次过磅" if seq % 2 else "二次过磅", "虚拟过磅员甲", f"虚拟司机{['甲', '乙', '丙'][seq % 3]}",
                  supplier, f"6217000000000{seq:05d}", "中国工商银行虚拟支行"],
        "inspect": [f"ZJ{code}{seq:03d}", weigh_id, inspect_time, inspect_goods if inspect_goods is not None else goods,
                    spec, grade, deduct, "1.20%", price, amount, "虚拟质检员甲", reviewer, invoice],
        "inbound": [f"RK{code}{seq:03d}", weigh_id, inbound_time, inbound, "虚拟仓管员甲", payment_time],
        "note": note,
    }


RECORDS = [
    # 1 完全合规 · 小额 → 待付款
    record(1, "2026-09-01", "虚拟赵建国001", "自然人", "SIM198501010001", "否", "废钢", "重型废钢", "B级", 2.65,
           12560, 5820, 40, 6700, "2026-09-01 09:00", "2026-09-01 09:15", "2026-09-01 09:30", "2026-09-02 10:00",
           SELF_INVOICE, note="基准合规样本，金额 ¥17,755（无需审批）"),
    # 2 完全合规 · 大额 → 待审批
    record(2, "2026-09-02", "虚拟钱志远002", "自然人", "SIM198601020002", "否", "废铜", "紫杂铜", "A级", 61.0,
           4680, 1980, 0, 2700, "2026-09-02 09:00", "2026-09-02 09:15", "2026-09-02 09:30", "2026-09-03 10:00",
           REVERSE_INVOICE, note="大额合规样本，金额 ¥164,700（需财务负责人审批）"),
    # 3 完全合规 · 公司主体（无身份证号属正常，未登记个体户）
    record(3, "2026-09-03", "X蓝海资源回收有限公司003", "公司", "", "否", "废铝", "破碎杂铝", "C级", 13.2,
           9850, 4650, 200, 5000, "2026-09-03 09:00", "2026-09-03 09:15", "2026-09-03 09:30", "2026-09-04 10:00",
           REVERSE_INVOICE, note="公司主体合规样本，金额 ¥66,000"),
    # 4 低风险：复核人签字缺失（R07）
    record(4, "2026-09-04", "虚拟孙宏达004", "自然人", "SIM198701040004", "否", "废电机", "混合电机", "B级", 5.5,
           8200, 3100, 100, 5000, "2026-09-04 09:00", "2026-09-04 09:15", "2026-09-04 09:30", "2026-09-05 10:00",
           SELF_INVOICE, reviewer="", note="低风险：复核人签字留空（R07，4 分）"),
    # 5 低风险：自然人身份证号缺失（R03）
    record(5, "2026-09-05", "虚拟李明轩005", "自然人", "", "否", "废钢", "重型废钢", "B级", 2.65,
           11000, 4000, 50, 6950, "2026-09-05 09:00", "2026-09-05 09:15", "2026-09-05 09:30", "2026-09-06 10:00",
           SELF_INVOICE, note="低风险：身份证号留空（R03，8 分）"),
    # 6 中风险：数量链偏差 35Kg（R02）
    record(6, "2026-09-06", "虚拟周国良006", "自然人", "SIM198801060006", "否", "废铝", "破碎杂铝", "C级", 13.2,
           9850, 4650, 235, 5000, "2026-09-06 09:00", "2026-09-06 09:15", "2026-09-06 09:30", "2026-09-07 10:00",
           SELF_INVOICE, note="中风险：净重-扣杂 与入库差 35Kg（R02，15 分）"),
    # 7 中风险：质检时间早于过磅（R04）
    record(7, "2026-09-07", "虚拟吴永福007", "自然人", "SIM198901070007", "否", "废铜", "紫杂铜", "A级", 61.0,
           4680, 1980, 0, 2700, "2026-09-07 09:00", "2026-09-07 08:30", "2026-09-07 09:30", "2026-09-08 10:00",
           REVERSE_INVOICE, note="中风险：质检早于过磅（R04，12 分）"),
    # 8 中风险：票据合规方式缺失（R08）
    record(8, "2026-09-08", "虚拟郑立军008", "自然人", "SIM199001080008", "否", "废铁", "重型废钢", "B级", 2.65,
           11000, 4000, 100, 6900, "2026-09-08 09:00", "2026-09-08 09:15", "2026-09-08 09:30", "2026-09-09 10:00",
           "", note="中风险：票据合规方式留空（R08，10 分）"),
    # 9 高风险：数量大幅偏差 1200Kg + 时间倒置（R02 + R04）
    record(9, "2026-09-09", "虚拟王秀英009", "自然人", "SIM199101090009", "否", "废钢", "重型废钢", "B级", 2.65,
           12560, 5820, 1240, 6700, "2026-09-09 09:00", "2026-09-09 08:30", "2026-09-09 09:30", "2026-09-10 10:00",
           SELF_INVOICE, note="高风险：数量偏差 1200Kg + 时间倒置（42 分）"),
    # 10 高风险：数量大幅偏差 + 个体户登记却走反向开票（R02 + R09）
    record(10, "2026-09-10", "虚拟冯丽华010", "自然人", "SIM199201100010", "是", "废铝", "破碎杂铝", "C级", 13.2,
           9850, 4650, 1400, 5000, "2026-09-10 09:00", "2026-09-10 09:15", "2026-09-10 09:30", "2026-09-11 10:00",
           REVERSE_INVOICE, note="高风险：数量偏差 1200Kg + 个体户反向开票（36 分）"),
]


def write_sheet(ws, headers, rows):
    """按模板样式写入表头与数据行。"""
    template_row = 2 if ws.max_row >= 2 else 1
    for column, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=column, value=header)
        source = ws.cell(row=1, column=column)
        if source.value is None or source.font is None:
            continue
        cell.font = copy(source.font)
    for offset, values in enumerate(rows):
        row_index = 2 + offset
        for column, value in enumerate(values, start=1):
            cell = ws.cell(row=row_index, column=column, value=value)
            style_source = ws.cell(row=template_row, column=column)
            cell.font = copy(style_source.font)
            cell.fill = copy(style_source.fill)
            cell.border = copy(style_source.border)
            cell.alignment = copy(style_source.alignment)
            cell.number_format = style_source.number_format


def main():
    if not os.path.exists(TEMPLATE):
        print(f"模板不存在：{TEMPLATE}", file=sys.stderr)
        return 1
    workbook = load_workbook(TEMPLATE)
    weigh_rows = [item["weigh"] for item in RECORDS]
    inspect_rows = [item["inspect"] for item in RECORDS]
    inbound_rows = [item["inbound"] for item in RECORDS]
    write_sheet(workbook["过磅单"], WEIGH_HEADERS, weigh_rows)
    write_sheet(workbook["质检单"], INSPECT_HEADERS, inspect_rows)
    write_sheet(workbook["入库单"], INBOUND_HEADERS, inbound_rows)
    # 清掉模板自带的示例行之外的残留（本模板每个工作表仅 1 行示例）
    for name, count in (("过磅单", len(weigh_rows)), ("质检单", len(inspect_rows)), ("入库单", len(inbound_rows))):
        ws = workbook[name]
        if ws.max_row > count + 1:
            ws.delete_rows(count + 2, ws.max_row - count - 1)
    workbook.save(OUTPUT)
    print(f"已生成：{OUTPUT}")
    print(f"过磅单 {len(weigh_rows)} 条 / 质检单 {len(inspect_rows)} 条 / 入库单 {len(inbound_rows)} 条")
    for item in RECORDS:
        print(f"  {item['weigh'][0]}  {item['weigh'][3]:<28} {item['note']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
