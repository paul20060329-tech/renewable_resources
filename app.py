import base64
import hashlib
import io
import json
import os
import sqlite3
from datetime import datetime, timedelta

from flask import Flask, jsonify, render_template, request, send_file, session, redirect, url_for
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "greenloop-demo-change-before-production")
DB_PATH = os.environ.get("DATABASE_PATH", "greenloop.db")

RULES = [
    ("R01", "单据关联一致性"), ("R02", "重量偏差校验"), ("R03", "自然人身份完整性"),
    ("R04", "单据时间逻辑"), ("R05", "等级计价匹配"), ("R06", "短期高频交易"),
    ("R07", "原始单据完整性"), ("R08", "票据合规二选一"), ("R09", "主体资格预警"),
]

PAYMENT_THRESHOLD = 50000
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_TEMPLATE_DIR = os.path.join(BASE_DIR, "local_templates")
LEDGER_TEMPLATE_PATH = os.path.join(BASE_DIR, "Data", "再生资源公司收购合规台账模板.xlsx")
# 按当前用户提交的权限矩阵整理：
#  - 财务负责人：只能看、驳回、付款审批（≥5万），不能修改单据；
#  - 会计：导入单据与核验，不能付款；
#  - 出纳：只对≤5万且已通过的业务执行付款，不参与核验与驳回；
#  - 后台管理员：拥有全部权限。
ROLE_PERMISSIONS = {
    "财务负责人": {"overview", "approval", "ledger"},
    "会计": {"overview", "ingest", "verify", "ledger"},
    "出纳": {"overview", "payment", "ledger"},
    "后台管理员": {"overview", "ingest", "verify", "approval", "payment", "ledger", "admin"},
}
ACCOUNTS = {
    "finance": ("123456", "财务负责人", "财务负责人"),
    "accounting": ("123456", "会计", "会计"),
    "cashier": ("123456", "出纳", "出纳"),
    "admin": ("123456", "后台管理员", "后台管理员"),
}

WEIGH_HEADERS = ["过磅单编号", "货物批次号", "过磅时间", "供货方名称", "供货方身份证号", "供货方类别", "个体户登记标记", "车牌号", "货物品名", "毛重 (Kg)", "皮重 (Kg)", "净重 (Kg)", "过磅方式", "过磅员签字", "司机签字", "收款户名", "银行卡号", "开户行"]
INSPECT_HEADERS = ["质检单编号", "关联过磅单编号", "质检时间", "货物品名", "货品规格", "质检等级", "质检扣杂数量 (Kg)", "含水率 / 杂质率", "计价单价 (元 / Kg)", "计价金额 (元)", "质检员签字", "复核人签字", "票据合规方式"]
INBOUND_HEADERS = ["入库单编号", "关联过磅单编号", "入库时间", "入库数量 (Kg)", "仓管员签字", "付款时间"]

# 大量虚拟姓名（仅用作演示，明显以"虚拟"/"X"前缀避免与真实姓名混用）。
VIRTUAL_SURNAMES = ["赵", "钱", "孙", "李", "周", "吴", "郑", "王", "冯", "陈", "褚", "卫", "蒋", "沈", "韩"]
VIRTUAL_GIVEN_NAMES = ["建国", "志远", "宏达", "明轩", "国良", "永福", "立军", "志伟", "秀英", "丽华", "文军", "秀梅", "建华", "彩霞", "海涛"]
VIRTUAL_COMPANY_PREFIX = ["蓝海", "绿野", "银拓", "钢城", "循环", "环保", "资源", "循环再生", "金属再生", "物资回收"]
VIRTUAL_COMPANY_SUFFIX = ["资源回收有限公司", "再生金属有限公司", "环保科技有限公司", "废旧物资回收部", "循环利用公司"]
VIRTUAL_BANK_BRANCHES = ["中国工商银行虚拟支行", "中国农业银行虚拟分行", "中国建设银行虚拟支行", "中国银行虚拟分行", "虚拟商业银行支行"]

def _pick_virtual_person_name(index):
    """不依赖真实姓名的伪中文姓名；保证只生成演示用假数据。"""
    surname = VIRTUAL_SURNAMES[index % len(VIRTUAL_SURNAMES)]
    given = VIRTUAL_GIVEN_NAMES[(index // 3) % len(VIRTUAL_GIVEN_NAMES)]
    return f"虚拟{surname}{given}{index:03d}"

def _pick_virtual_company_name(index):
    """明显的虚拟公司命名，避免与任何真实企业同名。"""
    prefix = VIRTUAL_COMPANY_PREFIX[index % len(VIRTUAL_COMPANY_PREFIX)]
    suffix = VIRTUAL_COMPANY_SUFFIX[(index // 2) % len(VIRTUAL_COMPANY_SUFFIX)]
    return f"X{prefix}{suffix}{index:03d}"


DEMO_ROWS = [
    {"id":"GB202608250001","batch":"PC202608250001","supplier":"虚拟自然人赵建国001","id_no":"SIM198501010001","supplier_type":"自然人","registered":"否","plate":"虚拟车牌-001","goods":"废不锈钢","gross":12560,"tare":5820,"net":6740,"method":"一次过磅","weigher":"过磅员甲","driver":"司机甲","weigh_date":"2026-08-25","weigh_time":"2026-08-25 14:00","inspect_id":"ZJ202608250001","inspect_date":"2026-08-25","inspect_time":"2026-08-25 14:15","inbound_id":"RK202608250001","inbound_date":"2026-08-25","inbound_time":"2026-08-25 14:30","payment_time":"2026-08-26 10:00","spec":"一级统料","grade":"A级","inspect_qty":35,"inbound_qty":6705,"moisture":0.52,"price":9.8,"amount":65709,"inspector":"质检员甲","reviewer":"复核员甲","invoice":"reverse","annual":382000,"frequency":2,"bank_name":"虚拟自然人赵建国001","bank_account":"621700000000000001","bank_branch":"中国工商银行虚拟支行","status":"已付款","approval":"无需审批","paid_amount":65709},
    {"id":"GB202608290002","batch":"PC202608290002","supplier":"虚拟自然人钱志远002","id_no":"","supplier_type":"自然人","registered":"是","plate":"虚拟车牌-002","goods":"废杂铝合金","gross":9850,"tare":4650,"net":5200,"method":"二次过磅","weigher":"","driver":"司机乙","weigh_date":"2026-08-29","weigh_time":"2026-08-29 09:00","inspect_id":"ZJ202608290002","inspect_date":"2026-08-29","inspect_time":"2026-08-29 08:30","inbound_id":"RK202608290002","inbound_date":"2026-08-29","inbound_time":"2026-08-29 10:00","payment_time":"2026-08-30 10:00","spec":"破碎杂铝","grade":"C级","inspect_qty":900,"inbound_qty":4300,"moisture":8.7,"price":13.2,"amount":56760,"inspector":"质检员乙","reviewer":"","invoice":"none","annual":4960000,"frequency":8,"bank_name":"虚拟自然人钱志远002","bank_account":"621700000000000002","bank_branch":"中国农业银行虚拟分行","status":"高风险","approval":"待逐项风险处置","paid_amount":0},
    {"id":"GB202609030003","batch":"PC202609030003","supplier":"虚拟自然人孙宏达003","id_no":"SIM198611040003","supplier_type":"自然人","registered":"否","plate":"虚拟车牌-003","goods":"废铁","gross":18320,"tare":7320,"net":11000,"method":"二次过磅","weigher":"过磅员乙","driver":"司机丙","weigh_date":"2026-09-03","weigh_time":"2026-09-03 10:00","inspect_id":"ZJ202609030003","inspect_date":"2026-09-03","inspect_time":"2026-09-03 10:15","inbound_id":"RK202609030003","inbound_date":"2026-09-03","inbound_time":"2026-09-03 10:30","payment_time":"2026-09-03 16:00","spec":"重型废钢","grade":"B级","inspect_qty":0,"inbound_qty":11000,"moisture":1.4,"price":2.65,"amount":29150,"inspector":"质检员丙","reviewer":"复核员乙","invoice":"self","annual":684000,"frequency":3,"bank_name":"虚拟自然人孙宏达003","bank_account":"621700000000000003","bank_branch":"中国建设银行虚拟支行","status":"待核验","approval":"无需审批","paid_amount":0},
    {"id":"GB202609040004","batch":"PC202609040004","supplier":"虚拟自然人李明轩004","id_no":"SIM197906180004","supplier_type":"自然人","registered":"否","plate":"虚拟车牌-004","goods":"废铜","gross":4680,"tare":1980,"net":2700,"method":"一次过磅","weigher":"过磅员甲","driver":"司机丁","weigh_date":"2026-09-04","weigh_time":"2026-09-04 11:00","inspect_id":"ZJ202609040004","inspect_date":"2026-09-04","inspect_time":"2026-09-04 11:15","inbound_id":"RK202609040004","inbound_date":"2026-09-04","inbound_time":"2026-09-04 11:30","payment_time":"2026-09-05 10:00","spec":"紫杂铜","grade":"A级","inspect_qty":0,"inbound_qty":2700,"moisture":0.2,"price":61.0,"amount":164700,"inspector":"质检员甲","reviewer":"复核员甲","invoice":"reverse","annual":1180000,"frequency":4,"bank_name":"虚拟自然人李明轩004","bank_account":"621700000000000004","bank_branch":"中国工商银行虚拟支行","status":"高风险","approval":"待逐项风险处置","paid_amount":0},
    {"id":"GB202609050005","batch":"PC202609050005","supplier":"X蓝海资源回收有限公司005","id_no":"","supplier_type":"公司","registered":"是","plate":"虚拟车牌-005","goods":"废电机","gross":8200,"tare":3100,"net":5100,"method":"一次过磅","weigher":"过磅员乙","driver":"司机戊","weigh_date":"2026-09-05","weigh_time":"2026-09-05 13:00","inspect_id":"ZJ202609050005","inspect_date":"2026-09-05","inspect_time":"2026-09-05 13:15","inbound_id":"RK202609050005","inbound_date":"2026-09-05","inbound_time":"2026-09-05 13:30","payment_time":"2026-09-05 17:00","spec":"混合电机","grade":"B级","inspect_qty":0,"inbound_qty":5100,"moisture":0.8,"price":5.5,"amount":28050,"inspector":"质检员乙","reviewer":"复核员乙","invoice":"reverse","annual":28050,"frequency":1,"bank_name":"X蓝海资源回收有限公司005","bank_account":"621700000000000005","bank_branch":"中国农业银行虚拟分行","status":"待核验","approval":"无需审批","paid_amount":0},
]


def build_demo_rows(count=100):
    """Create deterministic, fully fictitious demo data for a clean install/reset.

    - 90% 自然人 / 10% 公司
    - 全部使用明显"虚拟"/"X"前缀的姓名/单位/银行/车牌，避免与任何真实记录同名
    - 时间链：过磅 < 质检 < 入库 < 付款；数量链：净重 - 扣杂 = 入库数量
    - 一条已付款示例 + 若干风险/审批/待付款示例
    """
    goods = [("废钢", "重型废钢", 2.65), ("废铜", "紫杂铜", 61.0), ("废铝", "破碎杂铝", 13.2), ("废电机", "混合电机", 5.5)]
    rows = []
    start = datetime(2026, 8, 1, 9, 0)
    for index in range(1, count + 1):
        day = start + timedelta(days=index - 1)
        date_code = day.strftime("%Y%m%d")
        business_id = f"GB{date_code}{index:03d}"
        batch_id = f"PC{date_code}{index:03d}"
        # 自然人：index % 10 != 0 → 90%
        company = index % 10 == 0
        supplier_type = "公司" if company else "自然人"
        supplier = _pick_virtual_company_name(index) if company else _pick_virtual_person_name(index)
        identity = "" if company else f"SIM{19800000 + index:08d}"
        goods_name, spec, price = goods[(index - 1) % len(goods)]
        gross = 7000 + (index * 431) % 12000
        tare = 2200 + (index * 137) % 4200
        net = gross - tare
        deduct = (index * 13) % 180
        inbound_qty = net - deduct
        amount = round(inbound_qty * price, 2)
        weigh_time = day.strftime("%Y-%m-%d 09:00")
        inspect_time = day.strftime("%Y-%m-%d 09:15")
        inbound_time = day.strftime("%Y-%m-%d 09:30")
        payment_time = (day + timedelta(days=1)).strftime("%Y-%m-%d 10:00")
        # 风险命中刻意做得多样化：有的 1 项轻微异常（低风险），
        # 有的 2~3 项叠加（中/高风险）；index%40 的记录数量+时间同时出错。
        risk_quantity_small = index % 17 == 0          # 扣杂小幅不符（~35Kg）
        risk_quantity_large = index % 40 == 0 or index % 70 == 0   # 数量链大幅偏差（~1200Kg）→ 高风险
        risk_time = index % 23 == 0 or risk_quantity_large
        risk_identity = index % 29 == 0 and not company
        risk_invoice = index % 31 == 0
        risk_review = index % 13 == 0 and index % 40 != 0   # 复核签字缺失 + 高频交易 → 中风险
        risk_price = index % 47 == 0                        # 计价金额与数量×单价不符
        bank_branch = VIRTUAL_BANK_BRANCHES[index % len(VIRTUAL_BANK_BRANCHES)]
        row = {
            "id": business_id, "batch": batch_id, "supplier": supplier, "supplier_type": supplier_type,
            "id_no": identity, "registered": "是" if company else "否", "plate": f"虚拟车牌-{index:03d}",
            "goods": goods_name, "gross": gross, "tare": tare, "net": net, "method": "一次过磅" if index % 2 else "二次过磅",
            "weigher": f"虚拟过磅员{['甲', '乙', '丙', '丁'][index % 4]}", "driver": f"虚拟司机{['甲', '乙', '丙', '丁'][index % 4]}",
            "weigh_date": day.date().isoformat(), "weigh_time": weigh_time,
            "inspect_id": f"ZJ{date_code}{index:03d}", "inspect_date": day.date().isoformat(),
            "inspect_time": (day - timedelta(minutes=30)).strftime("%Y-%m-%d %H:%M") if risk_time else inspect_time,
            "inbound_id": f"RK{date_code}{index:03d}", "inbound_date": day.date().isoformat(),
            "inbound_time": inbound_time, "payment_time": payment_time,
            "spec": spec, "grade": "A级" if index % 3 == 0 else "B级",
            "inspect_qty": deduct + (1200 if risk_quantity_large else (35 if risk_quantity_small else 0)), "inbound_qty": inbound_qty,
            "settlement_qty": inbound_qty, "moisture": 0.8,
            "inspect_goods": goods_name, "inbound_goods": goods_name,
            "price": price, "amount": amount,
            "inspector": f"虚拟质检员{['甲', '乙', '丙', '丁'][index % 4]}",
            "reviewer": f"虚拟复核员{['甲', '乙', '丙', '丁'][index % 4]}",
            "invoice": "none" if risk_invoice else ("reverse" if company or index % 2 else "self"),
            "annual": 6000000 if index % 37 == 0 else 300000 + index * 7000,
            "frequency": 8 if (index % 19 == 0 or index % 13 == 0) else (index % 5 + 1),
            "bank_name": supplier, "bank_account": f"621700{(index * 7) % 10000:04d}{(index * 31) % 100000000:08d}",
            "bank_branch": bank_branch,
            "status": "待核验", "approval": "无需审批", "verified": False, "review_cycle": 1,
            "risk_decisions": {}, "rejection_history": [], "financial_approved": False, "paid_amount": 0,
        }
        if risk_identity:
            row["id_no"] = ""
        if risk_review:
            row["reviewer"] = ""
        if risk_price:
            row["amount"] = round(row["amount"] + 2400, 2)
        if index == 1:
            row["status"], row["verified"], row["paid_amount"] = "已付款", True, amount
        elif (
            risk_quantity_small or risk_quantity_large or risk_time or risk_identity or risk_invoice
            or risk_review or risk_price
            or row["annual"] >= 5000000 or row["frequency"] > 6
            or (row["registered"] == "是" and row["invoice"] == "reverse")
        ):
            row["status"], row["approval"] = "高风险", "待逐项风险处置"
        else:
            row["verified"] = True
            row["status"] = "待审批" if amount >= PAYMENT_THRESHOLD else "待付款"
            row["approval"] = "待财务负责人审批" if amount >= PAYMENT_THRESHOLD else "无需审批"
        rows.append(row)
    return rows


# 演示数据：90% 自然人 / 10% 公司，全部虚拟。
DEMO_ROWS = build_demo_rows(100)

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def sha256_fingerprint(value):
    return hashlib.sha256(str(value).encode("utf-8")).hexdigest()

def add_security_fields(row):
    canonical = {
        key: row.get(key, "")
        for key in (
            "id", "batch", "weigh_date", "supplier", "id_no", "registered",
            "plate", "goods", "gross", "tare", "net", "method", "weigher",
            "driver",
        )
    }
    serialized = json.dumps(canonical, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    row["id_hash"] = row.get("id_hash") or sha256_fingerprint(row.get("id", ""))
    row["batch_hash"] = row.get("batch_hash") or sha256_fingerprint(row.get("batch", ""))
    row["integrity_hash"] = sha256_fingerprint(serialized)
    row["hash_algorithm"] = "SHA-256"
    return row

def ensure_workflow_fields(row, status=None, approval=None):
    """Keep workflow state in the payload so old SQLite records can be upgraded safely."""
    defaults = {
        "supplier_type": "自然人", "inbound_id": "", "inbound_date": row.get("weigh_date", ""),
        "inbound_time": "", "payment_time": "", "inbound_qty": 0,
        # Legacy payloads did not store cross-sheet goods names; use the
        # weighing goods for those records while new imports keep their own
        # values so mismatches are detected by R01.
        "inspect_goods": row.get("goods", ""), "inbound_goods": row.get("goods", ""),
        "settlement_qty": row.get("inbound_qty", 0),
        "bank_name": row.get("supplier", ""),
        "bank_account": "", "bank_branch": "", "inspect_id": "", "inspect_date": row.get("weigh_date", ""),
        "inspect_time": "", "inspect_qty": 0, "rejection_history": [], "risk_decisions": {},
    }
    for key, value in defaults.items():
        if row.get(key) is None:
            row[key] = value
    current_status = status or row.get("status", "待核验")
    row["verified"] = bool(row.get("verified", current_status in ("已核验", "已付款")))
    row["verification_status"] = "已核验" if row["verified"] else "待核验"
    row.setdefault("review_cycle", 1)
    if not isinstance(row.get("risk_decisions"), dict): row["risk_decisions"] = {}
    if not isinstance(row.get("rejection_history"), list): row["rejection_history"] = []
    row.setdefault("financial_approved", False)
    row.setdefault("paid_amount", row.get("amount", 0) if current_status == "已付款" else 0)
    if approval is not None:
        row["approval"] = approval
    row["status"] = current_status
    row.setdefault("approval", "无需审批")
    return row

def role_can(permission):
    user = session.get("user") or {}
    return permission in ROLE_PERMISSIONS.get(user.get("role"), set())

def permission_error(permission):
    return jsonify({"error": f"当前账号没有{permission}权限"}), 403

def require_any_permission(*permissions):
    """Allow a workflow action to be performed by any of the listed roles."""
    if any(role_can(permission) for permission in permissions):
        return None
    return permission_error("或".join(permissions))

def masked_bank(value):
    value = str(value or "")
    return value if len(value) <= 8 else f"{value[:4]}****{value[-4:]}"

def parse_datetime(value, fallback=""):
    if isinstance(value, datetime):
        return value
    text = str(value or fallback).strip().replace("/", "-")
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y-%m-%dT%H:%M"):
        try: return datetime.strptime(text[:19], fmt)
        except ValueError: pass
    return datetime.min

def fmt_datetime(value, fallback=""):
    dt = parse_datetime(value, fallback)
    return dt.strftime("%Y-%m-%d %H:%M") if dt != datetime.min else (str(value or fallback)[:16])

# Keep a complete source-document snapshot with every rejection.  The ledger
# drawer can then show the exact values before rejection and after resubmission,
# instead of only a handful of summary fields.
SOURCE_SNAPSHOT_FIELDS = (
    "id", "batch", "supplier", "supplier_type", "id_no", "registered", "plate", "goods",
    "gross", "tare", "net", "method", "weigher", "driver", "weigh_date", "weigh_time",
    "inspect_id", "inspect_date", "inspect_time", "inspect_goods", "spec", "grade", "inspect_qty", "inbound_qty", "settlement_qty",
    "inbound_id", "inbound_date", "inbound_time", "payment_time", "moisture", "price", "amount",
    "inbound_goods", "inspector", "reviewer", "invoice", "bank_name", "bank_account", "bank_branch", "annual", "frequency",
)

def source_snapshot(row):
    return {key: row.get(key, "") for key in SOURCE_SNAPSHOT_FIELDS}

def rejection_entry(row, code, opinion):
    now = datetime.now().isoformat(timespec="seconds")
    snapshot = source_snapshot(row)
    return {
        "cycle": int(row.get("review_cycle", 1) or 1),
        "rejected_code": code,
        "opinion": opinion,
        "original": snapshot,
        "original_documents": snapshot,
        "rejected_at": now,
    }

def generate_document_ids():
    date_code = datetime.now().strftime("%Y%m%d")
    con = db()
    index = 1
    while True:
        suffix = f"{index:04d}"
        business_id = f"GB{date_code}{suffix}"
        batch_id = f"PC{date_code}{suffix}"
        exists = con.execute("SELECT 1 FROM businesses WHERE id=?", (business_id,)).fetchone()
        if not exists:
            con.close()
            return business_id, batch_id
        index += 1

def init_db():
    con = db()
    con.execute('''CREATE TABLE IF NOT EXISTS businesses (id TEXT PRIMARY KEY, payload TEXT NOT NULL, status TEXT NOT NULL, approval TEXT NOT NULL, updated_at TEXT NOT NULL)''')
    con.execute('''CREATE TABLE IF NOT EXISTS import_jobs (id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT NOT NULL, rows_count INTEGER NOT NULL, result TEXT NOT NULL, created_at TEXT NOT NULL)''')
    con.execute('''CREATE TABLE IF NOT EXISTS local_template_state (path TEXT PRIMARY KEY, mtime REAL NOT NULL)''')
    os.makedirs(LOCAL_TEMPLATE_DIR, exist_ok=True)
    # A template generated by a previous run is a download artifact, not a new
    # business import.  Seed its fingerprint so the first dashboard refresh on
    # a clean database does not duplicate that sample row.  A later edit still
    # changes the fingerprint and is picked up automatically.
    for name in os.listdir(LOCAL_TEMPLATE_DIR):
        path = os.path.join(LOCAL_TEMPLATE_DIR, name)
        if not os.path.isfile(path) or not name.lower().endswith(".xlsx"):
            continue
        try:
            fingerprint = f"{os.stat(path).st_mtime_ns}:{os.path.getsize(path)}"
            con.execute("INSERT OR IGNORE INTO local_template_state(path,mtime) VALUES(?,?)", (path, fingerprint))
        except OSError:
            continue
    if not con.execute("SELECT 1 FROM businesses LIMIT 1").fetchone():
        now = datetime.now().isoformat(timespec="seconds")
        for item in DEMO_ROWS:
            ensure_workflow_fields(item)
            add_security_fields(item)
            con.execute("INSERT INTO businesses VALUES (?,?,?,?,?)", (item["id"], json.dumps(item, ensure_ascii=False), item["status"], item["approval"], now))
        con.commit()
    else:
        # One-time compatibility migration for databases created before workflow states existed.
        for saved in con.execute("SELECT id,payload,status,approval FROM businesses").fetchall():
            payload = json.loads(saved[1]); changed = False
            before = json.dumps(payload, ensure_ascii=False, sort_keys=True)
            ensure_workflow_fields(payload, saved[2], saved[3])
            if saved[2] == "合规通过" and not payload["verified"]:
                payload["status"] = "待核验"
            if payload["status"] == "高风险" and payload["approval"] in ("例外审批通过", "待例外审批"):
                payload["approval"] = "待逐项风险处置"
            changed = before != json.dumps(payload, ensure_ascii=False, sort_keys=True)
            if changed:
                con.execute("UPDATE businesses SET payload=?,status=?,approval=?,updated_at=? WHERE id=?", (json.dumps(payload, ensure_ascii=False), payload["status"], payload["approval"], datetime.now().isoformat(timespec="seconds"), saved[0]))
        con.commit()
    con.close()

def reset_demo_data():
    con = db()
    con.execute("DELETE FROM businesses")
    con.execute("DELETE FROM import_jobs")
    # Resetting demo data must also treat the current downloaded workbooks as
    # already-seen artifacts.  Otherwise a template generated during an
    # earlier session can be imported again on the next dashboard refresh.
    for name in os.listdir(LOCAL_TEMPLATE_DIR):
        path = os.path.join(LOCAL_TEMPLATE_DIR, name)
        if not os.path.isfile(path) or not name.lower().endswith(".xlsx"):
            continue
        try:
            fingerprint = f"{os.stat(path).st_mtime_ns}:{os.path.getsize(path)}"
            con.execute("INSERT OR REPLACE INTO local_template_state(path,mtime) VALUES(?,?)", (path, fingerprint))
        except OSError:
            continue
    now = datetime.now().isoformat(timespec="seconds")
    for item in DEMO_ROWS:
        ensure_workflow_fields(item)
        add_security_fields(item)
        con.execute("INSERT INTO businesses VALUES (?,?,?,?,?)", (item["id"], json.dumps(item, ensure_ascii=False), item["status"], item["approval"], now))
    con.commit()
    con.close()

def check(row):
    gross = float(row.get("gross", 0) or 0)
    tare = float(row.get("tare", 0) or 0)
    net = float(row.get("net", 0) or 0)
    goods = str(row.get("goods", "") or "").strip()
    inspect_goods_value = row.get("inspect_goods")
    # Legacy/demo payloads may not have a separate quality-goods field.  An
    # explicitly blank imported field remains blank and correctly fails R01.
    inspect_goods = goods if inspect_goods_value is None else str(inspect_goods_value).strip()
    inbound_goods = str(row.get("inbound_goods", "") or goods).strip()
    net_ok = gross > tare >= 0 and abs(net - (gross - tare)) < 1
    relation = (
        bool(row.get("id"))
        and bool(row.get("batch"))
        and row.get("batch", "")[-3:] == row.get("id", "")[-3:]
        and bool(row.get("inspect_id"))
        and bool(row.get("inbound_id"))
        and bool(goods)
        and inspect_goods == goods
        and inbound_goods == goods
    )
    # 新规则：过磅净重 - 质检扣杂数量 = 入库数量。
    inbound_qty = clean_number(row.get("inbound_qty"))
    inspect_qty = float(row.get("inspect_qty", 0) or 0)
    quantity_ok = inspect_qty >= 0 and inbound_qty > 0 and abs(net - inspect_qty - inbound_qty) < 1
    price = float(row.get("price", 0) or 0)
    amount = float(row.get("amount", 0) or 0)
    settlement_value = row.get("settlement_qty")
    settlement_qty = clean_number(settlement_value, inbound_qty)
    if settlement_qty <= 0:
        settlement_qty = inbound_qty
    amount_ok = (
        settlement_qty > 0
        and price > 0
        and amount > 0
        and abs(amount - settlement_qty * price) < 1
        and row.get("grade") in ["A级", "B级", "C级"]
    )
    weigh_dt, inspect_dt, inbound_dt, payment_dt = [parse_datetime(row.get(k)) for k in ("weigh_time", "inspect_time", "inbound_time", "payment_time")]
    time_ok = bool(row.get("inbound_id")) and all(dt != datetime.min for dt in (weigh_dt, inspect_dt, inbound_dt, payment_dt)) and weigh_dt < inspect_dt < inbound_dt < payment_dt
    items = [
        ("R01", relation, "批次号、过磅/质检/入库关联及货物品名已交叉匹配" if relation else "批次号、关联单据或货物品名不一致"),
        ("R02", net_ok and quantity_ok, f"净重 {net:,} - 扣杂 {inspect_qty:,} = 入库 {inbound_qty:,} Kg" if quantity_ok else f"数量链不平：净重 {net:,.0f} - 扣杂 {inspect_qty:,.0f} ≠ 入库 {inbound_qty:,.0f} Kg，偏差 {abs(net - inspect_qty - inbound_qty):,.0f} Kg"),
        ("R03", row.get("supplier_type", "自然人") == "公司" or (bool(row.get("id_no")) and len(str(row.get("id_no"))) >= 8), "主体身份信息完整" if row.get("supplier_type", "自然人") == "公司" or row.get("id_no") else "缺少自然人身份证明"),
        ("R04", time_ok, "过磅<质检<入库<付款" if time_ok else "过磅、质检、入库、付款时间顺序不合规或存在缺失"),
        ("R05", amount_ok, f"结算数量 {settlement_qty:,} Kg × ¥{price:.2f} = ¥{amount:,.2f}" if amount_ok else "计价金额与结算数量、单价或等级不匹配"),
        ("R06", float(row.get("frequency", 0) or 0) <= 6, f"近30日交易 {row.get('frequency', 0)} 笔" if float(row.get("frequency", 0) or 0) <= 6 else f"近30日交易 {row.get('frequency', 0)} 笔，超过预警阈值"),
        ("R07", bool(row.get("weigher") and row.get("inspector") and row.get("reviewer")), "过磅、质检、复核留痕完整" if row.get("weigher") and row.get("inspector") and row.get("reviewer") else "原始单据签字或复核信息缺失"),
        ("R08", row.get("invoice") in ["reverse", "self"], "已完成反向开票授权" if row.get("invoice") == "reverse" else "已留存自然人代开发票" if row.get("invoice") == "self" else "未满足反向开票或自然人代开任一条件"),
        ("R09", not (row.get("registered") == "是" and row.get("invoice") == "reverse") and float(row.get("annual", 0) or 0) < 5000000, f"年度累计 ¥{float(row.get('annual', 0) or 0):,.0f}，主体资格正常" if float(row.get("annual", 0) or 0) < 4500000 else f"年度累计 ¥{float(row.get('annual', 0) or 0):,.0f}，临近500万阈值"),
    ]
    return [{"code": code, "name": dict(RULES)[code], "passed": bool(passed), "reason": reason} for code, passed, reason in items]

# 风险分级：按规则权重累加评分，再按偏差幅度加成，映射为高/中/低风险。
#  - 高风险：评分 ≥ 30，或数量链偏差 ≥ 1000 Kg；
#  - 中风险：评分 ≥ 10，或同时命中 2 项以上异常；
#  - 低风险：仅命中 1 项轻微异常（签字缺失、交易频次偏高等）。
RULE_WEIGHTS = {"R01": 10, "R02": 15, "R03": 8, "R04": 12, "R05": 8, "R06": 5, "R07": 4, "R08": 10, "R09": 6}

def assess_risk(checks, row):
    failed = [c for c in checks if not c["passed"]]
    if not failed:
        return {"level": "无风险", "score": 0, "failed_count": 0, "summary": "9项规则全部通过"}
    score = sum(RULE_WEIGHTS.get(c["code"], 5) for c in failed)
    # 数量链偏差越大越危险：净重 - 扣杂 与入库数量的差额。
    net = clean_number(row.get("net"))
    inspect_qty = clean_number(row.get("inspect_qty"))
    inbound_qty = clean_number(row.get("inbound_qty"))
    qty_dev = abs(net - inspect_qty - inbound_qty)
    qty_bonus = 0
    if qty_dev >= 1000: qty_bonus = 15
    elif qty_dev >= 100: qty_bonus = 8
    score += qty_bonus
    # 年度累计逼近/超过 500 万 → 主体资格风险加成。
    if clean_number(row.get("annual")) >= 5000000: score += 8
    # 短期高频交易明显超标 → 加成。
    if clean_number(row.get("frequency")) > 10: score += 5
    if len(failed) >= 2 and score < 10: score = 10  # 多项异常至少中风险
    if qty_dev >= 1000 or score >= 30:
        level = "高风险"
    elif score >= 10:
        level = "中风险"
    else:
        level = "低风险"
    summary = "{}项异常：{}".format(len(failed), "、".join("{} {}".format(c["code"], c["name"]) for c in failed))
    return {"level": level, "score": min(int(score), 100), "failed_count": len(failed), "summary": summary}

def get_rows():
    con = db(); saved = con.execute("SELECT payload,status,approval,updated_at FROM businesses ORDER BY updated_at DESC").fetchall(); con.close()
    records = []
    for x in saved:
        row = json.loads(x["payload"]); row.update(status=x["status"], approval=x["approval"], updated_at=x["updated_at"])
        ensure_workflow_fields(row, x["status"], x["approval"])
        add_security_fields(row)
        row["checks"] = check(row); row["passed"] = sum(v["passed"] for v in row["checks"])
        row["risk"] = assess_risk(row["checks"], row)
        failed_codes = {item["code"] for item in row["checks"] if not item["passed"]}
        row["risk_decisions"] = {code: value for code, value in row.get("risk_decisions", {}).items() if code in failed_codes}
        row["bank_account_masked"] = masked_bank(row.get("bank_account"))
        amount = float(row.get("amount", 0) or 0)
        paid_amount = float(row.get("paid_amount", 0) or 0)
        # The compliance ledger is an archive of completely settled business,
        # so a partial or zero-value payment must never be included.
        row["ledger_eligible"] = (
            row["passed"] == len(RULES)
            and row.get("status") == "已付款"
            and bool(row.get("verified"))
            and amount > 0
            and paid_amount >= amount - 0.01
        )
        row["current_stage"] = current_stage(row)
        records.append(row)
    return records

def current_stage(row):
    if row.get("status") == "已付款": return "已付款"
    if row.get("status") == "审批驳回": return "审批驳回"
    if row.get("passed", 0) < len(RULES): return "风险待驳回"
    if not row.get("verified"): return "待核验"
    if float(row.get("amount", 0) or 0) >= PAYMENT_THRESHOLD and not row.get("financial_approved"): return "待审批"
    return "待付款"

def get_business(business_id):
    return next((row for row in get_rows() if row["id"] == business_id), None)

def save_business(row, status=None, approval=None):
    ensure_workflow_fields(row, status or row.get("status"), approval)
    add_security_fields(row)
    now = datetime.now().isoformat(timespec="seconds")
    con = db()
    con.execute("UPDATE businesses SET payload=?,status=?,approval=?,updated_at=? WHERE id=?", (json.dumps(row, ensure_ascii=False), row["status"], row["approval"], now, row["id"]))
    con.commit(); con.close()
    row["updated_at"] = now
    return row

def clean_number(value, default=0):
    if value is None or value == "": return default
    try: return float(str(value).replace(",", "").replace("%", ""))
    except (TypeError, ValueError): return default

def clean_date(value, fallback=None):
    if isinstance(value, datetime): return value.date().isoformat()
    if value:
        return str(value)[:10]
    return datetime.now().date().isoformat() if fallback is None else fallback

def sheet_rows(ws):
    values = list(ws.iter_rows(values_only=True))
    if not values: return []
    headers = [str(x).strip() if x is not None else "" for x in values[0]]
    # 空白单元格归一化为空串：否则 str(None) 会变成字面量 "None"，
    # 既会污染界面显示，也会让"签字留空"这类缺失被误判为已填写。
    return [{key: ("" if value is None else value) for key, value in zip(headers, row)}
            for row in values[1:] if any(v is not None and v != "" for v in row)]

def find_sheet(workbook, name):
    return workbook[name] if name in workbook.sheetnames else None


def ledger_sheet_to_import_rows(ws):
    """Accept the supplied 59-column ledger template as an import source too."""
    values = list(ws.iter_rows(min_row=3, values_only=True))
    if not values:
        return [], [], []
    headers = [str(value).strip() if value is not None else "" for value in values[0]]
    rows = [dict(zip(headers, row)) for row in values[1:] if any(value not in (None, "") for value in row)]
    weigh_rows, inspect_rows, inbound_rows = [], [], []
    for item in rows:
        if not item.get("过磅单号") and not item.get("供货人姓名"):
            continue
        business_id = str(item.get("过磅单号") or item.get("收购工单号") or "").strip()
        if not business_id:
            continue
        supplier_type = str(item.get("供货人类别", "自然人"))
        supplier_type = "公司" if any(word in supplier_type for word in ("公司", "企业")) else "自然人"
        gross = clean_number(item.get("过磅毛重(吨)")) * 1000
        tare = clean_number(item.get("过磅皮重(吨)")) * 1000
        net_value = clean_number(item.get("过磅净重(吨)"), None)
        net = (gross - tare) if net_value is None else net_value * 1000
        inbound_value = clean_number(item.get("入库实收数量(吨)"), None)
        # A blank inbound quantity remains blank/zero and therefore fails R02;
        # it must not be silently inferred as a valid zero-quantity receipt.
        inbound_qty = 0 if inbound_value is None else inbound_value * 1000
        settlement_value = clean_number(item.get("扣杂后结算数量(吨)"), None)
        settlement_qty = None if settlement_value is None else settlement_value * 1000
        deduct = max(net - inbound_qty, 0)
        weigh_rows.append({"过磅单编号": business_id, "货物批次号": str(item.get("收购工单号", business_id)), "过磅时间": item.get("过磅时间"), "供货方名称": item.get("供货人姓名", ""), "供货方身份证号": item.get("身份证号", ""), "供货方类别": supplier_type, "个体户登记标记": "否", "车牌号": "虚拟车牌-导入", "货物品名": item.get("物资名称", ""), "毛重 (Kg)": gross, "皮重 (Kg)": tare, "净重 (Kg)": net, "过磅方式": "模板导入", "过磅员签字": item.get("制单人", "模板导入"), "司机签字": item.get("制单人", "模板导入"), "收款户名": item.get("收款户名", item.get("供货人姓名", "")), "银行卡号": item.get("收款账号", ""), "开户行": item.get("收款开户行", "")})
        inspect_rows.append({"质检单编号": item.get("质检单号", ""), "关联过磅单编号": business_id, "质检时间": item.get("质检时间"), "货物品名": item.get("物资名称", ""), "货品规格": item.get("品类规格", ""), "质检等级": "B级", "质检扣杂数量 (Kg)": deduct, "结算数量 (Kg)": settlement_qty, "含水率 / 杂质率": item.get("质检扣杂率(%)", 0), "计价单价 (元 / Kg)": clean_number(item.get("收购单价(元/吨)")) / 1000, "计价金额 (元)": clean_number(item.get("结算金额(元)")), "质检员签字": item.get("制单人", "模板导入"), "复核人签字": item.get("财务审核人", "模板导入"), "票据合规方式": "自然人自行代开发票"})
        inbound_rows.append({"入库单编号": item.get("入库单号", ""), "关联过磅单编号": business_id, "入库时间": item.get("入库时间"), "入库数量 (Kg)": inbound_qty, "仓管员签字": item.get("制单人", "模板导入"), "付款时间": item.get("付款日期")})
    return weigh_rows, inspect_rows, inbound_rows

def persist_imported_rows(weigh_rows, inspect_rows, inbound_rows=None):
    inspections = {str(x.get("关联过磅单编号", "")).strip(): x for x in inspect_rows}
    inbounds = {str(x.get("关联过磅单编号", "")).strip(): x for x in (inbound_rows or [])}
    con = db(); now = datetime.now().isoformat(timespec="seconds"); imported = 0
    for raw in weigh_rows:
        business_id = str(raw.get("过磅单编号", "")).strip()
        if not business_id: continue
        quality = inspections.get(business_id, {}); inbound = inbounds.get(business_id, {})
        gross = clean_number(raw.get("毛重 (Kg)"))
        tare = clean_number(raw.get("皮重 (Kg)"))
        # Net weight is derived when the workbook leaves the field blank (for
        # example, when a user clears a formula before saving the file).
        net_value = clean_number(raw.get("净重 (Kg)"), None)
        net = gross - tare if net_value is None else net_value
        deduct = clean_number(quality.get("质检扣杂数量 (Kg)"), None)
        quality_qty = clean_number(quality.get("质检数量 (Kg)"), None)
        if deduct is None: deduct = max(net - (quality_qty or 0), 0)
        inbound_value = clean_number(inbound.get("入库数量 (Kg)"), None)
        inbound_qty = 0 if inbound_value is None else inbound_value
        price, amount = clean_number(quality.get("计价单价 (元 / Kg)")), clean_number(quality.get("计价金额 (元)"))
        weigh_date_value = raw.get("业务日期") or raw.get("过磅时间")
        business_date = clean_date(weigh_date_value, fallback="")
        settlement_value = clean_number(quality.get("结算数量 (Kg)"), None)
        if settlement_value is None:
            settlement_value = inbound_qty
        row = {"id":business_id,"batch":str(raw.get("货物批次号", "")),"supplier":str(raw.get("供货方名称", "")),"supplier_type":str(raw.get("供货方类别", "自然人")) or "自然人","id_no":str(raw.get("供货方身份证号", "")),"registered":str(raw.get("个体户登记标记", "否")),"plate":str(raw.get("车牌号", "")),"goods":str(raw.get("货物品名", "")),"gross":gross,"tare":tare,"net":net,"method":str(raw.get("过磅方式", "")),"weigher":str(raw.get("过磅员签字", "")),"driver":str(raw.get("司机签字", "")),"weigh_date":business_date,"weigh_time":fmt_datetime(raw.get("过磅时间")),"inspect_id":str(quality.get("质检单编号", "")),"inspect_date":clean_date(quality.get("质检时间") or quality.get("质检日期"), fallback=""),"inspect_time":fmt_datetime(quality.get("质检时间") or quality.get("质检日期")),"inspect_goods":str(quality.get("货物品名", "")),"inbound_id":str(inbound.get("入库单编号", "")),"inbound_date":clean_date(inbound.get("入库时间"), fallback=""),"inbound_time":fmt_datetime(inbound.get("入库时间")),"inbound_goods":str(inbound.get("货物品名", "")),"payment_time":fmt_datetime(inbound.get("付款时间") or raw.get("付款时间")),"spec":str(quality.get("货品规格", "")),"grade":str(quality.get("质检等级", "")),"inspect_qty":deduct,"inbound_qty":inbound_qty,"settlement_qty":settlement_value,"moisture":clean_number(quality.get("含水率 / 杂质率")),"price":price,"amount":amount or round(settlement_value * price, 2),"inspector":str(quality.get("质检员签字", "")),"reviewer":str(quality.get("复核人签字", quality.get("复核人签字", ""))),"invoice":"reverse" if "反向" in str(quality.get("票据合规方式", "")) else "self" if "代开" in str(quality.get("票据合规方式", "")) else "none","annual":0,"frequency":0,"bank_name":str(raw.get("收款户名", raw.get("供货方名称", ""))),"bank_account":str(raw.get("银行卡号", "")),"bank_branch":str(raw.get("开户行", "")),"status":"待核验","approval":"无需审批","verified":False,"review_cycle":1,"risk_decisions":{},"paid_amount":0}
        add_security_fields(row)
        is_risk = any(not item["passed"] for item in check(row))
        if is_risk:
            status, approval = "高风险", "待逐项风险处置"
        else:
            # Clean imported records are immediately routed by amount: small
            # payments are ready for the cashier, while larger ones wait for
            # the finance lead's approval.
            row["verified"] = True
            status = "待审批" if float(row.get("amount", 0) or 0) >= PAYMENT_THRESHOLD else "待付款"
            approval = "待财务负责人审批" if status == "待审批" else "无需审批"
        row["status"], row["approval"] = status, approval
        ensure_workflow_fields(row, status, approval)
        con.execute("INSERT OR REPLACE INTO businesses VALUES (?,?,?,?,?)", (business_id, json.dumps(row, ensure_ascii=False), status, approval, now)); imported += 1
    con.commit(); con.close(); return imported


def sync_local_templates(auto=False):
    """Import changed workbooks from local_templates; used by the button and dashboard polling."""
    os.makedirs(LOCAL_TEMPLATE_DIR, exist_ok=True)
    con = db()
    saved_state = {row["path"]: str(row["mtime"]) for row in con.execute("SELECT path,mtime FROM local_template_state").fetchall()}
    con.close()
    total = 0; files = 0; changed_files = []
    paths = [os.path.join(LOCAL_TEMPLATE_DIR, name) for name in os.listdir(LOCAL_TEMPLATE_DIR) if name.lower().endswith(".xlsx")]
    for path in paths:
        mtime = f"{os.stat(path).st_mtime_ns}:{os.path.getsize(path)}"
        # Manual sync and dashboard polling use the same fingerprint guard.
        # A button click should pick up edits, not re-import every unchanged
        # workbook and create duplicate import-history entries.
        if saved_state.get(path) == mtime:
            continue
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            weigh_rows = sheet_rows(wb["过磅单"]) if "过磅单" in wb.sheetnames else []
            inspect_rows = sheet_rows(wb["质检单"]) if "质检单" in wb.sheetnames else []
            inbound_rows = sheet_rows(wb["入库单"]) if "入库单" in wb.sheetnames else []
            if not weigh_rows and not inspect_rows and not inbound_rows and "合规台账" in wb.sheetnames:
                weigh_rows, inspect_rows, inbound_rows = ledger_sheet_to_import_rows(wb["合规台账"])
            total += persist_imported_rows(weigh_rows, inspect_rows, inbound_rows)
            files += 1; changed_files.append(os.path.basename(path))
            saved_state[path] = mtime
        except Exception as exc:
            # Leave a malformed workbook untouched so the user can correct it locally.
            app.logger.warning("本地模板同步失败 %s: %s", path, exc)
            continue
    con = db()
    for path, mtime in saved_state.items():
        con.execute("INSERT OR REPLACE INTO local_template_state(path,mtime) VALUES(?,?)", (path, mtime))
    con.commit(); con.close()
    if total and changed_files:
        con = db()
        con.execute("INSERT INTO import_jobs(filename,rows_count,result,created_at) VALUES(?,?,?,?)", ("、".join(changed_files), total, "本地模板自动同步并执行9项规则", datetime.now().isoformat(timespec="seconds")))
        con.commit(); con.close()
    return {"count": total, "files": files, "changed": changed_files}

def signed_in(): return bool(session.get("user"))

@app.route("/")
def home():
    return redirect(url_for("login")) if not signed_in() else render_template("index.html", user=session["user"])

@app.route("/login", methods=["GET", "POST"])
def login():
    if signed_in(): return redirect(url_for("home"))
    error = None
    if request.method == "POST":
        username=request.form.get("username", "").strip(); password=request.form.get("password", "")
        # 演示账户；部署时应替换为数据库密码哈希或企业 SSO。
        account = ACCOUNTS.get(username)
        if account and password == account[0]:
            session["user"] = {"name": account[1], "role": account[2], "permissions": sorted(ROLE_PERMISSIONS[account[2]])}
            return redirect(url_for("home"))
        error="账号或密码不正确"
    return render_template("login.html", error=error)

@app.get("/logout")
def logout(): session.clear(); return redirect(url_for("login"))

@app.before_request
def guard_api():
    if request.path.startswith("/api/") and not signed_in(): return jsonify({"error":"登录已失效"}), 401

def require_permission(permission):
    if not role_can(permission): return permission_error(permission)
    return None

@app.get("/api/dashboard")
def dashboard():
    denied = require_permission("overview")
    if denied: return denied
    # A modified workbook in local_templates is picked up on the next dashboard
    # refresh, so users can edit locally without re-uploading the file.
    if role_can("ingest"):
        sync_local_templates(auto=True)
    rows = get_rows(); total = len(rows); risks = [r for r in rows if r["passed"] < 9]
    imports_con = db()
    imports = imports_con.execute("SELECT filename,rows_count,result,created_at FROM import_jobs ORDER BY id DESC LIMIT 5").fetchall()
    imports_con.close()
    return jsonify({"stats":{"total":total,"passed":sum(r["passed"] == len(RULES) for r in rows),"risk":len(risks),"paid":sum(r["status"] == "已付款" for r in rows),"amount":sum(float(r.get("paid_amount",0) or 0) for r in rows)},"records":rows,"rules":[{"code":a,"name":b,"enabled":True} for a,b in RULES],"imports":[dict(x) for x in imports],"permissions":session["user"].get("permissions",[])})

@app.post("/api/business/manual")
def create_manual_business():
    denied = require_permission("ingest")
    if denied: return denied
    data = request.get_json(silent=True) or {}
    required = ["supplier", "plate", "goods", "gross", "tare", "method", "weigher", "driver"]
    missing = [field for field in required if not str(data.get(field, "")).strip()]
    if missing:
        return jsonify({"error": "请填写供货方、车牌号、货物品名、重量、过磅方式和签字信息"}), 400
    gross, tare = clean_number(data.get("gross")), clean_number(data.get("tare"))
    if gross <= tare or tare < 0:
        return jsonify({"error": "毛重必须大于皮重，皮重不能小于 0"}), 400
    business_id, batch_id = generate_document_ids()
    now_date = clean_date(data.get("weigh_date"))
    row = {
        "id": business_id,
        "batch": batch_id,
        "supplier": str(data.get("supplier", "")).strip(),
        "id_no": str(data.get("id_no", "")).strip(),
        "registered": str(data.get("registered", "否")).strip() or "否",
        "plate": str(data.get("plate", "")).strip(),
        "goods": str(data.get("goods", "")).strip(),
        "gross": gross,
        "tare": tare,
        "net": round(gross - tare, 2),
        "method": str(data.get("method", "")).strip(),
        "weigher": str(data.get("weigher", "")).strip(),
        "driver": str(data.get("driver", "")).strip(),
        "weigh_date": now_date,
        "weigh_time": fmt_datetime(data.get("weigh_time")),
        "inspect_id": "",
        "inspect_date": now_date,
        "inspect_time": fmt_datetime(data.get("inspect_time")),
        "inbound_id": "",
        "inbound_date": now_date,
        "inbound_time": fmt_datetime(data.get("inbound_time")),
        "payment_time": fmt_datetime(data.get("payment_time")),
        "spec": "",
        "grade": "",
        "inspect_qty": 0,
        "inbound_qty": 0,
        "moisture": 0,
        "price": 0,
        "amount": 0,
        "inspector": "",
        "reviewer": "",
        "bank_name": str(data.get("bank_name", "")).strip(),
        "bank_account": str(data.get("bank_account", "")).strip(),
        "bank_branch": str(data.get("bank_branch", "")).strip(),
        "invoice": "none",
        "annual": 0,
        "frequency": 0,
        "status": "待核验",
        "approval": "无需审批",
        "verified": False,
        "review_cycle": 1,
        "risk_decisions": {},
    }
    ensure_workflow_fields(row)
    add_security_fields(row)
    con = db()
    con.execute(
        "INSERT INTO businesses VALUES (?,?,?,?,?)",
        (business_id, json.dumps(row, ensure_ascii=False), row["status"], row["approval"], datetime.now().isoformat(timespec="seconds")),
    )
    con.commit()
    con.close()
    return jsonify({"ok": True, "message": "手动单据已保存，编号和批次号已自动生成并完成 SHA-256 指纹留痕。", "record": row})

@app.post("/api/business/<business_id>/approve")
def approve(business_id):
    denied = require_permission("approval")
    if denied: return denied
    business = get_business(business_id)
    if not business: return jsonify({"error":"业务不存在"}), 404
    if business["status"] == "已付款": return jsonify({"error":"已付款业务不能再次审批"}), 400
    if float(business.get("amount", 0) or 0) < PAYMENT_THRESHOLD:
        return jsonify({"error":"金额低于5万元，无需财务负责人审批"}), 400
    if business["passed"] < 9:
        return jsonify({"error":"风险业务必须逐条驳回并完成补件，修正后再次核验，不能用例外审批跳过"}), 400
    if not business["verified"]:
        return jsonify({"error":"请先在核验中心确认‘已核验’"}), 400
    business["financial_approved"] = True
    save_business(business, "待付款" if business["verified"] else business["status"], "财务负责人已审批")
    return jsonify({"ok":True, "message":"财务负责人审批通过，业务可进入付款中心。"})

@app.post("/api/business/<business_id>/reject")
def reject(business_id):
    data = request.get_json(silent=True) or {}
    code = data.get("code")
    if code: return reject_risk(business_id, code)
    # Risk rejection belongs to the review workflow.  Finance can reject from
    # payment approval, while accounting/cashier can reject from verification.
    denied = require_any_permission("approval", "verify")
    if denied: return denied
    business = get_business(business_id)
    if not business: return jsonify({"error":"业务不存在"}), 404
    if business["status"] == "已付款": return jsonify({"error":"已付款业务不能再次驳回"}), 400
    failed = [item for item in business["checks"] if not item["passed"]]
    if not failed: return jsonify({"error":"当前没有风险项可驳回"}), 400
    opinion = data.get("opinion", "风险业务一键驳回，退回补件")
    for item in failed:
        business.setdefault("risk_decisions", {})[item["code"]] = {"decision":"rejected", "opinion": opinion, "updated_at": datetime.now().isoformat(timespec="seconds")}
    business.setdefault("rejection_history", []).append(rejection_entry(business, "ALL", opinion))
    business["verified"] = False
    save_business(business, "审批驳回", "一键驳回，待补件")
    return jsonify({"ok": True, "message": "风险业务已一键驳回，等待补件修改"})

@app.post("/api/business/<business_id>/risk/<code>/reject")
def reject_risk(business_id, code):
    denied = require_any_permission("approval", "verify")
    if denied: return denied
    business = get_business(business_id)
    if not business: return jsonify({"error":"业务不存在"}), 404
    if business["status"] == "已付款": return jsonify({"error":"已付款业务不能再次处置风险"}), 400
    check_item = next((item for item in business["checks"] if item["code"] == code), None)
    if not check_item: return jsonify({"error":"风险规则不存在"}), 404
    if check_item["passed"]: return jsonify({"error":"该规则当前已通过，无需驳回"}), 400
    opinion = (request.get_json(silent=True) or {}).get("opinion", "未通过风险项，退回补件")
    business.setdefault("risk_decisions", {})[code] = {"decision":"rejected", "opinion": opinion, "updated_at":datetime.now().isoformat(timespec="seconds")}
    business.setdefault("rejection_history", []).append(rejection_entry(business, code, opinion))
    business["verified"] = False
    failed = [item for item in business["checks"] if not item["passed"]]
    handled = sum(1 for item in failed if business["risk_decisions"].get(item["code"], {}).get("decision") == "rejected")
    status = "审批驳回" if handled == len(failed) else "高风险"
    approval = "管理层已逐项驳回，待补件" if status == "审批驳回" else f"待逐项风险处置 ({handled}/{len(failed)})"
    save_business(business, status, approval)
    return jsonify({"ok":True, "message": "该风险项已驳回，全部风险项已退回补件。" if status == "审批驳回" else f"已驳回 {code}，还需处理 {len(failed)-handled} 项风险。"})

@app.post("/api/business/<business_id>/resubmit")
def resubmit(business_id):
    # Financial users may view/reject/approve only; only the backend administrator
    # can edit source documents and submit a revised version.
    denied = require_permission("admin")
    if denied: return denied
    business = get_business(business_id)
    if not business: return jsonify({"error":"业务不存在"}), 404
    if business["status"] != "审批驳回": return jsonify({"error":"请先逐项驳回全部风险依据，再提交补件"}), 400
    data = request.get_json(silent=True) or {}
    editable = ("batch", "id_no", "registered", "plate", "goods", "inspect_goods", "inbound_goods", "gross", "tare", "method", "weigher", "driver", "weigh_date", "weigh_time", "inspect_id", "inspect_date", "inspect_time", "inbound_id", "inbound_date", "inbound_time", "payment_time", "spec", "grade", "inspect_qty", "inbound_qty", "settlement_qty", "moisture", "price", "amount", "inspector", "reviewer", "invoice", "annual", "frequency", "bank_name", "bank_account", "bank_branch")
    for key in editable:
        if key in data:
            business[key] = data[key]
    for key in ("gross", "tare", "net", "inspect_qty", "inbound_qty", "settlement_qty", "moisture", "price", "amount", "annual", "frequency"):
        if key in business: business[key] = clean_number(business[key])
    business["net"] = round(business["gross"] - business["tare"], 2)
    if business.get("rejection_history"):
        revised = source_snapshot(business)
        business["rejection_history"][-1]["revised"] = revised
        business["rejection_history"][-1]["revised_documents"] = revised
    business["review_cycle"] = int(business.get("review_cycle", 1)) + 1
    business["verified"] = False; business["risk_decisions"] = {}; business["financial_approved"] = False
    checks = check(business); passed = sum(item["passed"] for item in checks)
    if passed == 9:
        business["verified"] = True
        status = "待审批" if float(business.get("amount", 0) or 0) >= PAYMENT_THRESHOLD else "待付款"
        approval = "待财务负责人审批" if status == "待审批" else "无需审批"
    else:
        status, approval = "高风险", "待逐项风险处置"
    save_business(business, status, approval)
    business["checks"], business["passed"] = checks, passed
    return jsonify({"ok":True, "message": "补件已提交，系统已重新执行 9 项规则，并按金额自动进入下一环节。", "record":business})

@app.post("/api/business/<business_id>/verify")
def verify(business_id):
    denied = require_permission("verify")
    if denied: return denied
    business = get_business(business_id)
    if not business: return jsonify({"error":"业务不存在"}), 404
    if business["status"] == "已付款": return jsonify({"error":"已付款业务已经核验"}), 400
    if business["passed"] < 9: return jsonify({"error":"仍有风险项未通过，不能标记为已核验"}), 400
    business["verified"] = True
    status = "待审批" if float(business.get("amount", 0) or 0) >= PAYMENT_THRESHOLD else "待付款"
    approval = "待财务负责人审批" if status == "待审批" else "无需审批"
    save_business(business, status, approval)
    return jsonify({"ok":True, "message":"已核验，业务现在可以进入付款中心。"})

@app.post("/api/business/<business_id>/pay")
def pay(business_id):
    denied = require_permission("payment")
    if denied: return denied
    rows = {r["id"]:r for r in get_rows()}; row=rows.get(business_id)
    if not row: return jsonify({"error":"业务不存在"}),404
    if row["passed"] < 9: return jsonify({"error":"风险业务必须补件修正并通过全部规则后才可付款"}),400
    if not row["verified"]: return jsonify({"error":"请先完成‘已核验’操作，核验后才能付款"}),400
    if float(row.get("amount", 0) or 0) >= PAYMENT_THRESHOLD and not row.get("financial_approved"): return jsonify({"error":"金额达到5万元，需财务负责人审批"}),400
    row["paid_amount"] = float(row.get("amount", 0) or 0)
    # Keep the event chain valid even when a demo record is dated in the
    # future relative to the local clock.
    payment_dt = datetime.now()
    inbound_dt = parse_datetime(row.get("inbound_time"))
    if inbound_dt != datetime.min and payment_dt <= inbound_dt:
        payment_dt = inbound_dt + timedelta(minutes=1)
    row["payment_time"] = payment_dt.strftime("%Y-%m-%d %H:%M")
    save_business(row, "已付款", "财务负责人已审批" if float(row.get("amount",0) or 0) >= PAYMENT_THRESHOLD else "无需审批")
    return jsonify({"ok":True,"message":"财务付款确认成功，已同步至全局合规台账。"})

@app.post("/api/import")
def upload():
    denied = require_permission("ingest")
    if denied: return denied
    files=request.files.getlist("files")
    if not files: return jsonify({"error":"请选择 Excel 文件"}),400
    weigh_rows, inspect_rows, inbound_rows, filenames = [], [], [], []
    for f in files:
        try:
            wb=load_workbook(f,read_only=True,data_only=True); filenames.append(f.filename)
            weigh = find_sheet(wb, "过磅单"); inspect = find_sheet(wb, "质检单"); inbound = find_sheet(wb, "入库单")
            if weigh: weigh_rows += sheet_rows(weigh)
            if inspect: inspect_rows += sheet_rows(inspect)
            if inbound: inbound_rows += sheet_rows(inbound)
            if not weigh and not inspect and not inbound and "合规台账" in wb.sheetnames:
                lw, li, lb = ledger_sheet_to_import_rows(wb["合规台账"])
                weigh_rows += lw; inspect_rows += li; inbound_rows += lb
            if not weigh and not inspect and not inbound:
                first = sheet_rows(wb.active)
                headers = set(first[0].keys()) if first else set()
                if "过磅单编号" in headers: weigh_rows += first
                elif "关联过磅单编号" in headers and "入库数量 (Kg)" in headers: inbound_rows += first
                elif "关联过磅单编号" in headers: inspect_rows += first
        except Exception: return jsonify({"error":f"无法读取 {f.filename}，请使用标准 .xlsx 模板"}),400
    if not weigh_rows: return jsonify({"error":"未识别到过磅单数据。请使用下载的标准模板。"}),400
    count = persist_imported_rows(weigh_rows, inspect_rows, inbound_rows)
    con=db(); con.execute("INSERT INTO import_jobs(filename,rows_count,result,created_at) VALUES(?,?,?,?)", ("、".join(filenames),count,"已入库并执行9项规则",datetime.now().isoformat(timespec="seconds")));con.commit();con.close()
    return jsonify({"ok":True,"count":count,"weigh":len(weigh_rows),"inspect":len(inspect_rows),"inbound":len(inbound_rows),"message":f"成功写入 {count} 笔业务，已自动匹配质检单、入库单并完成合规核验。"})

@app.delete("/api/import-history")
def clear_import_history():
    denied = require_permission("ingest")
    if denied: return denied
    con = db(); con.execute("DELETE FROM import_jobs"); con.commit(); con.close()
    return jsonify({"ok": True, "message": "最近导入记录已清除"})

@app.post("/api/local-sync")
def local_sync():
    denied = require_permission("ingest")
    if denied: return denied
    # File fingerprints make an explicit sync idempotent too: only a workbook
    # saved locally since the previous sync is imported again.
    result = sync_local_templates(auto=True)
    return jsonify({"ok": True, "count": result["count"], "files": result["files"], "message": f"已同步本地模板 {result['files']} 个，更新 {result['count']} 笔业务"})

@app.get("/api/templates/standard")
def standard_template():
    denied = require_permission("ingest")
    if denied: return denied
    wb = Workbook(); weigh = wb.active; weigh.title = "过磅单"; inspect = wb.create_sheet("质检单")
    inbound = wb.create_sheet("入库单")
    weigh.append(WEIGH_HEADERS); inspect.append(INSPECT_HEADERS)
    inbound.append(INBOUND_HEADERS)
    weigh.append(["GB20260905001", "PC20260905001", "2026-09-05 09:00", "虚拟赵建国示例", "SIM199001010001", "自然人", "否", "虚拟车牌-001", "废铁", 12000, 4000, 8000, "二次过磅", "虚拟过磅员甲", "虚拟司机甲", "虚拟赵建国示例", "621700000000009999", "中国工商银行虚拟支行"])
    inspect.append(["ZJ20260905001", "GB20260905001", "2026-09-05 09:15", "废铁", "重型废钢", "B级", 100, "1.2%", 2.65, 20935, "虚拟质检员甲", "虚拟复核员甲", "自然人自行代开发票"])
    inbound.append(["RK20260905001", "GB20260905001", "2026-09-05 09:30", 7900, "虚拟仓管员甲", "2026-09-06 10:00"])
    for ws in (weigh, inspect, inbound):
        ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]: cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill("solid", fgColor="167A59")
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 18
    stream=io.BytesIO(); wb.save(stream); stream.seek(0)
    local_path = os.path.join(LOCAL_TEMPLATE_DIR, "绿循智审_本地同步模板.xlsx")
    with open(local_path, "wb") as local_file:
        local_file.write(stream.getvalue())
    template_fingerprint = f"{os.stat(local_path).st_mtime_ns}:{os.path.getsize(local_path)}"
    con = db(); con.execute("INSERT OR REPLACE INTO local_template_state(path,mtime) VALUES(?,?)", (local_path, template_fingerprint)); con.commit(); con.close()
    stream.seek(0)
    return send_file(stream,as_attachment=True,download_name="绿循智审_标准导入模板.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/api/templates/ledger")
def ledger_template():
    denied = require_permission("ingest")
    if denied: return denied
    if not os.path.exists(LEDGER_TEMPLATE_PATH):
        return jsonify({"error": "合规台账模板不存在"}), 404
    wb = load_workbook(LEDGER_TEMPLATE_PATH)
    ws = wb["合规台账"]
    if ws.max_row >= 4:
        ws.delete_rows(4, ws.max_row - 3)
    stream = io.BytesIO(); wb.save(stream)
    # Keep a server-local working copy so edits to the downloaded template can
    # be picked up by the same polling path as the standard three-sheet file.
    os.makedirs(LOCAL_TEMPLATE_DIR, exist_ok=True)
    local_path = os.path.join(LOCAL_TEMPLATE_DIR, "再生资源公司收购合规台账模板.xlsx")
    with open(local_path, "wb") as local_file:
        local_file.write(stream.getvalue())
    template_fingerprint = f"{os.stat(local_path).st_mtime_ns}:{os.path.getsize(local_path)}"
    con = db(); con.execute("INSERT OR REPLACE INTO local_template_state(path,mtime) VALUES(?,?)", (local_path, template_fingerprint)); con.commit(); con.close()
    stream.seek(0)
    return send_file(stream, as_attachment=True, download_name="再生资源公司收购合规台账模板.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/api/admin")
def admin_data():
    denied = require_permission("admin")
    if denied: return denied
    return jsonify({"storage":"SQLite · greenloop.db","backend":"Flask · Python","rules":[{"code":a,"name":b,"owner":"合规风控部","enabled":True,"updated":"2026-09-08"} for a,b in RULES],"accounts":[{"account":k,"name":v[1],"role":v[2],"permissions":sorted(ROLE_PERMISSIONS[v[2]]),"status":"启用"} for k,v in ACCOUNTS.items()]})

@app.post("/api/admin/reset-demo")
def reset_demo():
    denied = require_permission("admin")
    if denied: return denied
    reset_demo_data()
    return jsonify({"ok": True, "message": "演示数据已恢复：100 条全虚拟业务（自然人 90%、公司 10%），含风险、审批和已付款示例。"})

@app.get("/api/export")
def export():
    denied = require_permission("ledger")
    if denied: return denied
    if os.path.exists(LEDGER_TEMPLATE_PATH):
        wb = load_workbook(LEDGER_TEMPLATE_PATH)
        ws = wb["合规台账"]
        if ws.max_row >= 4:
            ws.delete_rows(4, ws.max_row - 3)
    else:
        wb = Workbook(); ws = wb.active; ws.title = "合规台账"
        ws.append(["台账编号", "业务日期", "业务类型"])
    rows = get_rows()
    sort = request.args.get("sort", "risk_desc")
    if sort == "risk_desc": rows.sort(key=lambda r: (r["passed"] == 9, r["passed"], r["updated_at"]))
    elif sort == "risk_asc": rows.sort(key=lambda r: (r["passed"] != 9, -r["passed"], r["updated_at"]))
    elif sort == "amount_desc": rows.sort(key=lambda r: r["amount"], reverse=True)
    elif sort == "updated_desc": rows.sort(key=lambda r: r["updated_at"], reverse=True)
    for r in rows:
        if not r.get("ledger_eligible"): continue
        paid = float(r.get("paid_amount", 0) or 0)
        amount = float(r.get("amount", 0) or 0)
        kg = 1000.0
        inbound_qty = float(r.get("inbound_qty", 0) or 0)
        settlement_qty = float(r.get("settlement_qty", 0) or 0) or inbound_qty
        settlement_tonnes = settlement_qty / kg
        inbound_tonnes = inbound_qty / kg
        quantity_difference = inbound_tonnes - settlement_tonnes
        difference_rate = quantity_difference / settlement_tonnes * 100 if settlement_tonnes else 0
        own_receiving = "是" if r.get("bank_name") == r.get("supplier") else "否"
        attachment_count = 3 + len(r.get("rejection_history", []))
        ws.append([
            f"HG{r['id'][2:]}", r.get("weigh_date", ""), "自然人收购" if r.get("supplier_type") != "公司" else "企业采购", f"SG{r['id'][2:]}", r.get("supplier", ""), r.get("id_no", ""), "", r.get("supplier_type", "自然人"), r["id"], r.get("weigh_time", ""), r.get("inspect_id", ""), r.get("inspect_time", ""), r.get("inbound_id", ""), r.get("inbound_time", ""), f"JS{r['id'][2:]}", f"FX{r['id'][2:]}", r.get("goods", ""), r.get("spec", ""), "吨", float(r.get("gross", 0) or 0) / kg, float(r.get("tare", 0) or 0) / kg, float(r.get("net", 0) or 0) / kg, r.get("moisture", 0), settlement_tonnes, inbound_tonnes, quantity_difference, difference_rate, float(r.get("price", 0) or 0) * kg, amount, 0, paid, "", r.get("payment_time", ""), "银行转账", "虚拟结算账户", r.get("bank_name", ""), r.get("bank_account", ""), r.get("bank_branch", ""), f"FH{r['id'][2:]}", own_receiving, "通过", "通过", "通过", "通过", r.get("annual", 0), "通过", "通过", "合规", "无", "", "系统", "财务负责人" if amount >= PAYMENT_THRESHOLD else "无需审批", r.get("updated_at", ""), "财务负责人" if amount >= PAYMENT_THRESHOLD else "无需审批", "已归档", r.get("updated_at", ""), f"ARCHIVE/{r['id']}", attachment_count, "规则全通过且已全额付款",
        ])
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A4"
    for column_index in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(column_index)].width = max(ws.column_dimensions[get_column_letter(column_index)].width or 0, 15)
    stream=io.BytesIO(); wb.save(stream); stream.seek(0)
    return send_file(stream,as_attachment=True,download_name="绿循智审_合规台账.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    init_db(); app.run(debug=False, host="0.0.0.0", port=int(os.environ.get("PORT",5000)))
else: init_db()
