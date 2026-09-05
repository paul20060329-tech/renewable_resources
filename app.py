import base64
import hashlib
import io
import json
import os
import sqlite3
from datetime import datetime

from flask import Flask, jsonify, render_template, request, send_file, session, redirect, url_for
from openpyxl import Workbook, load_workbook

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "greenloop-demo-change-before-production")
DB_PATH = os.environ.get("DATABASE_PATH", "greenloop.db")

RULES = [
    ("R01", "单据关联一致性"), ("R02", "重量偏差校验"), ("R03", "自然人身份完整性"),
    ("R04", "单据日期逻辑"), ("R05", "等级计价匹配"), ("R06", "短期高频交易"),
    ("R07", "原始单据完整性"), ("R08", "票据合规二选一"), ("R09", "主体资格预警"),
]

WEIGH_HEADERS = ["过磅单编号", "货物批次号", "业务日期", "供货方名称", "供货方身份证号", "个体户登记标记", "车牌号", "货物品名", "毛重 (Kg)", "皮重 (Kg)", "净重 (Kg)", "过磅方式", "过磅员签字", "司机签字"]
INSPECT_HEADERS = ["质检单编号", "关联过磅单编号", "质检日期", "货物品名", "货品规格", "质检等级", "质检数量 (Kg)", "含水率 / 杂质率", "计价单价 (元 / Kg)", "计价金额 (元)", "质检员签字", "复核人签字", "票据合规方式"]

DEMO_ROWS = [
    {"id":"GB20260825246","batch":"PC20260825246","supplier":"张建国","id_no":"510104197803152417","registered":"否","plate":"川A3F72P","goods":"废不锈钢","gross":12560,"tare":5820,"net":6740,"method":"一次过磅","weigher":"李军","driver":"王强","weigh_date":"2026-08-25","inspect_id":"ZJ20260825246","inspect_date":"2026-08-26","spec":"一级统料","grade":"A级","inspect_qty":6705,"moisture":0.52,"price":9.8,"amount":65709,"inspector":"刘梅","reviewer":"赵晓","invoice":"reverse","annual":382000,"frequency":2,"status":"已付款","approval":"无需审批"},
    {"id":"GB20260829182","batch":"PC20260829183","supplier":"陈丽","id_no":"","registered":"是","plate":"川A8K91L","goods":"废杂铝合金","gross":9850,"tare":4650,"net":5200,"method":"二次过磅","weigher":"","driver":"陈丽","weigh_date":"2026-08-29","inspect_id":"ZJ20260829182","inspect_date":"2026-08-29","spec":"破碎杂铝","grade":"C级","inspect_qty":4300,"moisture":8.7,"price":13.2,"amount":68640,"inspector":"孙杰","reviewer":"","invoice":"none","annual":4960000,"frequency":8,"status":"高风险","approval":"待逐项风险处置"},
    {"id":"GB20260903001","batch":"PC20260903001","supplier":"周伟","id_no":"510105198611042010","registered":"否","plate":"川A6M28Q","goods":"废铁","gross":18320,"tare":7320,"net":11000,"method":"二次过磅","weigher":"王磊","driver":"周伟","weigh_date":"2026-09-03","inspect_id":"ZJ20260903001","inspect_date":"2026-09-03","spec":"重型废钢","grade":"B级","inspect_qty":10870,"moisture":1.4,"price":2.65,"amount":28805.5,"inspector":"陈敏","reviewer":"何露","invoice":"self","annual":684000,"frequency":3,"status":"待核验","approval":"无需审批"},
    {"id":"GB20260904007","batch":"PC20260904007","supplier":"刘芳","id_no":"510107197906184328","registered":"否","plate":"川A9D53K","goods":"废铜","gross":4680,"tare":1980,"net":2700,"method":"一次过磅","weigher":"李军","driver":"刘芳","weigh_date":"2026-09-04","inspect_id":"ZJ20260904007","inspect_date":"2026-09-03","spec":"紫杂铜","grade":"A级","inspect_qty":2700,"moisture":0.2,"price":61.0,"amount":164700,"inspector":"刘梅","reviewer":"赵晓","invoice":"reverse","annual":1180000,"frequency":4,"status":"高风险","approval":"待逐项风险处置"},
]

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
    current_status = status or row.get("status", "待核验")
    row["verified"] = bool(row.get("verified", current_status in ("已核验", "已付款")))
    row["verification_status"] = "已核验" if row["verified"] else "待核验"
    row.setdefault("review_cycle", 1)
    row.setdefault("risk_decisions", {})
    if approval is not None:
        row["approval"] = approval
    row["status"] = current_status
    row.setdefault("approval", "无需审批")
    return row

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
    now = datetime.now().isoformat(timespec="seconds")
    for item in DEMO_ROWS:
        ensure_workflow_fields(item)
        add_security_fields(item)
        con.execute("INSERT INTO businesses VALUES (?,?,?,?,?)", (item["id"], json.dumps(item, ensure_ascii=False), item["status"], item["approval"], now))
    con.commit()
    con.close()

def check(row):
    net_ok = row["gross"] > row["tare"] >= 0 and abs(row["net"] - (row["gross"] - row["tare"])) < 1
    relation = row["batch"][-3:] == row["id"][-3:] and bool(row["inspect_id"])
    weight_diff = abs(row["net"] - row["inspect_qty"]) / row["net"] if row["net"] else 1
    amount_ok = abs(row["amount"] - row["inspect_qty"] * row["price"]) < 1 and row["grade"] in ["A级", "B级", "C级"]
    items = [
        ("R01", relation and row["goods"], "批次号、关联过磅单与货物品名已交叉匹配" if relation else "批次号或关联单据不一致"),
        ("R02", net_ok and weight_diff <= .05, f"过磅净重 {row['net']:,} Kg，质检数量 {row['inspect_qty']:,} Kg，偏差 {weight_diff:.1%}"),
        ("R03", bool(row["id_no"]) and len(row["id_no"]) >= 15, "自然人身份信息完整" if row["id_no"] else "缺少自然人身份证明"),
        ("R04", row["inspect_date"] >= row["weigh_date"], "质检日期不早于过磅日期" if row["inspect_date"] >= row["weigh_date"] else "质检日期早于过磅日期，存在倒签风险"),
        ("R05", amount_ok, f"{row['inspect_qty']:,} Kg × ¥{row['price']:.2f} = ¥{row['amount']:,.2f}" if amount_ok else "计价金额与数量、单价或等级不匹配"),
        ("R06", row["frequency"] <= 6, f"近30日交易 {row['frequency']} 笔" if row["frequency"] <= 6 else f"近30日交易 {row['frequency']} 笔，超过预警阈值"),
        ("R07", bool(row["weigher"] and row["inspector"] and row["reviewer"]), "过磅、质检、复核留痕完整" if row["weigher"] and row["inspector"] and row["reviewer"] else "原始单据签字或复核信息缺失"),
        ("R08", row["invoice"] in ["reverse", "self"], "已完成反向开票授权" if row["invoice"] == "reverse" else "已留存自然人代开发票" if row["invoice"] == "self" else "未满足反向开票或自然人代开任一条件"),
        ("R09", not (row["registered"] == "是" and row["invoice"] == "reverse") and row["annual"] < 5000000, f"年度累计 ¥{row['annual']:,.0f}，主体资格正常" if row["annual"] < 4500000 else f"年度累计 ¥{row['annual']:,.0f}，临近500万阈值"),
    ]
    return [{"code": code, "name": dict(RULES)[code], "passed": bool(passed), "reason": reason} for code, passed, reason in items]

def get_rows():
    con = db(); saved = con.execute("SELECT payload,status,approval,updated_at FROM businesses ORDER BY updated_at DESC").fetchall(); con.close()
    records = []
    for x in saved:
        row = json.loads(x["payload"]); row.update(status=x["status"], approval=x["approval"], updated_at=x["updated_at"])
        ensure_workflow_fields(row, x["status"], x["approval"])
        add_security_fields(row)
        row["checks"] = check(row); row["passed"] = sum(v["passed"] for v in row["checks"])
        failed_codes = {item["code"] for item in row["checks"] if not item["passed"]}
        row["risk_decisions"] = {code: value for code, value in row.get("risk_decisions", {}).items() if code in failed_codes}
        records.append(row)
    return records

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

def clean_date(value):
    if isinstance(value, datetime): return value.date().isoformat()
    return str(value)[:10] if value else datetime.now().date().isoformat()

def sheet_rows(ws):
    values = list(ws.iter_rows(values_only=True))
    if not values: return []
    headers = [str(x).strip() if x is not None else "" for x in values[0]]
    return [dict(zip(headers, row)) for row in values[1:] if any(v is not None and v != "" for v in row)]

def find_sheet(workbook, name):
    return workbook[name] if name in workbook.sheetnames else None

def persist_imported_rows(weigh_rows, inspect_rows):
    inspections = {str(x.get("关联过磅单编号", "")).strip(): x for x in inspect_rows}
    con = db(); now = datetime.now().isoformat(timespec="seconds"); imported = 0
    for raw in weigh_rows:
        business_id = str(raw.get("过磅单编号", "")).strip()
        if not business_id: continue
        quality = inspections.get(business_id, {})
        gross, tare, net = clean_number(raw.get("毛重 (Kg)")), clean_number(raw.get("皮重 (Kg)")), clean_number(raw.get("净重 (Kg)"))
        qty, price, amount = clean_number(quality.get("质检数量 (Kg)")), clean_number(quality.get("计价单价 (元 / Kg)")), clean_number(quality.get("计价金额 (元)"))
        row = {"id":business_id,"batch":str(raw.get("货物批次号", "")),"supplier":str(raw.get("供货方名称", "")),"id_no":str(raw.get("供货方身份证号", "")),"registered":str(raw.get("个体户登记标记", "否")),"plate":str(raw.get("车牌号", "")),"goods":str(raw.get("货物品名", "")),"gross":gross,"tare":tare,"net":net,"method":str(raw.get("过磅方式", "")),"weigher":str(raw.get("过磅员签字", "")),"driver":str(raw.get("司机签字", "")),"weigh_date":clean_date(raw.get("业务日期")),"inspect_id":str(quality.get("质检单编号", "")),"inspect_date":clean_date(quality.get("质检日期")),"spec":str(quality.get("货品规格", "")),"grade":str(quality.get("质检等级", "")),"inspect_qty":qty,"moisture":clean_number(quality.get("含水率 / 杂质率")),"price":price,"amount":amount,"inspector":str(quality.get("质检员签字", "")),"reviewer":str(quality.get("复核人签字", "")),"invoice":"reverse" if "反向" in str(quality.get("票据合规方式", "")) else "self" if "代开" in str(quality.get("票据合规方式", "")) else "none","annual":0,"frequency":0,"status":"待核验","approval":"无需审批","verified":False,"review_cycle":1,"risk_decisions":{}}
        add_security_fields(row)
        is_risk = any(not item["passed"] for item in check(row))
        status = "高风险" if is_risk else "待核验"
        approval = "待逐项风险处置" if is_risk else "无需审批"
        row["status"], row["approval"] = status, approval
        ensure_workflow_fields(row, status, approval)
        con.execute("INSERT OR REPLACE INTO businesses VALUES (?,?,?,?,?)", (business_id, json.dumps(row, ensure_ascii=False), status, approval, now)); imported += 1
    con.commit(); con.close(); return imported

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
        accounts = {
            "001": ("123456", "赛事演示管理员", "合规管理"),
            "admin": ("admin123", "系统管理员", "合规管理"),
            "finance": ("finance123", "财务审核员", "财务审核"),
        }
        account = accounts.get(username)
        if account and password == account[0]:
            session["user"] = {"name": account[1], "role": account[2]}
            return redirect(url_for("home"))
        error="账号或密码不正确"
    return render_template("login.html", error=error)

@app.get("/logout")
def logout(): session.clear(); return redirect(url_for("login"))

@app.before_request
def guard_api():
    if request.path.startswith("/api/") and not signed_in(): return jsonify({"error":"登录已失效"}), 401

@app.get("/api/dashboard")
def dashboard():
    rows = get_rows(); total = len(rows); risks = [r for r in rows if r["passed"] < 9]
    imports = db().execute("SELECT filename,rows_count,result,created_at FROM import_jobs ORDER BY id DESC LIMIT 5").fetchall()
    return jsonify({"stats":{"total":total,"passed":sum(r["passed"] == 9 for r in rows),"risk":len(risks),"paid":sum(r["status"] == "已付款" for r in rows),"amount":sum(r["amount"] for r in rows)},"records":rows,"rules":[{"code":a,"name":b,"enabled":True} for a,b in RULES],"imports":[dict(x) for x in imports]})

@app.post("/api/business/manual")
def create_manual_business():
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
        "inspect_id": "",
        "inspect_date": now_date,
        "spec": "",
        "grade": "",
        "inspect_qty": 0,
        "moisture": 0,
        "price": 0,
        "amount": 0,
        "inspector": "",
        "reviewer": "",
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
    business = get_business(business_id)
    if not business: return jsonify({"error":"业务不存在"}), 404
    if business["status"] == "已付款": return jsonify({"error":"已付款业务不能再次审批"}), 400
    if business["passed"] < 9:
        return jsonify({"error":"风险业务必须逐条驳回并完成补件，修正后再次核验，不能用例外审批跳过"}), 400
    if not business["verified"]:
        return jsonify({"error":"请先在核验中心确认‘已核验’"}), 400
    return jsonify({"ok":True, "message":"该业务已完成核验，可进入付款中心。"})

@app.post("/api/business/<business_id>/reject")
def reject(business_id):
    code = (request.get_json(silent=True) or {}).get("code")
    if not code:
        return jsonify({"error":"风险依据必须逐条驳回，请提供规则编号 code"}), 400
    return reject_risk(business_id, code)

@app.post("/api/business/<business_id>/risk/<code>/reject")
def reject_risk(business_id, code):
    business = get_business(business_id)
    if not business: return jsonify({"error":"业务不存在"}), 404
    if business["status"] == "已付款": return jsonify({"error":"已付款业务不能再次处置风险"}), 400
    check_item = next((item for item in business["checks"] if item["code"] == code), None)
    if not check_item: return jsonify({"error":"风险规则不存在"}), 404
    if check_item["passed"]: return jsonify({"error":"该规则当前已通过，无需驳回"}), 400
    business.setdefault("risk_decisions", {})[code] = {"decision":"rejected", "updated_at":datetime.now().isoformat(timespec="seconds")}
    business["verified"] = False
    failed = [item for item in business["checks"] if not item["passed"]]
    handled = sum(1 for item in failed if business["risk_decisions"].get(item["code"], {}).get("decision") == "rejected")
    status = "审批驳回" if handled == len(failed) else "高风险"
    approval = "管理层已逐项驳回，待补件" if status == "审批驳回" else f"待逐项风险处置 ({handled}/{len(failed)})"
    save_business(business, status, approval)
    return jsonify({"ok":True, "message": "该风险项已驳回，全部风险项已退回补件。" if status == "审批驳回" else f"已驳回 {code}，还需处理 {len(failed)-handled} 项风险。"})

@app.post("/api/business/<business_id>/resubmit")
def resubmit(business_id):
    business = get_business(business_id)
    if not business: return jsonify({"error":"业务不存在"}), 404
    if business["status"] != "审批驳回": return jsonify({"error":"请先逐项驳回全部风险依据，再提交补件"}), 400
    data = request.get_json(silent=True) or {}
    editable = ("batch", "id_no", "registered", "plate", "goods", "gross", "tare", "method", "weigher", "driver", "weigh_date", "inspect_id", "inspect_date", "spec", "grade", "inspect_qty", "moisture", "price", "amount", "inspector", "reviewer", "invoice", "annual", "frequency")
    for key in editable:
        if key in data:
            business[key] = data[key]
    for key in ("gross", "tare", "net", "inspect_qty", "moisture", "price", "amount", "annual", "frequency"):
        if key in business: business[key] = clean_number(business[key])
    business["net"] = round(business["gross"] - business["tare"], 2)
    business["review_cycle"] = int(business.get("review_cycle", 1)) + 1
    business["verified"] = False; business["risk_decisions"] = {}
    checks = check(business); passed = sum(item["passed"] for item in checks)
    status = "待核验" if passed == 9 else "高风险"
    approval = "无需审批" if passed == 9 else "待逐项风险处置"
    save_business(business, status, approval)
    business["checks"], business["passed"] = checks, passed
    return jsonify({"ok":True, "message": "补件已提交，系统已重新执行 9 项规则，请再次核验。", "record":business})

@app.post("/api/business/<business_id>/verify")
def verify(business_id):
    business = get_business(business_id)
    if not business: return jsonify({"error":"业务不存在"}), 404
    if business["status"] == "已付款": return jsonify({"error":"已付款业务已经核验"}), 400
    if business["passed"] < 9: return jsonify({"error":"仍有风险项未通过，不能标记为已核验"}), 400
    business["verified"] = True
    save_business(business, "已核验", "无需审批")
    return jsonify({"ok":True, "message":"已核验，业务现在可以进入付款中心。"})

@app.post("/api/business/<business_id>/pay")
def pay(business_id):
    rows = {r["id"]:r for r in get_rows()}; row=rows.get(business_id)
    if not row: return jsonify({"error":"业务不存在"}),404
    if row["passed"] < 9: return jsonify({"error":"风险业务必须补件修正并通过全部规则后才可付款"}),400
    if not row["verified"] or row["status"] != "已核验": return jsonify({"error":"请先完成‘已核验’操作，核验后才能付款"}),400
    con=db(); con.execute("UPDATE businesses SET status=?, updated_at=? WHERE id=?",("已付款",datetime.now().isoformat(timespec="seconds"),business_id));con.commit();con.close()
    return jsonify({"ok":True,"message":"财务付款确认成功，已同步至全局合规台账。"})

@app.get("/api/business/<business_id>/payment-code")
def payment_code(business_id):
    rows = {r["id"]: r for r in get_rows()}
    row = rows.get(business_id)
    if not row:
        return jsonify({"error": "业务不存在"}), 404
    if row["passed"] < 9 or not row["verified"]:
        return jsonify({"error": "该业务尚未满足付款条件"}), 403
    import qrcode
    mode = request.args.get("kind", "trace")
    endpoint = "payment_landing" if mode == "wechat" else "trace_landing"
    payload = url_for(endpoint, business_id=row["id"], _external=True)
    image = qrcode.make(payload)
    stream = io.BytesIO()
    image.save(stream, format="PNG")
    stream.seek(0)
    return send_file(stream, mimetype="image/png", max_age=0)

@app.get("/trace/<business_id>")
def trace_landing(business_id):
    row = get_business(business_id)
    if not row: return render_template("payment.html", missing=True), 404
    return render_template("payment.html", row=row, mode="trace", missing=False)

@app.get("/payment/<business_id>")
def payment_landing(business_id):
    row = get_business(business_id)
    if not row: return render_template("payment.html", missing=True), 404
    return render_template("payment.html", row=row, mode="wechat", missing=False)

@app.post("/api/import")
def upload():
    files=request.files.getlist("files")
    if not files: return jsonify({"error":"请选择 Excel 文件"}),400
    weigh_rows, inspect_rows, filenames = [], [], []
    for f in files:
        try:
            wb=load_workbook(f,read_only=True,data_only=True); filenames.append(f.filename)
            weigh = find_sheet(wb, "过磅单"); inspect = find_sheet(wb, "质检单")
            if weigh: weigh_rows += sheet_rows(weigh)
            if inspect: inspect_rows += sheet_rows(inspect)
            if not weigh and not inspect:
                first = sheet_rows(wb.active)
                headers = set(first[0].keys()) if first else set()
                if "过磅单编号" in headers: weigh_rows += first
                elif "关联过磅单编号" in headers: inspect_rows += first
        except Exception: return jsonify({"error":f"无法读取 {f.filename}，请使用标准 .xlsx 模板"}),400
    if not weigh_rows: return jsonify({"error":"未识别到过磅单数据。请使用下载的标准模板。"}),400
    count = persist_imported_rows(weigh_rows, inspect_rows)
    con=db(); con.execute("INSERT INTO import_jobs(filename,rows_count,result,created_at) VALUES(?,?,?,?)", ("、".join(filenames),count,"已入库并执行9项规则",datetime.now().isoformat(timespec="seconds")));con.commit();con.close()
    return jsonify({"ok":True,"count":count,"weigh":len(weigh_rows),"inspect":len(inspect_rows),"message":f"成功写入 {count} 笔业务，已自动匹配质检单并完成合规核验。"})

@app.get("/api/templates/standard")
def standard_template():
    wb = Workbook(); weigh = wb.active; weigh.title = "过磅单"; inspect = wb.create_sheet("质检单")
    weigh.append(WEIGH_HEADERS); inspect.append(INSPECT_HEADERS)
    weigh.append(["GB20260905001", "PC20260905001", "2026-09-05", "示例供应商", "510000199001011234", "否", "川A1B23C", "废铁", 12000, 4000, 8000, "二次过磅", "过磅员姓名", "司机姓名"])
    inspect.append(["ZJ20260905001", "GB20260905001", "2026-09-05", "废铁", "重型废钢", "B级", 7900, "1.2%", 2.65, 20935, "质检员姓名", "复核人姓名", "自然人自行代开发票"])
    for ws in (weigh, inspect):
        ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]: cell.font = __import__('openpyxl').styles.Font(bold=True, color="FFFFFF"); cell.fill = __import__('openpyxl').styles.PatternFill("solid", fgColor="167A59")
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 18
    stream=io.BytesIO(); wb.save(stream); stream.seek(0)
    return send_file(stream,as_attachment=True,download_name="绿循智审_标准导入模板.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/api/admin")
def admin_data():
    return jsonify({"storage":"SQLite · greenloop.db","backend":"Flask · Python","rules":[{"code":a,"name":b,"owner":"合规风控部","enabled":True,"updated":"2026-09-04"} for a,b in RULES],"accounts":[{"account":"001","name":"赛事演示管理员","role":"合规管理","status":"启用"},{"account":"finance","name":"财务审核员","role":"财务审核","status":"启用"}]})

@app.post("/api/admin/reset-demo")
def reset_demo():
    reset_demo_data()
    return jsonify({"ok": True, "message": "演示数据已恢复：包含合规通过、高风险待审批、例外审批和已付款业务。"})

@app.get("/api/export")
def export():
    wb=Workbook(); ws=wb.active; ws.title="合规台账"
    headers = ["业务编号","供应商","货物品名","净重(Kg)","金额(元)","通过规则","风险等级","核验状态","审批状态","付款状态","最后更新时间"]
    ws.append(headers)
    rows = get_rows()
    sort = request.args.get("sort", "risk_desc")
    if sort == "risk_desc": rows.sort(key=lambda r: (r["passed"] == 9, r["passed"], r["updated_at"]))
    elif sort == "risk_asc": rows.sort(key=lambda r: (r["passed"] != 9, -r["passed"], r["updated_at"]))
    elif sort == "amount_desc": rows.sort(key=lambda r: r["amount"], reverse=True)
    elif sort == "updated_desc": rows.sort(key=lambda r: r["updated_at"], reverse=True)
    for r in rows:
        ws.append([r["id"],r["supplier"],r["goods"],r["net"],r["amount"],f"{r['passed']}/9","高风险" if r["passed"] < 9 else "合规",r["verification_status"],r["approval"],r["status"],r["updated_at"]])
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    for cell in ws[1]:
        cell.font = __import__('openpyxl').styles.Font(bold=True, color="FFFFFF")
        cell.fill = __import__('openpyxl').styles.PatternFill("solid", fgColor="167A59")
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width=18
    stream=io.BytesIO(); wb.save(stream); stream.seek(0)
    return send_file(stream,as_attachment=True,download_name="绿循智审_合规台账.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    init_db(); app.run(debug=False, host="0.0.0.0", port=int(os.environ.get("PORT",5000)))
else: init_db()
