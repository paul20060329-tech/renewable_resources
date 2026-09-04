import base64
import io
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

DEMO_ROWS = [
    {"id":"GB20260825246","batch":"PC20260825246","supplier":"张建国","id_no":"510104197803152417","registered":"否","plate":"川A3F72P","goods":"废不锈钢","gross":12560,"tare":5820,"net":6740,"method":"一次过磅","weigher":"李军","driver":"王强","weigh_date":"2026-08-25","inspect_id":"ZJ20260825246","inspect_date":"2026-08-26","spec":"一级统料","grade":"A级","inspect_qty":6705,"moisture":0.52,"price":9.8,"amount":65709,"inspector":"刘梅","reviewer":"赵晓","invoice":"reverse","annual":382000,"frequency":2,"status":"已付款","approval":"无需审批"},
    {"id":"GB20260829182","batch":"PC20260829183","supplier":"陈丽","id_no":"","registered":"是","plate":"川A8K91L","goods":"废杂铝合金","gross":9850,"tare":4650,"net":5200,"method":"二次过磅","weigher":"","driver":"陈丽","weigh_date":"2026-08-29","inspect_id":"ZJ20260829182","inspect_date":"2026-08-29","spec":"破碎杂铝","grade":"C级","inspect_qty":4300,"moisture":8.7,"price":13.2,"amount":68640,"inspector":"孙杰","reviewer":"","invoice":"none","annual":4960000,"frequency":8,"status":"高风险","approval":"待例外审批"},
    {"id":"GB20260903001","batch":"PC20260903001","supplier":"周伟","id_no":"510105198611042010","registered":"否","plate":"川A6M28Q","goods":"废铁","gross":18320,"tare":7320,"net":11000,"method":"二次过磅","weigher":"王磊","driver":"周伟","weigh_date":"2026-09-03","inspect_id":"ZJ20260903001","inspect_date":"2026-09-03","spec":"重型废钢","grade":"B级","inspect_qty":10870,"moisture":1.4,"price":2.65,"amount":28805.5,"inspector":"陈敏","reviewer":"何露","invoice":"self","annual":684000,"frequency":3,"status":"合规通过","approval":"无需审批"},
    {"id":"GB20260904007","batch":"PC20260904007","supplier":"刘芳","id_no":"510107197906184328","registered":"否","plate":"川A9D53K","goods":"废铜","gross":4680,"tare":1980,"net":2700,"method":"一次过磅","weigher":"李军","driver":"刘芳","weigh_date":"2026-09-04","inspect_id":"ZJ20260904007","inspect_date":"2026-09-03","spec":"紫杂铜","grade":"A级","inspect_qty":2700,"moisture":0.2,"price":61.0,"amount":164700,"inspector":"刘梅","reviewer":"赵晓","invoice":"reverse","annual":1180000,"frequency":4,"status":"高风险","approval":"例外审批通过"},
]

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    con = db()
    con.execute('''CREATE TABLE IF NOT EXISTS businesses (id TEXT PRIMARY KEY, payload TEXT NOT NULL, status TEXT NOT NULL, approval TEXT NOT NULL, updated_at TEXT NOT NULL)''')
    if not con.execute("SELECT 1 FROM businesses LIMIT 1").fetchone():
        import json
        now = datetime.now().isoformat(timespec="seconds")
        for item in DEMO_ROWS:
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
    import json
    con = db(); saved = con.execute("SELECT payload,status,approval,updated_at FROM businesses ORDER BY updated_at DESC").fetchall(); con.close()
    records = []
    for x in saved:
        row = json.loads(x["payload"]); row.update(status=x["status"], approval=x["approval"], updated_at=x["updated_at"])
        row["checks"] = check(row); row["passed"] = sum(v["passed"] for v in row["checks"]); records.append(row)
    return records

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
    return jsonify({"stats":{"total":total,"passed":sum(r["passed"] == 9 for r in rows),"risk":len(risks),"paid":sum(r["status"] == "已付款" for r in rows),"amount":sum(r["amount"] for r in rows)},"records":rows,"rules":[{"code":a,"name":b} for a,b in RULES]})

@app.post("/api/business/<business_id>/approve")
def approve(business_id):
    con = db(); row = con.execute("SELECT payload FROM businesses WHERE id=?", (business_id,)).fetchone()
    if not row: return jsonify({"error":"业务不存在"}), 404
    con.execute("UPDATE businesses SET status=?, approval=?, updated_at=? WHERE id=?", ("例外审批通过", "管理层例外审批通过", datetime.now().isoformat(timespec="seconds"), business_id)); con.commit(); con.close()
    return jsonify({"ok":True, "message":"已记录管理层例外审批，业务现可进入付款审核。"})

@app.post("/api/business/<business_id>/pay")
def pay(business_id):
    rows = {r["id"]:r for r in get_rows()}; row=rows.get(business_id)
    if not row: return jsonify({"error":"业务不存在"}),404
    if row["passed"] < 9 and "通过" not in row["approval"]: return jsonify({"error":"高风险业务须完成例外审批后才可付款"}),400
    con=db(); con.execute("UPDATE businesses SET status=?, updated_at=? WHERE id=?",("已付款",datetime.now().isoformat(timespec="seconds"),business_id));con.commit();con.close()
    return jsonify({"ok":True,"message":"财务付款确认成功，已同步至全局合规台账。"})

@app.get("/api/business/<business_id>/payment-code")
def payment_code(business_id):
    rows = {r["id"]: r for r in get_rows()}
    row = rows.get(business_id)
    if not row:
        return jsonify({"error": "业务不存在"}), 404
    if row["passed"] < 9 and "通过" not in row["approval"]:
        return jsonify({"error": "该业务尚未满足付款条件"}), 403
    import qrcode
    payload = f"GREENLOOP|{row['id']}|{row['supplier']}|{row['amount']:.2f}|PAYMENT_REVIEW"
    image = qrcode.make(payload)
    stream = io.BytesIO()
    image.save(stream, format="PNG")
    stream.seek(0)
    return send_file(stream, mimetype="image/png", max_age=0)

@app.post("/api/import")
def upload():
    files=request.files.getlist("files")
    if not files: return jsonify({"error":"请选择 Excel 文件"}),400
    counts=[]
    for f in files:
        try:
            wb=load_workbook(f,read_only=True,data_only=True); ws=wb.active; counts.append({"name":f.filename,"rows":max(ws.max_row-1,0)})
        except Exception: return jsonify({"error":f"无法读取 {f.filename}，请使用 .xlsx 格式"}),400
    return jsonify({"ok":True,"files":counts,"message":"文件结构已读取。比赛演示版保留原始字段并进入待核验队列。"})

@app.get("/api/export")
def export():
    wb=Workbook(); ws=wb.active; ws.title="合规台账"
    ws.append(["业务编号","供应商","货物品名","净重(Kg)","金额(元)","通过规则","风险等级","审批状态","付款状态","最后更新时间"])
    for r in get_rows(): ws.append([r["id"],r["supplier"],r["goods"],r["net"],r["amount"],f"{r['passed']}/9","合规" if r["passed"]==9 else "高风险",r["approval"],r["status"],r["updated_at"]])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width=18
    stream=io.BytesIO(); wb.save(stream); stream.seek(0)
    return send_file(stream,as_attachment=True,download_name="绿循智审_合规台账.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    init_db(); app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT",5000)))
else: init_db()
