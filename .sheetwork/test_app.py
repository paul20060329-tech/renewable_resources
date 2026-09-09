import io
import json
import os
import sys
from pathlib import Path

TEST_DB = Path(__file__).with_name("test_api.db")
if TEST_DB.exists():
    TEST_DB.unlink()
os.environ["DATABASE_PATH"] = str(TEST_DB)
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook
import app

client = app.app.test_client()

def login(username):
    session = app.app.test_client()
    response = session.post("/login", data={"username": username, "password": "123456"}, follow_redirects=False)
    assert response.status_code == 302, (username, response.status_code)
    return session

clients = {name: login(name) for name in app.ACCOUNTS}
roles = {}
for name, session in clients.items():
    response = session.get("/api/dashboard")
    assert response.status_code == 200, (name, response.status_code, response.get_data(as_text=True))
    data = response.get_json()
    roles[name] = set(data["permissions"])
    html = session.get("/").get_data(as_text=True)
    assert ("data-view=\"approval\"" in html) and ("付款审批" in html)
    print("ROLE", name, sorted(data["permissions"]))

assert roles["finance"] == {"overview", "approval", "ledger"}
assert roles["accounting"] == {"overview", "ingest", "verify", "ledger"}
assert roles["cashier"] == {"overview", "ingest", "verify", "payment", "ledger"}
assert roles["admin"] == {"overview", "ingest", "verify", "approval", "payment", "ledger", "admin"}

for name in ("finance", "accounting", "cashier"):
    html = clients[name].get("/").get_data(as_text=True)
    for permission in (set(app.ROLE_PERMISSIONS[app.ACCOUNTS[name][2]]) ^ set(app.ROLE_PERMISSIONS["后台管理员"])):
        if permission not in app.ROLE_PERMISSIONS[app.ACCOUNTS[name][2]]:
            marker = f'data-view="{permission}"'
            marker_pos = html.find(marker)
            assert marker_pos >= 0 and "hidden" in html[marker_pos:marker_pos + 180]

admin = clients["admin"]
assert admin.get("/api/admin").status_code == 200
assert clients["finance"].get("/api/admin").status_code == 403
assert clients["finance"].get("/api/templates/standard").status_code == 403
assert clients["accounting"].get("/api/templates/standard").status_code == 200
standard = clients["accounting"].get("/api/templates/standard")
standard_wb = load_workbook(io.BytesIO(standard.data), read_only=True, data_only=True)
assert standard_wb.sheetnames == ["过磅单", "质检单", "入库单"]
assert "银行卡号" in [c.value for c in standard_wb["过磅单"][1]]
assert not any("微信" in str(c.value or "") or "二维码" in str(c.value or "") for ws in standard_wb.worksheets for row in ws.iter_rows() for c in row)

ledger = clients["accounting"].get("/api/templates/ledger")
ledger_wb = load_workbook(io.BytesIO(ledger.data), read_only=True, data_only=True)
assert ledger_wb.sheetnames == ["合规台账", "字段说明", "银行税务检查要点"]
assert ledger_wb["合规台账"].max_row == 3

rows = {row["id"]: row for row in clients["admin"].get("/api/dashboard").get_json()["records"]}
high = next(row for row in rows.values() if row["passed"] == 9 and row["verified"] and row["amount"] >= app.PAYMENT_THRESHOLD and row["status"] != "已付款")
low = next(row for row in rows.values() if row["passed"] == 9 and row["verified"] and row["amount"] < app.PAYMENT_THRESHOLD and row["status"] != "已付款")
risk = next(row for row in rows.values() if row["passed"] < 9 and row["status"] != "审批驳回")
assert high["current_stage"] == "待审批"
assert low["current_stage"] == "待付款"
assert clients["cashier"].post(f"/api/business/{high['id']}/approve").status_code == 403
assert clients["finance"].post(f"/api/business/{high['id']}/approve").status_code == 200
assert clients["cashier"].post(f"/api/business/{high['id']}/pay").status_code == 200
assert clients["finance"].post(f"/api/business/{low['id']}/pay").status_code == 403
assert clients["accounting"].post(f"/api/business/{risk['id']}/reject", json={"opinion": "测试驳回"}).status_code == 200
assert clients["admin"].get("/api/export").status_code == 200

# Re-importing the new 59-column demo workbook is accepted and replaces rows by ID.
demo_path = Path(app.BASE_DIR) / "Data" / "绿循智审_100条虚拟导入数据.xlsx"
with demo_path.open("rb") as handle:
    imported = clients["accounting"].post("/api/import", data={"files": (handle, demo_path.name)}, content_type="multipart/form-data")
assert imported.status_code == 200, imported.get_data(as_text=True)
assert imported.get_json()["count"] == 100

# The local sync endpoint is idempotent when no downloaded workbook changed.
synced = clients["cashier"].post("/api/local-sync")
assert synced.status_code == 200
assert synced.get_json()["count"] == 0

print("CHECKS OK")
