# 绿循智审 - 再生金属合规校验平台

## 本地运行

首次运行：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --index-url https://pypi.org/simple -r requirements.txt
python app.py
```

后续运行只需：

```powershell
.\.venv\Scripts\Activate.ps1
python app.py
```

打开 `http://127.0.0.1:5000`。

## 部署 Render

1. 将项目推送至 GitHub。
2. 在 Render 选择 **New +** → **Blueprint**，连接该仓库。
3. Render 会读取 `render.yaml` 并自动安装、启动。
4. 免费实例休眠后首次访问会有短暂冷启动；比赛正式使用建议本地运行或使用付费实例。

## 账号

`001 / 123456`：赛事演示管理员，可访问审批和后台管理。

## 如何实验

1. 登录后进入“单据导入”，点击“下载标准 Excel 模板”。
2. 在模板的“过磅单”“质检单”工作表填写一笔业务，保持第一行字段名不变。
3. 上传模板，系统会真正写入 SQLite 台账，并自动执行 9 项规则。
4. 到“核验中心”查看规则矩阵，点击“证据详情”查看每项规则的判断依据。
5. 到“风险审批”，对风险业务点击“例外审批通过”或“驳回补件”。
6. 到“付款二维码”，查看该业务的真实二维码，再点击“确认付款”。
7. 到“合规台账”导出 Excel 留档；到“后台管理”查看规则、账号和后端数据服务状态。

## 数据管理

本地业务数据保存在项目根目录的 `greenloop.db`。比赛演示如需恢复初始数据，在停止服务后删除这个文件，再重新运行 `python app.py`。
