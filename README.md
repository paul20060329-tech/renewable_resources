# 绿循智审 - 再生金属合规校验平台

## 本地运行

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

打开 `http://127.0.0.1:5000`。

## 部署 Render

1. 将项目推送至 GitHub。
2. 在 Render 选择 **New +** → **Blueprint**，连接该仓库。
3. Render 会读取 `render.yaml` 并自动安装、启动。
4. 免费实例休眠后首次访问会有短暂冷启动；比赛正式使用建议本地运行或使用付费实例。

## 演示流程

1. 在“数据接入”查看两类标准单据及字段映射。
2. 在“合规核验”筛选风险业务，打开业务详情查看九项规则证据。
3. 对高风险业务执行“例外审批”，再进入“付款审核”生成付款凭证。
4. 在“合规台账”导出留档数据。
