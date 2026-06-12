# TOKEN消耗自动播报

每天 13:30 取截止到前一天的数据，生成 VIP-THINK 风格的 AI 金额消耗海报，并发送到钉钉群。

## 规则

- 数据日期：默认取前一天，中文文案里的“当日”也指前一天。
- 总金额：只统计监督名单内成员的金额总计。
- 低消耗点名：按金额口径，金额低于 100 的成员进入点名。
- 0 消耗点名：单独统计前一天金额为 0 的成员。
- 未匹配：SmartBI 数据里没有匹配到监督名单成员时，显示为“未匹配”。
- 通知形式：优先发送图片海报，标题带 VIP-THINK 标识。
- 钉盘清理：上传到钉盘的 token-usage 图片保留 168 小时，也就是 7 天后清理。

## 文件说明

- `scripts/token_usage_dingtalk_report.py`：生成日报、海报和钉钉通知。
- `scripts/smartbi_cli.py`：SmartBI 只读导出能力，被日报脚本调用。
- `configs/token_usage_dingtalk.example.json`：公开版主配置示例，不含真实群 ID 和人员 open id。
- `configs/token_usage_dingtalk.smartbi_xlsx.example.json`：手动导出 Excel 兜底示例。
- `configs/token_usage_smartbi_tasks.example.json`：SmartBI 自动导出任务示例，不含真实 report id。
- `.env.example`：本地环境变量示例。

## 本地运行

先复制示例配置到本地私有配置：

```powershell
Copy-Item configs/token_usage_dingtalk.example.json configs/token_usage_dingtalk.local.json
Copy-Item configs/token_usage_smartbi_tasks.example.json configs/token_usage_smartbi_tasks.local.json
```

把真实凭证只写入本地 `.env` 或 shell 环境变量，不要提交到 GitHub：

```text
SMARTBI_USERNAME=
SMARTBI_PASSWORD=
DINGTALK_GROUP_OPEN_CONVERSATION_ID=
```

Dry-run：

```powershell
python scripts/token_usage_dingtalk_report.py --config configs/token_usage_dingtalk.local.json --json
```

确认海报和数据无误后发送：

```powershell
python scripts/token_usage_dingtalk_report.py --config configs/token_usage_dingtalk.local.json --send --json
```

## 安全边界

不要上传以下内容：

- `.env`
- `configs/*.local.json`
- SmartBI 密码、钉钉群 ID、open_dingtalk_id
- `outputs/` 里的导出文件、日报 JSON/MD/PNG
- 浏览器 cookie、截图、日志
