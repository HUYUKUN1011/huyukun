# TOKEN消耗自动播报说明

这套脚本用于每天 13:30 播报前一天 AI 金额消耗，输出图片海报并发送到钉钉群。

## 核心口径

- 取数时间：每天 13:30。
- 数据范围：截止到前一天。
- 总金额：监督名单内成员金额总计。
- 低消耗点名：金额低于 100。
- 0 消耗点名：前一天金额为 0。
- 未匹配：SmartBI 记录无法匹配到监督名单成员时标注为“未匹配”。
- 图片清理：钉盘图片保留 168 小时。

## 公开配置

公开仓库只保留示例配置。真实人员、群 ID、SmartBI report id、密码和本地输出不要提交。

## 常用命令

```powershell
python scripts/token_usage_dingtalk_report.py --config configs/token_usage_dingtalk.local.json --json
python scripts/token_usage_dingtalk_report.py --config configs/token_usage_dingtalk.local.json --send --json
```
