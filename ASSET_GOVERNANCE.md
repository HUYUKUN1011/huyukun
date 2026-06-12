# 资产治理登记表

本文件记录 `HUYUKUN1011/huyukun` 仓库内资产的治理状态。该仓库按公开 GitHub 仓库标准管理，不保存真实凭证、内部链接、成员名单或业务数据。

## 仓库总览

| 项目 | 内容 |
| --- | --- |
| 仓库 | `HUYUKUN1011/huyukun` |
| 所属团队 | 教学 |
| 负责人 | 胡煜坤 |
| 备份负责人 | 待补充 |
| 最近维护日期 | 2026-06-12 |
| 仓库建议状态 | governed |

## 状态定义

| 状态 | 含义 |
| --- | --- |
| governed | 资产用途、负责人、运行说明、示例配置、安全边界和维护状态已明确。 |
| partial_governed | 已有部分说明，但仍缺少脚本、运行验证、来源说明或安全校验记录。 |
| not_governed | 尚未按资产治理要求整理。 |
| indexed_only | 仅被收录到资产清单，尚未补齐治理信息。 |
| unknown | 资产用途、负责人或状态无法确认。 |
| exclude | 明确不纳入治理统计。 |

## 资产登记

| 资产名称 | 路径 | 用途 | 负责人 | 备份负责人 | 所属团队 | 治理状态 | 运行/验证方式 | 配置文件 | 安全说明 | 最近维护日期 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| TOKEN消耗自动播报 | `TOKEN消耗自动播报/` | 每天生成 VIP-THINK 风格的 AI 金额消耗海报，并支持发送到钉钉群。 | 胡煜坤 | 待补充 | 教学 | governed | 参考 `TOKEN消耗自动播报/README.md`，先执行 dry-run：`python scripts/token_usage_dingtalk_report.py --config configs/token_usage_dingtalk.local.json --json`；确认无误后再加 `--send`。 | `.env.example`、`configs/*.example.json` | 真实 SmartBI 凭证、钉钉群 ID、open_dingtalk_id、导出文件和日志不得提交。 | 2026-06-12 |
| 周报自动化提醒 | `周报自动化提醒/` | 教学运营周报提醒、会后待办确认、下周周报模板整理与飞书写回。 | 胡煜坤 | 待补充 | 教学 | governed | 参考 `周报自动化提醒/README.md`，默认 dry-run：`.\Write-FeishuOperationWeeklyReport.ps1 -Mode dry-run -ConfigPath .\operation_weekly_report_automation.config.example.json`。 | `operation_weekly_report_automation.config.example.json` | 真实飞书链接、钉钉群会话 ID、机器人 webhook、成员名单、会议记录原文不得提交。 | 2026-06-12 |
| 班级申请审批 | `SKILL.md` | 自动审批中英文班级申请邮件，按关键词搜索邮件并以回复形式发送确认。 | 胡煜坤 | 待补充 | 教学 | partial_governed | 当前仅有技能说明；需要补齐或关联实际脚本 `approve-class-emails.ps1` 后，才能完成运行验证。 | 通过命令参数传入邮箱和授权码，后续应补充 `.env.example` 或参数示例。 | IMAP/SMTP 授权码、真实邮箱、抄送人策略和邮件内容不得泄露；公开说明中只能保留示例值。 | 2026-06-12 |
| 薪酬表技能上传包 | `compensation-workbook-skill-upload.zip` | `compensation-workbook` 技能上传包，用于生成薪酬 workbook。zip 内含 `SKILL.md`、`agents/openai.yaml`、`scripts/create_compensation_workbook.py`、`README.md`。 | 胡煜坤 | 待补充 | 教学 | partial_governed | 需要补充来源、生成方式、使用场景和安全校验记录；使用前应先解压到本地检查内容。 | zip 内含技能配置和脚本，暂无根目录公开配置说明。 | 保留 zip 前需确认不含真实薪酬数据、人员明细、内部模板、凭证或私有路径。 | 2026-06-12 |

## 待补齐事项

- 为 `SKILL.md / 班级申请审批` 补齐或关联实际脚本 `approve-class-emails.ps1`，并增加 dry-run 或最小验证方式。
- 为 `compensation-workbook-skill-upload.zip` 补充来源、生成方式、使用场景和敏感信息检查结果；如后续频繁维护，建议解压成目录管理。
- 补充备份负责人。
- 新增资产时同步更新根目录 README 和本登记表。
