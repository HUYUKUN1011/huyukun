---
name: 班级申请审批
description: 自动审批中英文班级申请邮件，搜索主题含"中英"的邮件，以回复形式标记已读并回复确认（含抄送），保存到已发送
---

## 功能
- 连接腾讯企业邮箱 IMAP，扫描近30天邮件
- 搜索主题含关键词（默认"中英"）的邮件
- 自动标记已读/已回复
- 以 **回复形式**（In-Reply-To + References 头）发送审批回复，在客户端显示为原邮件回复
- 自动 **抄送** ruiyu.zong@hltn.com, jian.chen@hltn.com
- 保存回复到"Sent Messages"已发送文件夹

## 使用前提
- 需要 **IMAP/SMTP 授权码**（登录网页版邮箱 → 设置 → 邮箱绑定 → 生成新密码）
- IMAP: imap.exmail.qq.com:993 (SSL)
- SMTP: smtp.exmail.qq.com:465 (SSL)

## 运行
```powershell
.\approve-class-emails.ps1 -Email "your@company.com" -AuthCode "your_code" -Keyword "中英"
```

## 参数
- `-Email`: 邮箱地址
- `-AuthCode`: IMAP/SMTP 授权码（不是登录密码）
- `-Keyword`: 搜索关键词，默认 "class application"
