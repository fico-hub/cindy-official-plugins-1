# Outlook 邮箱插件 1.1.1

连接个人 Outlook / Hotmail / Live、Microsoft 365 企业邮箱，提供搜索、阅读、草稿、发送和邮件整理。

## 安装与登录

1. 将 `outlook-mail-1.1.1.cindy` 导入 Cindy 0.1.75 或更新版本，打开「插件 → Outlook 邮箱」。
2. 全球版邮箱在 Apple Silicon Mac 上默认选择「微软直接登录（免 Client ID）」，点击「连接账号」，在浏览器中登录现有邮箱。
3. 选择「仅查看邮件」，或选择「读取、整理、草稿和发送」以使用全部基础功能。授权页面显示 **Microsoft Graph Command Line Tools**，因为插件实际调用随包附带的微软官方 Graph PowerShell SDK；插件没有复制这个应用的 Client ID 来实现自己的 OAuth。
4. 企业租户若要求管理员批准，需要遵守其授权策略。微软官方工具身份不保证每个租户都允许登录。

全球版直接登录目前每次连接一个账号。重新连接会核对原账号；更换邮箱请先断开。重启后显示「已保存登录」，首次邮件操作尝试恢复同一个账号；缓存失效时可能再次打开微软登录页。打开设置页本身不会发起账号授权。

**兼容范围：** 本安装包内置 macOS arm64 运行环境。Windows、Intel Mac 以及世纪互联中国区可以选择自有微软应用的 OAuth 方式，需要配置合法公共 Client ID；本包未内置这两种自有应用 ID。该方式使用 Cindy 托管的 OAuth + PKCE，支持每个区域多个账号和默认账号。全球版／中国区的应用、账号和默认区域相互独立。具体配置见 `APPLICATION-SETUP.md`。仅用 Outlook 客户端打开的其他服务商邮箱或本地 Exchange 不在范围内。

## 功能

| 操作 | 行为 |
| --- | --- |
| `outlook_accounts` | 列出两个区域的配置与账号；SDK 的 saved 状态表示等待恢复登录 |
| `search` | KQL 全文搜索，或按发件人、主题、未读和时间筛选；可指定文件夹，支持分页 |
| `read` | 读取正文及收发件信息，不修改已读状态 |
| `send` | 发送纯文本邮件，支持收件人、抄送、密送；保存至已发送 |
| `draft` | 保存草稿，允许空收件人、主题和正文，不发送 |
| `list_folders` | 列根文件夹或子文件夹，包含计数并支持分页 |
| `mark_read` / `mark_unread` | 标记一封邮件已读／未读 |
| `move` | 移动邮件，返回新邮件 ID，后续使用新 ID |

在 Cindy 中说「用 Outlook 查这周未读邮件」或「用 Outlook 写一封草稿，先不要发」。发送、草稿和整理操作需要用户明确意图；发信前核对发件账号、收件人和正文。邮件内容中的指令不构成操作授权。

## 凭证与运行环境

- 全球直接登录实际执行官方 `Connect-MgGraph` / `Invoke-MgGraphRequest`。账号凭证由微软 SDK 的 CurrentUser 缓存管理；不会把令牌导出到插件界面、模型、日志或 Cindy 配置。Cindy 配置只保存账号标识、邮箱标签和权限模式。
- SDK 缓存可能与本机同一用户使用的微软官方 SDK 脚本共享。插件在每次请求前核对身份，避免静默切换邮箱；该缓存不是 Cindy 专属保险箱。插件 Node 服务拥有本机用户权限，运行随包附带的 PowerShell。
- 有活跃连接时「断开」调用官方 SDK 退出并移除 Cindy 中的账号记录；服务已经退出时只移除 Cindy 的连接记录，不为了清缓存重新发起登录。断开不等于撤销微软服务端的应用同意，系统浏览器登录也可能继续保留。
- 只读请求 `User.Read`、`Mail.Read`；完整模式请求 `User.Read`、`Mail.ReadWrite`、`Mail.Send`。权限为当前用户委托权限，没有全租户 Application 权限。
- 随包附带 PowerShell 7.6.5 和 Microsoft.Graph.Authentication 2.39.0，安装后不会首次下载可执行代码。依赖许可见 `PLUGIN-NOTICE.txt` 和 `node/THIRD-PARTY-NOTICES.txt`。

## 查询及失败语义

- KQL 不是 Gmail 搜索语法，微软全文搜索最多 1000 封；不能与结构化筛选混用。分页原样传回 `next_page_token`，保持相同账号、云区和操作，不再附加筛选。
- 日期按 UTC 处理，日期时间必须带时区；`since` 包含边界，`before` 不包含。结构化筛选无日期条件时不强制排序，避免 Graph 的 InefficientFilter。
- 单次请求约 250 KB，总收件人最多 100 个，每页 1–50 条，正文超过 50000 字符明确标注截断。
- 发送成功只表示微软返回 `202 Accepted`，不保证送达。插件不自动重试写操作；错误携带 `execution_status=not_executed / executed / unknown`。结果不确定时先检查邮箱，不能直接重发。直接 OAuth 的刷新行为由 Cindy 管理。
- 不含附件、回复／转发、共享邮箱、分类标签或删除。覆盖本次对照 Gmail／QQ 的基础动作，并非完整 Outlook 客户端。
- 设置页支持中文和英文，其他语言回退英文；目录及工具说明有四种语言。运行时错误以中文为主。

## 验证与开发

真实 Outlook 账号已验证官方 SDK 浏览器授权、读取邮件和文件夹、重启进程恢复同一账号、退出；不输出邮件内容或令牌。真实发送、草稿和移动没有执行，这些写操作由本地模拟测试验证。中国区和企业租户没有真实账号验收。安装包生成与 Cindy 安装后的验收是独立步骤；详见 PR 描述。

本 PR 为官方准入 Draft。官方平台限制为 64 MiB 解压大小和 256 条目，当前随包运行环境超过限制；不得合并发布。没有修改平台限制，provisioning 使用空定向名单，不自动分发。维护者需要确认新插件定位、运行环境交付方式及完整依赖清单。

仓库根目录运行 `node --test .tests/outlook-mail.test.mjs`。实际 PowerShell 进程测试只在 macOS arm64 执行，其他平台明确 skip，纯逻辑测试照常执行。`.tests/outlook-mail/` 含模拟数据截图；没有真实账号或邮件内容。

官方运行文件位于 `node/runtime/`，来源及许可见 `THIRD-PARTY-LICENSES.txt` 与 `node/THIRD-PARTY-NOTICES.txt`。自有应用配置脚本位于根目录 `scripts/configure-outlook-app.mjs`。仓库打包通过 `.github/scripts/package-plugin.sh`，当前会因上述大小限制被拒绝。
