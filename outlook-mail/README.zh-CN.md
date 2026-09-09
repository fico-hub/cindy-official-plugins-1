# Outlook 邮箱插件 2.0.0

搜索、阅读、起草、发送和整理 Outlook、Hotmail、Live 与 Microsoft 365 云端邮件。全球版和世纪互联中国区分别管理多账号及默认账号。不支持其他邮箱服务商或仅使用 Outlook 客户端打开的本地 Exchange 邮箱。

邮件请求统一经过：

`插件邮件逻辑 → cindy.fetch → Host 注入 OAuth Token → Microsoft Graph`

插件运行于 Cindy 浏览器沙箱，不包含 Node worker、PowerShell、微软 SDK、本机可执行文件或运行时下载。Host 负责打开微软登录、校验回调、交换与刷新令牌、保存凭证，并仅向声明的 Graph 主机注入 Authorization。插件经 `/oauth` 获取账号元数据，无法读取或保存令牌。每次邮件请求带当前 `callId` 和明确解析的 `authAccount`。

## 连接

安装到 Cindy 0.1.75 或更新版本，在插件详情页选择邮箱所属区域并连接账号。设置页支持中文、英文、日文和韩文。发布者或企业 IT 需要为各服务区域配置已注册的公共微软应用 Client ID，无需客户端密钥。当前源码尚无已注册的 Client ID，未配置发布者应用或自定义应用前，连接按钮保持禁用。详见[应用配置](APPLICATION-SETUP.md)。发布者配置完成后，普通用户只需授权自己的账号。

全球版申请 User.Read、Mail.ReadWrite、Mail.Send、openid 和 offline_access；中国区使用对应中国区 Graph 资源权限。这些均为用户委托权限，不是全租户 Application 权限；租户策略可能要求管理员批准。连接、断开和默认账号管理全部交给 Host，插件 KV 仅保存默认服务区域。

## 邮件操作

提供账号列表、KQL／结构化搜索、阅读、带 CC/BCC 的纯文本发送、草稿、分页文件夹、已读／未读及移动。读取不标记已读。写操作必须来自明确用户意图，并核对发件账号、收件人和内容。发送成功只表示微软接受，不代表送达。执行状态为 executed 或 unknown 的失败，必须先检查邮箱再决定是否重复。插件不重试邮件请求；OAuth 刷新及相关重试由 Host 决定。

请求上限约 250 KB、100 个收件人、每页 50 条。分页绑定账号、区域与操作。Microsoft KQL 搜索最多返回 1000 封；正文超过 50000 字符明确标记截断；移动后使用返回的新 ID。不包含附件、回复／转发、共享邮箱、分类或删除。

## 升级和验证

2.0.0 替换 SDK 授权路径，旧 SDK 会话与保存的账号元数据不再使用，需要通过 Host OAuth 重新连接。不会读取或清理其他微软工具的凭据缓存，也不会从旧 KV 自动恢复登录。

运行 `node --test .tests/outlook-mail.test.mjs` 和仓库规定的契约、本地化、provisioning、发布流程检查。提交后使用 `.github/scripts/package-plugin.sh outlook-mail /tmp/outlook-mail-2.0.0.cindy` 打包。浏览器测试使用模拟 Host 接口与虚构账号，不作为真实 OAuth 证据；见 `.tests/outlook-mail/settings-browser.mjs`。

由于尚无已注册的应用 ID，本版本尚未完成真实 Host OAuth 登录或正式客户端邮件验收，不能沿用早期 SDK 开发包的实机结果。新插件准入和最终实机验收仍需维护者审查。provisioning 是空定向名单，不自动分发；PR 保持正式待审，测试及打包限制不作放宽。
