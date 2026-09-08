# 微软应用注册与发布配置

本说明只适用于「使用自有微软应用」登录及世纪互联中国区。1.1.0 在 Apple Silicon Mac 上的全球版默认使用微软官方 SDK 登录，不需要用户注册应用。

自有应用由插件维护者／企业 IT 一次性配置。发布者可配置公开 Client ID 后重新打包，普通用户不应逐个创建应用。当前包未内置自有应用 ID，因此不能承诺中国区或其他设备零配置连接。

## 两个独立注册

| 项目 | 全球版 | 世纪互联中国区 |
| --- | --- | --- |
| 注册门户 | https://portal.azure.com | https://portal.azure.cn |
| 目标账号 | 任意组织目录 + 个人 Microsoft 账号 | 任意中国区组织目录 |
| 授权域名 | login.microsoftonline.com | login.chinacloudapi.cn |
| Graph 域名 | graph.microsoft.com | microsoftgraph.chinacloudapi.cn |
| 桌面回调地址 | http://127.0.0.1:53687/callback | http://127.0.0.1:53688/callback |
| 插件凭证键 | outlook_global | outlook_china |

1. 在相应门户的 Microsoft Entra ID → 应用注册中新建应用，名称可设为「Cindy Outlook Mail」。选择表中支持的账号类型。全球版需同时支持个人账号和组织账号；中国区不使用个人账号。
2. 配置为**移动和桌面应用／公共客户端**，注册表中精确的回调地址。不要注册成需要客户端密钥的 Web 应用或 SPA。
3. 微软门户普通输入框可能拒绝 `http://127.0.0.1`。按当前门户支持的应用清单编辑功能配置公共客户端回调（Microsoft Graph 格式为 `publicClient.redirectUris`，旧式清单使用 `replyUrlsWithType` + `InstalledClient`）。不要擅自将地址换成 `localhost`，因为 Cindy 固定监听并发送 `127.0.0.1` 回调。
4. 添加 Microsoft Graph **委托权限** `User.Read`、`Mail.ReadWrite`、`Mail.Send`。插件另外请求 `openid` 和 `offline_access`，用于登录及长期刷新。不要申请 Application 权限或全租户邮件访问。
5. 公共桌面应用无需生成客户端密钥。复制「应用程序（客户端）ID」，在插件所选区域的「应用配置」填写。需要管理员同意时由租户管理员完成；实际企业可能还要求发布者验证或允许此应用。
6. 使用获授权的测试账号登录；测试读取与搜索、保存草稿、标记状态、移动测试邮件。只有得到测试收件人和内容授权后再做发送测试。
7. 在全球版至少验证个人 Outlook 和组织 Microsoft 365 各一个账号；中国区验证一个世纪互联测试账号。验证重新授权后同一账号不重复添加，以及实际令牌过期刷新。

## 发布者内置公开 Client ID

最终发布时把对应 UUID 写入 `ghost.json` 的 `network.secrets[].oauth.clientId`，每个区域使用自己注册的应用。只内置 **Client ID**，不要写入密码、客户端密钥、Access Token 或 Refresh Token。

父目录的 `scripts/configure-outlook-app.mjs` 可做结构化更新：

```sh
node scripts/configure-outlook-app.mjs global <已注册的全球版应用Client-ID>
node scripts/configure-outlook-app.mjs china <已注册的中国区应用Client-ID>
```

脚本只接受公开 UUID，只更新选定区域的 clientId。注册设置和真实登录验证仍必须完成，UUID 格式正确不代表应用已正确注册。没有中国区应用时，不能宣称中国区开箱即用。

完整实测后用 Cindy Forge 重新打包。官方仓的新插件准入需要提案、维护者确认、明确受众和实机验证；当前工作没有创建 issue、PR 或发布。

## 官方参考

- 云区与端点：https://learn.microsoft.com/en-us/graph/deployments
- 桌面应用注册：https://learn.microsoft.com/en-us/entra/identity-platform/scenario-desktop-app-registration
- 回调限制及 127.0.0.1：https://learn.microsoft.com/en-us/entra/identity-platform/reply-url
- Graph 邮件列表：https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0
- 发送接受状态：https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0
