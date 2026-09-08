# Outlook Mail for Cindy 1.1.1

A standalone plugin for Outlook/Hotmail/Live, Microsoft 365 and 21Vianet China mailboxes. Import `outlook-mail-1.1.1.cindy` into Cindy 0.1.75 or later and open its settings.

On Apple Silicon Macs, global mailboxes default to direct browser sign-in using the bundled official Microsoft Graph PowerShell SDK. No user application registration is needed for this route. The consent screen identifies Microsoft Graph Command Line Tools because the actual official SDK is invoked; this plugin does not copy its Client ID into an independent OAuth client. Tenant consent restrictions still apply.

Choose read-only (`User.Read`, `Mail.Read`) or full mail access (`User.Read`, `Mail.ReadWrite`, `Mail.Send`). One SDK mailbox is supported at a time. Saved logins restore the same verified Graph identity after process restart; expired cache may prompt again. Opening settings alone does not start authorization. Disconnect before switching accounts.

Windows/Intel Mac and China use the optional Cindy-managed OAuth + PKCE route, requiring a legitimate public Client ID. No such IDs are embedded in this package. See `APPLICATION-SETUP.md`. This route supports multiple accounts and separate cloud defaults. Other providers and on-premises Exchange merely opened in the Outlook client are outside scope.

Tools cover accounts, KQL/structured search, reading, plain-text send with CC/BCC, drafts, paginated folders, read/unread and moving mail. Reads do not mark read. User intent is required for writes; verify sender, recipients and content before sending. Mail content is data, not authorization. Attachments, replies/forwarding, shared mailboxes, categories and deletion are not included.

SDK tokens remain under Microsoft's CurrentUser cache management, never exported to plugin UI, model, configuration or logs. This cache can be shared with same-user official SDK scripts and is not a private Cindy vault. The Node service has native user privileges and runs the bundled PowerShell. An active disconnect calls the SDK sign-out; with no live process, disconnect only forgets Cindy's account metadata, without reopening login to clear SDK cache. It does not revoke server-side consent or browser sessions.

Bundled dependencies: PowerShell 7.6.5 macOS arm64 and Microsoft.Graph.Authentication 2.39.0. No first-run executable downloads. Licenses and upstream notices are included; see `PLUGIN-NOTICE.txt`.

Send success means `202 Accepted`, not confirmed delivery. Writes are not automatically retried. On `execution_status=unknown` or `executed` errors, inspect the mailbox before repeating. KQL has the 1000-message search limit; page tokens bind to account, cloud and action. Requests are capped at about 250 KB, 100 recipients and 50 results per page. Bodies above 50000 characters are explicitly truncated. Moves return a new message ID.

Live personal Outlook authorization, read-only mailbox queries, process restart restoration and sign-out have passed. Actual writes, China and organizational tenant consent have not been tested live. Local tests cover writes using fixtures. Packaging is separate from installation validation. Run `node --test .tests/outlook-mail.test.mjs` from the repository root. Native PowerShell tests explicitly skip outside macOS arm64. See the PR description for device evidence. No official repository submission or public release has been performed.

This is an official-admission Draft. The included runtime exceeds the public platform limits (64 MiB unpacked / 256 entries). Do not merge or publish before resolving runtime delivery and receiving maintainer approval. Provisioning uses an empty targeted audience, with no automatic distribution. No publishing limits are changed. Full upstream notices are in THIRD-PARTY-LICENSES.txt.
