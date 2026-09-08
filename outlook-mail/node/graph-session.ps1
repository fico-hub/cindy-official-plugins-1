# Development transport: official SDK owns authentication and all access tokens.
param([Parameter(Mandatory)][string] $ModulePath)
$ErrorActionPreference = 'Stop'
$PSStyle.OutputRendering = 'PlainText'
$WarningPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'
# Keep protocol output separate from SDK libraries writing diagnostics to Console.
$protocolOutput = [Console]::Out
[Console]::SetOut([IO.TextWriter]::Null)
Import-Module $ModulePath -ErrorAction Stop
Set-MgRequestContext -MaxRetry 0 -ClientTimeout 30 | Out-Null
$account = $null
$accessMode = $null
$sdkAccount = $null
$sdkTenant = $null
$sdkScope = 'Process'
function Emit($Value) {
    $json = ConvertTo-Json -InputObject $Value -Depth 30 -Compress
    if ([Text.Encoding]::UTF8.GetByteCount($json) -gt 900000) {
        $json = ConvertTo-Json -InputObject @{id=$Value.id;result=@{ok=$true;status=$Value.result.status;body='';truncated=$true}} -Compress -Depth 5
    }
    $script:protocolOutput.WriteLine($json)
}
function Require-Account($Expected) {
    $ctx = Get-MgContext
    if (!$script:account -or !$ctx -or $ctx.Environment -ne 'Global' -or $ctx.ContextScope -ne $script:sdkScope -or $ctx.Account -cne $script:sdkAccount -or $ctx.TenantId -cne $script:sdkTenant -or $Expected -cne $script:account.id) {
        throw 'ACCOUNT_NOT_CONNECTED'
    }
}
try {
    while ($null -ne ($line = [Console]::In.ReadLine())) {
        $request = $null
        $stage = 'validation'
        try {
            if ($line.Length -gt 300000) { throw 'INVALID_REQUEST' }
            $request = ConvertFrom-Json -InputObject $line -AsHashtable
            switch ($request.method) {
                'status' {
                    Emit @{id=$request.id;result=@{ready=$true;account=$account;mode=$accessMode;cloud='global';token_persistence=$(if($sdkScope -eq 'CurrentUser'){'microsoft_sdk'}else{'process_only'})}}
                }
                'connect' {
                    if ($account) { throw 'ALREADY_CONNECTED' }
                    $audience = $request.params.audience
                    $mode = $request.params.mode
                    $flow = $request.params.flow
                    $scope = $request.params.contextScope
                    $hint = $request.params.loginHint
                    $expectedId = $request.params.expectedId
                    if ($audience -notin @('common','consumers','organizations') -or $mode -notin @('read','full') -or $flow -notin @('browser','device') -or $scope -notin @('Process','CurrentUser')) { throw 'INVALID_REQUEST' }
                    if ($hint -and ($hint -isnot [string] -or $hint.Length -gt 320 -or $hint -match '[\r\n\x00]')) { throw 'INVALID_REQUEST' }
                    if ($scope -eq 'CurrentUser' -and $flow -ne 'browser') { throw 'INVALID_REQUEST' }
                    $scopes = if ($mode -eq 'full') { @('User.Read','Mail.ReadWrite','Mail.Send') } else { @('User.Read','Mail.Read') }
                    # Do not provide ClientId or extract the SDK's default application ID.
                    $stage = 'authentication'
                    $auth = @{Environment='Global';TenantId=$audience;Scopes=$scopes;ContextScope=$scope;NoWelcome=$true}
                    if ($hint) { $auth.LoginHint = $hint }
                    if ($flow -eq 'device') { $auth.UseDeviceAuthentication=$true }
                    Connect-MgGraph @auth | ForEach-Object {
                        if ($_ -is [string] -and $_ -match 'https://login\.microsoft\.com/device' -and $_ -match 'code ([A-Z0-9]{9})') {
                            Emit @{event='login_required';verification_uri='https://login.microsoft.com/device';user_code=$Matches[1];wait_seconds=120}
                        }
                    }
                    $stage = 'context_validation'
                    $ctx = Get-MgContext
                    if (!$ctx -or $ctx.AuthType -ne 'Delegated' -or $ctx.Environment -ne 'Global' -or $ctx.ContextScope -ne $scope) { throw 'INVALID_AUTH_CONTEXT' }
                    $stage = 'profile_validation'
                    $profile = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me?$select=id,mail,userPrincipalName' -OutputType Json | ConvertFrom-Json -AsHashtable
                    if (!$profile.id) { throw 'PROFILE_ID_MISSING' }
                    if ($expectedId -and ('sdk:' + $profile.id) -cne $expectedId) { throw 'ACCOUNT_MISMATCH' }
                    # A personal account may lack a username claim in the SDK context.
                    # The Graph /me id is authoritative; keep the SDK context binding separately.
                    $label = $profile.mail ?? $profile.userPrincipalName ?? $ctx.Account ?? $profile.id
                    $account = @{id=('sdk:' + $profile.id);login=$label}
                    $sdkAccount = $ctx.Account; $sdkTenant = $ctx.TenantId
                    $sdkScope = $scope
                    $accessMode = $mode
                    Emit @{id=$request.id;result=@{account=$account;mode=$accessMode}}
                }
                'request' {
                    Require-Account $request.params.account
                    $p = $request.params
                    $uri = [Uri]$p.url
                    $path = $uri.AbsolutePath
                    if ($uri.Scheme -ne 'https' -or $uri.Host -ne 'graph.microsoft.com' -or $uri.Port -ne 443 -or $uri.UserInfo -or $uri.Fragment) { throw 'INVALID_REQUEST' }
                    $readPath = '^/v1\.0/me(?:/messages(?:/[^/]+)?|/mailFolders(?:/[^/]+/(?:messages|childFolders))?)$'
                    $valid = ($p.method -ceq 'GET' -and $path -cmatch $readPath -and !$p.ContainsKey('body')) -or
                        ($p.method -ceq 'PATCH' -and $path -cmatch '^/v1\.0/me/messages/[^/]+$') -or
                        ($p.method -ceq 'POST' -and $path -cmatch '^/v1\.0/me/(?:sendMail|messages(?:/[^/]+/move)?)$')
                    if (!$valid -or $uri.OriginalString -match '/(?:\.|%2e){1,2}(?:/|\?|$)') { throw 'INVALID_REQUEST' }
                    if ($p.method -ne 'GET' -and $accessMode -ne 'full') { throw 'WRITE_SCOPE_REQUIRED' }
                    $invoke = @{Method=$p.method;Uri=$uri;OutputType='HttpResponseMessage';SkipHttpErrorCheck=$true;Headers=@{Accept='application/json';Prefer='outlook.body-content-type="text"'}}
                    if ($p.ContainsKey('body')) {
                        if ($p.body -isnot [string] -or [Text.Encoding]::UTF8.GetByteCount($p.body) -gt 250000) { throw 'INVALID_REQUEST' }
                        $invoke.Body = $p.body
                        $invoke.ContentType = 'application/json'
                    }
                    $stage = 'graph_request'
                    $response = Invoke-MgGraphRequest @invoke
                    try {
                        $body = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        $truncated = [Text.Encoding]::UTF8.GetByteCount($body) -gt 500000
                        Emit @{id=$request.id;result=@{ok=$true;status=[int]$response.StatusCode;body=$(if($truncated){''}else{$body});truncated=$truncated}}
                    } finally { $response.Dispose() }
                }
                'disconnect' {
                    if (Get-MgContext) { Disconnect-MgGraph | Out-Null }
                    $account=$null; $accessMode=$null; $sdkAccount=$null; $sdkTenant=$null; $sdkScope='Process'
                    Emit @{id=$request.id;result=@{disconnected=$true}}
                }
                default { throw 'INVALID_REQUEST' }
            }
        } catch {
            $known = @('INVALID_REQUEST','ACCOUNT_NOT_CONNECTED','ALREADY_CONNECTED','INVALID_AUTH_CONTEXT','PROFILE_ID_MISSING','ACCOUNT_MISMATCH','WRITE_SCOPE_REQUIRED')
            $code = if ($_.Exception.Message -cin $known) { $_.Exception.Message } else { 'SDK_REQUEST_FAILED' }
            # Raw SDK errors can contain HTTP headers or mailbox data: never forward them.
            $failure=$_.Exception; $types=@(); $messages=@()
            while ($failure -and $types.Count -lt 8) { $types+=$failure.GetType().Name; $messages+=$failure.Message; $failure=$failure.InnerException }
            $description=$messages -join ' '
            $microsoftCodes=@([regex]::Matches($description,'AADSTS\d+') | ForEach-Object {$_.Value} | Select-Object -Unique -First 5)
            if ($code -eq 'SDK_REQUEST_FAILED' -and $description -match '(?i)timed?\s*out|timeout|expired_token') { $code='SDK_TIMEOUT' }
            $contextCheck=$null
            if ($code -eq 'INVALID_AUTH_CONTEXT') {
                $current=Get-MgContext
                $contextCheck=@{present=($null -ne $current);auth_type=[string]$current.AuthType;environment=[string]$current.Environment;context_scope=[string]$current.ContextScope;has_account_claim=[bool]$current.Account}
            }
            Emit @{id=$request.id;error=@{code=$code;stage=$stage;error_types=$types;microsoft_codes=$microsoftCodes;context_check=$contextCheck}}
        }
    }
} finally {
    # CurrentUser uses the SDK's OS-protected cache. Only explicit disconnect clears it.
    if ((Get-MgContext) -and $sdkScope -eq 'Process') { Disconnect-MgGraph | Out-Null }
}
