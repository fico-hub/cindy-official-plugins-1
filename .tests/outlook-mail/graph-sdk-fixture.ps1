# Offline transport fixture. Functions shadow SDK cmdlets; no login or network.
Import-Module (Join-Path $PSScriptRoot '../../outlook-mail/node/runtime/modules/Microsoft.Graph.Authentication/2.39.0/Microsoft.Graph.Authentication.psd1')
$global:fixtureContext=$null
function Connect-MgGraph {
    param($Environment,$TenantId,$Scopes,$ContextScope,$NoWelcome,$LoginHint,$UseDeviceAuthentication)
    $global:fixtureContext=[pscustomobject]@{AuthType='Delegated';Environment=$Environment;ContextScope=$ContextScope;Account=$null;TenantId='fixture-tenant'}
}
function Get-MgContext { $global:fixtureContext }
function Disconnect-MgGraph {
    [Console]::WriteLine('Fixture SDK diagnostic: this must not enter the JSON protocol')
    Write-Warning 'Fixture SDK warning: also not a protocol message'
    $global:fixtureContext=$null
}
function Invoke-MgGraphRequest {
    param($Method,$Uri,$OutputType,$SkipHttpErrorCheck,$Headers,$Body,$ContentType)
    if ($OutputType -eq 'Json' -and $Method -eq 'GET' -and ([uri]$Uri).AbsolutePath -eq '/v1.0/me') {
        return '{"id":"fixture-account","mail":"fixture@example.test"}'
    }
    if ($OutputType -ne 'HttpResponseMessage') { throw 'UNEXPECTED_FIXTURE_CALL' }
    $response=[System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]::OK)
    $response.Content=[System.Net.Http.StringContent]::new('{"fixture_network_called":true,"value":[]}')
    return $response
}
foreach($name in @('Connect-MgGraph','Get-MgContext','Disconnect-MgGraph','Invoke-MgGraphRequest')) {
    if ((Get-Command $name).CommandType -ne 'Function') { throw 'FIXTURE_NOT_ISOLATED' }
}
. (Join-Path $PSScriptRoot '../../outlook-mail/node/graph-session.ps1') -ModulePath (Join-Path $PSScriptRoot '../../outlook-mail/node/runtime/modules/Microsoft.Graph.Authentication/2.39.0/Microsoft.Graph.Authentication.psd1')
