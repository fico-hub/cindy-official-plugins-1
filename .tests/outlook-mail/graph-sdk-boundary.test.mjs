const nativeTest=(name,fn)=>test(name,{skip:process.platform!=='darwin'||process.arch!=='arm64'},fn);
import test from 'node:test';
import assert from 'node:assert/strict';
import {execFileSync} from 'node:child_process';
import {fileURLToPath} from 'node:url';

const pwsh=fileURLToPath(new URL('../../outlook-mail/node/runtime/pwsh/pwsh',import.meta.url));
const fixture=fileURLToPath(new URL('./graph-sdk-fixture.ps1',import.meta.url));
const connect={method:'connect',params:{audience:'common',mode:'full',flow:'browser',contextScope:'Process'}};
const request={account:'sdk:fixture-account',url:'https://graph.microsoft.com/v1.0/me/messages',method:'GET'};
function run(commands){
  const input=commands.map((x,i)=>JSON.stringify({id:i+1,...x})).join('\n')+'\n';
  return execFileSync(pwsh,['-NoLogo','-NoProfile','-File',fixture],{input,encoding:'utf8',timeout:10000})
    .trim().split(/\r?\n/).map(JSON.parse);
}
nativeTest('PowerShell accepts a real profile id when the SDK has no account-name claim',()=>{
  const r=run([connect,{method:'request',params:request}]);
  assert.equal(r[0].result.account.id,'sdk:fixture-account');
  assert.equal(r[0].result.account.login,'fixture@example.test');
  assert.equal(r[1].result.status,200);
});
nativeTest('PowerShell prevents restoration to a different mailbox before exposing it',()=>{
  const r=run([{...connect,params:{...connect.params,contextScope:'CurrentUser',loginHint:'fixture@example.test',expectedId:'sdk:another-account'}}]);
  assert.equal(r[0].error.code,'ACCOUNT_MISMATCH');
  assert.equal(r[0].result,undefined);
});
nativeTest('PowerShell rejects foreign endpoints, traversal, credentials and non-mail routes',()=>{
  const urls=['https://example.test/v1.0/me/messages','http://graph.microsoft.com/v1.0/me/messages',
    'https://graph.microsoft.com:444/v1.0/me/messages','https://user@graph.microsoft.com/v1.0/me/messages',
    'https://graph.microsoft.com/v1.0/me/messages#fragment','https://graph.microsoft.com/v1.0/users',
    'https://graph.microsoft.com/v1.0/me/messages/%2e%2e/messages','https://graph.microsoft.com/v1.0/me/messages/../messages'];
  const r=run([connect,...urls.map(url=>({method:'request',params:{...request,url}}))]);
  for(const x of r.slice(1))assert.equal(x.error.code,'INVALID_REQUEST');
});
nativeTest('PowerShell itself enforces read-only consent and exact mailbox binding',()=>{
  const r=run([{...connect,params:{...connect.params,mode:'read'}},
    {method:'request',params:{...request,method:'POST',url:'https://graph.microsoft.com/v1.0/me/sendMail',body:'{}'}},
    {method:'request',params:{...request,account:'sdk:another-account'}},
    {method:'disconnect'},{method:'request',params:request}]);
  assert.equal(r[1].error.code,'WRITE_SCOPE_REQUIRED');assert.equal(r[2].error.code,'ACCOUNT_NOT_CONNECTED');
  assert.equal(r[3].result.disconnected,true);assert.equal(r[4].error.code,'ACCOUNT_NOT_CONNECTED');
});
nativeTest('PowerShell accepts legitimate nested mailbox paths and rejects unrequested methods',()=>{
  const allowed=[['GET','/mailFolders'],['GET','/mailFolders/inbox/messages'],['GET','/mailFolders/id/childFolders'],
    ['POST','/messages'],['POST','/sendMail'],['POST','/messages/id/move'],['PATCH','/messages/id']];
  const commands=allowed.map(([method,path])=>({method:'request',params:{...request,method,url:'https://graph.microsoft.com/v1.0/me'+path,...(method==='GET'?{}:{body:'{}'})}}));
  const r=run([connect,...commands,{method:'request',params:{...request,method:'DELETE'}}]);
  for(const x of r.slice(1,-1))assert.equal(x.result.status,200);
  assert.equal(r.at(-1).error.code,'INVALID_REQUEST');
});
