import {fileURLToPath} from 'node:url';
const nativeTest=(name,fn)=>test(name,{skip:process.platform!=='darwin'||process.arch!=='arm64'},fn);
import test from 'node:test';
import assert from 'node:assert/strict';
import {GraphSession,createMailAdapter} from '../../outlook-mail/node/graph-session.mjs';

nativeTest('real PowerShell and official Graph module start through JSON protocol with no connected account', async () => {
  const session=new GraphSession({pwsh:fileURLToPath(new URL('../../outlook-mail/node/runtime/pwsh/pwsh',import.meta.url)),modulePath:fileURLToPath(new URL('../../outlook-mail/node/runtime/modules/Microsoft.Graph.Authentication/2.39.0/Microsoft.Graph.Authentication.psd1',import.meta.url))});
  try {
    const status=await session.status();
    assert.equal(status.ready,true); assert.equal(status.account,null);
    assert.equal(status.cloud,'global'); assert.equal(status.token_persistence,'process_only');
    await assert.rejects(session.request({url:'https://graph.microsoft.com/v1.0/me/messages',method:'GET',authAccount:'missing'}),/ACCOUNT_NOT_CONNECTED/);
    assert.equal((await session.status()).account,null);
  } finally {session.close();}
});
nativeTest('real SDK transport rejects unknown methods and unsupported auth options before launching login', async () => {
  const session=new GraphSession({pwsh:fileURLToPath(new URL('../../outlook-mail/node/runtime/pwsh/pwsh',import.meta.url)),modulePath:fileURLToPath(new URL('../../outlook-mail/node/runtime/modules/Microsoft.Graph.Authentication/2.39.0/Microsoft.Graph.Authentication.psd1',import.meta.url))}); let challenges=0; session.on('login',()=>challenges++);
  try {
    await assert.rejects(session.call('arbitrary-code',{command:'Write-Output secret'}),/INVALID_REQUEST/);
    await assert.rejects(session.connect({audience:'china',mode:'full'}),/INVALID_REQUEST/);
    await assert.rejects(session.connect({audience:'common',mode:'arbitrary'}),/INVALID_REQUEST/);
    await assert.rejects(session.connect({flow:'arbitrary'}),/INVALID_REQUEST/);
    assert.equal(challenges,0);
    assert.equal((await session.status()).account,null);
  } finally {session.close();}
});
nativeTest('real transport rejects concurrent requests, shuts down and refuses calls after close', async () => {
  const session=new GraphSession({pwsh:fileURLToPath(new URL('../../outlook-mail/node/runtime/pwsh/pwsh',import.meta.url)),modulePath:fileURLToPath(new URL('../../outlook-mail/node/runtime/modules/Microsoft.Graph.Authentication/2.39.0/Microsoft.Graph.Authentication.psd1',import.meta.url))});
  try {
    const first=session.status();
    await assert.rejects(session.status(),/SDK_BUSY/);
    assert.equal((await first).ready,true);
    session.close();
    await assert.rejects(session.status(),/SDK_PROCESS_CLOSED/);
  } finally {session.close();}
});
test('missing PowerShell returns a controlled failure without a hanging request',async()=>{
  const session=new GraphSession({pwsh:'/nonexistent-cindy-outlook-runtime/pwsh'});
  try {await assert.rejects(session.status(),/SDK_RUNTIME_UNAVAILABLE|SDK_PROCESS_CLOSED/);} finally {session.close();}
});
function fake({mode='full',account='sdk:one',response={ok:true,status:200,body:'{"value":[]}'}}={}) {
  const state={mode,account:{id:account,login:'test@example.test'}};
  const calls=[];
  const session={status:async()=>state,request:async r=>{calls.push(r);return response;}};
  return {state,calls,adapter:createMailAdapter(session,{sourcePath:new URL('../../outlook-mail/main.js',import.meta.url)})};
}
test('existing mail core binds SDK account and accepts mailbox search through transport',async()=>{
  const f=fake(); const result=await f.adapter.run({action:'search',subject:'invoice'});
  assert.equal(result.account,'sdk:one'); assert.equal(result.returned_count,0);
  assert.equal(f.calls[0].authAccount,'sdk:one');assert.equal(f.calls[0].method,'GET');
  assert.equal(new URL(f.calls[0].url).searchParams.get('$filter'),"contains(subject,'invoice')");
  assert.equal(f.calls[0].headers.Authorization,undefined);
});
test('read-only SDK consent prevents every mutation without sending a request',async()=>{
  const f=fake({mode:'read'});
  const mail={to:'test@example.test',subject:'Fixture only',body_text:'Never sent'};
  for (const args of [{action:'send',...mail},{action:'draft',...mail},{action:'mark_read',message_id:'m'},
    {action:'mark_unread',message_id:'m'},{action:'move',message_id:'m',target_folder:'inbox'}]) {
    await assert.rejects(f.adapter.run(args),e=>e.code==='WRITE_SCOPE_REQUIRED'&&e.execution==='not_executed');
  }
  assert.equal(f.calls.length,0);
  await f.adapter.run({action:'search'}); assert.equal(f.calls.length,1);
});
test('SDK cannot use another account or the China cloud through the existing mail core',async()=>{
  const f=fake();
  await assert.rejects(f.adapter.run({action:'search',account:'sdk:someone-else'}),e=>e.code==='ACCOUNT_NOT_FOUND');
  await assert.rejects(f.adapter.run({action:'search',cloud:'china'}),e=>e.code==='CLIENT_NOT_CONFIGURED');
  assert.equal(f.calls.length,0);
});
test('SDK transport preserves send acceptance and ambiguous failure without retrying',async()=>{
  const args={action:'send',to:'test@example.test',subject:'Fixture',body_text:'No real mail'};
  const accepted=fake({response:{ok:true,status:202,body:''}});
  const result=await accepted.adapter.run(args);
  assert.equal(result.accepted,true);assert.equal(result.delivery_confirmed,false);assert.equal(accepted.calls.length,1);
  const failed=fake({response:{ok:true,status:503,body:'{}'}});
  await assert.rejects(failed.adapter.run(args),e=>e.execution==='unknown');assert.equal(failed.calls.length,1);
});
