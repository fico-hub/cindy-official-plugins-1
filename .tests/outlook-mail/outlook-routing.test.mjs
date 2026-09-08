import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const source=fs.readFileSync(new URL('../../outlook-mail/main.js',import.meta.url),'utf8');
function fixture(config={},nodeValue={ok:true,value:{items:[]}}){
  const native=[],node=[],results=[],broadcasts=[];let listener,channel;
  const oauth=[{key:'outlook_global',clientConfigured:true,accounts:[{id:'g1',label:'global',status:'connected',isDefault:true}]},
    {key:'outlook_china',clientConfigured:true,accounts:[{id:'c1',label:'china',status:'connected',isDefault:true}]}];
  const context=vm.createContext({URL,URLSearchParams,TextEncoder,TextDecoder,btoa,atob,
    BroadcastChannel:class{constructor(){channel=this;}postMessage(m){broadcasts.push(m);}},
    fetch:async(path,init)=>{if(path==='/kv'&&init?.method==='PUT')config=JSON.parse(init.body);return {ok:true,json:async()=>path==='/oauth'?oauth:config};},
    cindy:{onHostMessage:f=>listener=f,send:r=>results.push(r),
      fetch:async r=>{native.push(r);return {ok:true,status:200,body:'{"value":[]}'};},
      node:{request:async r=>{node.push(r);return {ok:true,result:nodeValue};}}}});
  vm.runInContext(source,context);
  return {native,node,broadcasts,get config(){return config;},
    call:async(args,tool='outlook_mail')=>{await listener({type:'tool-call',tool,args,callId:'routing-test'});return results.at(-1);},
    settings:async m=>channel.onmessage({data:m})};
}
test('fresh global settings route mail to the bundled SDK without native OAuth',async()=>{
  const h=fixture({sdk_account:{id:'sdk:one',login:'user@example.test'},sdk_mode:'read'});
  assert((await h.call({action:'search'})).ok);assert.equal(h.native.length,0);
  assert.equal(h.node[0].method,'mail/action');assert.equal(h.node[0].params.saved.id,'sdk:one');assert.equal(h.node[0].params.mode,'read');
});
test('explicit account wins over default cloud and selected login method',async()=>{
  const h=fixture({default_cloud:'china',global_connection:'direct'});
  assert((await h.call({action:'search',account:'sdk:one'})).ok);assert.equal(h.node.length,1);assert.equal(h.native.length,0);
  assert((await h.call({action:'search',account:'g1'})).ok);assert.equal(h.native[0].authAccount,'g1');
});
test('China and explicitly selected own-app mode preserve native OAuth routing',async()=>{
  const h=fixture();assert((await h.call({action:'search',cloud:'china'})).ok);assert.equal(h.node.length,0);assert.equal(h.native[0].authAccount,'c1');
  const direct=fixture({global_connection:'direct'});assert((await direct.call({action:'search'})).ok);assert.equal(direct.native[0].authAccount,'g1');assert.equal(direct.node.length,0);
});
test('invalid writes fail before launching SDK and uncertain results remain uncertain',async()=>{
  const h=fixture({}, {ok:false,code:'SDK_TIMEOUT',execution_status:'unknown'});
  const invalid=await h.call({action:'send',to:'not-an-email'});assert.equal(invalid.ok,false);assert.equal(h.node.length,0);
  const r=await h.call({action:'send',to:'reader@example.test',subject:'fixture',body_text:'never sent'});
  assert.equal(r.ok,false);assert.match(r.message,/execution_status=unknown/);assert.equal(h.node.length,1);
});
test('completed settings messages are replayed without reconnecting or repeating state writes',async()=>{
  const account={id:'sdk:one',login:'user@example.test'};
  const h=fixture({default_cloud:'china'},{ok:true,value:{account,mode:'read'}});
  const request={type:'settings-request',reqId:'once',action:'connect',payload:{mode:'read'}};
  await h.settings(request);await h.settings(request);
  assert.equal(h.node.length,1);assert.equal(h.broadcasts.length,2);
  assert.deepEqual(h.broadcasts[0],h.broadcasts[1]);assert.equal(h.config.default_cloud,'china');assert.equal(h.config.sdk_account.id,account.id);
});
