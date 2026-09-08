const nativeTest=(name,fn)=>test(name,{skip:process.platform!=='darwin'||process.arch!=='arm64'},fn);
import test from 'node:test';
import assert from 'node:assert/strict';
import {EventEmitter} from 'node:events';
import {OutlookService} from '../../outlook-mail/node/service.mjs';

function fixture(){
  const events=[];
  class Session extends EventEmitter {
    constructor(){super();this.account=null;this.mode=null;}
    async status(){return {account:this.account,mode:this.mode};}
    async connect(p){events.push({connect:p});this.account={id:'sdk:test',login:'user@example.test'};this.mode=p.mode;return this.status();}
    async request(p){events.push({request:p});return {ok:true,status:200,body:'{"value":[]}'};}
    async call(method){events.push({method});this.account=null;return {disconnected:true};}
    close(){this.emit('closed');}
  }
  return {events,service:new OutlookService({createSession:()=>new Session(),platform:'darwin',arch:'arm64'})};
}
const saved={id:'sdk:test',login:'user@example.test'};
nativeTest('packaged runtime can start without credentials',async()=>{
  const service=new OutlookService();
  try{await service.prepare();assert.equal(service.status().runtime_ready,true);assert.equal(service.status().connected,false);}
  finally{service.close();}
});
test('saved login restores the exact mailbox before executing the mail action',async()=>{
  const {events,service}=fixture();
  try{
    const r=await service.mail({saved,mode:'read',args:{action:'search'}});
    assert.equal(r.account,saved.id);assert.equal(events[0].connect.expectedId,saved.id);
    assert.equal(events[0].connect.loginHint,saved.login);assert.equal(events[0].connect.contextScope,'CurrentUser');
    assert.equal(events[1].request.authAccount,saved.id);
  }finally{service.close();}
});
test('malformed actions and mismatched accounts never launch restoration',async()=>{
  const {events,service}=fixture();
  try{
    await assert.rejects(service.mail({saved,args:{action:'send',to:'not-an-email',subject:'x',body_text:'x'}}));
    await assert.rejects(service.mail({saved,args:{action:'search',account:'sdk:another'}}),/ACCOUNT_NOT_FOUND/);
    await assert.rejects(service.mail({saved,args:{action:'search',cloud:'china'}}),/SDK_CHINA_UNSUPPORTED/);
    assert.equal(events.length,0);
  }finally{service.close();}
});
test('restored read-only login cannot send and closing the process clears live status',async()=>{
  const {events,service}=fixture();
  try{
    await assert.rejects(service.mail({saved,mode:'read',args:{action:'send',to:'user@example.test',subject:'fixture',body_text:'never sent'}}),/WRITE_SCOPE_REQUIRED/);
    assert.equal(events.filter(e=>e.request).length,0);assert.equal(service.status().connected,true);
    service.close();assert.equal(service.status().connected,false);
  }finally{service.close();}
});
test('explicit disconnect clears active SDK session and does not launch a missing session',async()=>{
  const {events,service}=fixture();
  await service.disconnect();assert.equal(events.length,0);
  await service.connect({mode:'read'});await service.disconnect();
  assert.equal(events.at(-1).method,'disconnect');assert.equal(service.status().connected,false);
});
test('unsupported devices are rejected before launching the bundled binary',async()=>{
  const s=new OutlookService({platform:'win32',arch:'x64'});
  assert.equal(s.status().supported,false);await assert.rejects(s.connect(),/SDK_PLATFORM_UNSUPPORTED/);
});
