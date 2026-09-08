import {GraphSession,createMailAdapter} from './graph-session.mjs';
import {fileURLToPath} from 'node:url';
import fs from 'node:fs';
const local=p=>fileURLToPath(new URL(p,import.meta.url));
const fail=(code,execution='not_executed')=>Object.assign(new Error(code),{code,execution});

export class OutlookService {
  constructor({createSession,platform=process.platform,arch=process.arch,sourcePath=local('../main.js')}={}) {
    this.supported=platform==='darwin'&&arch==='arm64';
    this.createSession=createSession||(()=>new GraphSession({pwsh:local('./runtime/pwsh/pwsh'),
      modulePath:local('./runtime/modules/Microsoft.Graph.Authentication/2.39.0/Microsoft.Graph.Authentication.psd1'),
      scriptPath:local('./graph-session.ps1')}));
    this.runtimeReady=!!createSession||fs.existsSync(local('./runtime/pwsh/pwsh'));
    this.session=null;this.account=null;this.mode=null;this.busy=false;
    this.adapter=createMailAdapter({status:()=>this.session.status(),request:r=>this.session.request(r)},{sourcePath});
  }
  status(){return {supported:this.supported,runtime_ready:this.runtimeReady,connected:!!this.account,
    account:this.account,mode:this.mode,busy:this.busy,persistence:'microsoft_sdk'};}
  async prepare(){
    if(!this.supported)throw fail('SDK_PLATFORM_UNSUPPORTED');
    if(!this.runtimeReady)throw fail('SDK_RUNTIME_UNAVAILABLE');
    if(!this.session){
      const s=this.createSession();this.session=s;
      s.on('closed',()=>{if(this.session===s){this.session=null;this.account=null;this.mode=null;}});
      await s.status();
    }
  }
  saved(value){
    if(!value||typeof value.id!=='string'||!value.id.startsWith('sdk:')||value.id.length>2048||
      typeof value.login!=='string'||!value.login.trim()||value.login.length>320||/[\x00-\x1f]/.test(value.login))throw fail('ACCOUNT_NOT_CONNECTED');
    return {id:value.id,login:value.login};
  }
  async connect(p={}){
    if(this.busy)throw fail('SDK_BUSY');
    if(p.mode!==undefined&&!['read','full'].includes(p.mode))throw fail('INVALID_ARGUMENT');
    if(p.loginHint!==undefined&&(typeof p.loginHint!=='string'||p.loginHint.length>320||/[\x00-\x1f]/.test(p.loginHint)))throw fail('INVALID_ARGUMENT');
    const saved=p.saved?this.saved(p.saved):null;
    this.busy=true;
    try{
      if(this.session)this.session.close();
      await this.prepare();
      const r=await this.session.connect({mode:p.mode||'full',flow:'browser',contextScope:'CurrentUser',
        loginHint:saved?.login||p.loginHint||undefined,expectedId:saved?.id});
      this.account=r.account;this.mode=r.mode;
      return this.status();
    }finally{this.busy=false;}
  }
  async mail(p={}){
    this.adapter.validate(p.args);
    if(p.args.cloud==='china')throw fail('SDK_CHINA_UNSUPPORTED');
    if(this.busy)throw fail('SDK_BUSY');
    // Validate the intended identity before any automatic cache restoration.
    const intended=this.account||this.saved(p.saved);
    if(p.args.account&&p.args.account!==intended.id)throw fail('ACCOUNT_NOT_FOUND');
    if(!this.account)await this.connect({saved:intended,mode:p.mode==='read'?'read':'full'});
    this.busy=true;
    try{return await this.adapter.run(p.args);}finally{this.busy=false;}
  }
  async disconnect(){
    if(this.busy)throw fail('SDK_BUSY');
    this.busy=true;
    try{
      if(this.session&&this.account)await this.session.call('disconnect');
      return {disconnected:true};
    }finally{this.session?.close();this.account=null;this.mode=null;this.busy=false;}
  }
  close(){this.session?.close();}
}

export function safeError(error){
  const known=new Set(['INVALID_ARGUMENT','INVALID_REQUEST','ACCOUNT_NOT_CONNECTED','ACCOUNT_NOT_FOUND','ACCOUNT_MISMATCH',
    'SDK_PLATFORM_UNSUPPORTED','SDK_RUNTIME_UNAVAILABLE','SDK_CHINA_UNSUPPORTED','SDK_BUSY','SDK_TIMEOUT','SDK_REQUEST_FAILED',
    'SDK_PROCESS_CLOSED','SDK_PROTOCOL_ERROR','SDK_RESPONSE_TOO_LARGE','WRITE_SCOPE_REQUIRED','GRAPH_REQUEST_FAILED','INVALID_RESPONSE']);
  return {ok:false,code:known.has(error.code)?error.code:'SDK_REQUEST_FAILED',
    execution_status:['not_executed','executed','unknown'].includes(error.execution)?error.execution:'unknown',
    microsoft_codes:Array.isArray(error.microsoft_codes)?error.microsoft_codes.filter(x=>/^AADSTS\d+$/.test(x)).slice(0,5):[]};
}
