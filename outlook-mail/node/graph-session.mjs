import {spawn} from 'node:child_process';
import {EventEmitter} from 'node:events';
import {fileURLToPath} from 'node:url';
import fs from 'node:fs';
import vm from 'node:vm';

// Development adapter only. Runtime dependencies are staged by the developer.
// Production packaging must bundle them; never download executables on first use.
export class GraphSession extends EventEmitter {
  #child; #pending = new Map(); #seq = 0; #buffer = ''; #closed = false;
  constructor({pwsh = fileURLToPath(new URL('../.outlook-runtime/pwsh/pwsh', import.meta.url)),
    modulePath = fileURLToPath(new URL('../.outlook-runtime/modules/Microsoft.Graph.Authentication/2.39.0/Microsoft.Graph.Authentication.psd1', import.meta.url)),
    scriptPath = fileURLToPath(new URL('./graph-session.ps1',import.meta.url))} = {}) {
    super();
    this.#child = spawn(pwsh, ['-NoLogo','-NoProfile','-File',scriptPath,'-ModulePath',modulePath], {stdio:['pipe','pipe','pipe']});
    this.#child.stdout.setEncoding('utf8');
    this.#child.stdout.on('data', chunk => {
      this.#buffer += chunk;
      if (Buffer.byteLength(this.#buffer) > 1_000_000) { this.close('SDK_RESPONSE_TOO_LARGE'); return; }
      let end;
      while ((end=this.#buffer.indexOf('\n')) !== -1) {
        const line=this.#buffer.slice(0,end).trim(); this.#buffer=this.#buffer.slice(end+1);
        if (!line) continue;
        let response;
        try { response=JSON.parse(line); } catch { this.close('SDK_PROTOCOL_ERROR'); return; }
        if (response.event === 'login_required') {
          if (response.verification_uri === 'https://login.microsoft.com/device' && /^[A-Z0-9]{9}$/.test(response.user_code)) {
            this.emit('login', {url:response.verification_uri,code:response.user_code,wait_seconds:120});
          }
          continue;
        }
        const pending=this.#pending.get(response.id);
        if (!pending) continue;
        clearTimeout(pending.timer); this.#pending.delete(response.id);
        if (response.error) pending.reject(Object.assign(new Error(response.error.code),{code:response.error.code,stage:response.error.stage,
          error_types:response.error.error_types,microsoft_codes:response.error.microsoft_codes,context_check:response.error.context_check}));
        else pending.resolve(response.result);
      }
    });
    // Drain but do not forward raw SDK diagnostic text or credentials.
    this.#child.stderr.on('data', () => {});
    this.#child.stdin.on('error', () => this.close('SDK_PROCESS_CLOSED'));
    this.#child.on('error', () => this.close('SDK_RUNTIME_UNAVAILABLE'));
    this.#child.on('close', () => this.close('SDK_PROCESS_CLOSED'));
  }
  call(method, params = {}, timeoutMs = 35000) {
    if (this.#closed) return Promise.reject(new Error('SDK_PROCESS_CLOSED'));
    if (this.#pending.size) return Promise.reject(new Error('SDK_BUSY'));
    const id=++this.#seq;
    const line=JSON.stringify({id,method,params})+'\n';
    if (Buffer.byteLength(line)>300000) return Promise.reject(new Error('INVALID_REQUEST'));
    return new Promise((resolve,reject) => {
      const timer=setTimeout(()=>this.close('SDK_TIMEOUT'),timeoutMs);
      this.#pending.set(id,{resolve,reject,timer});
      this.#child.stdin.write(line);
    });
  }
  connect({audience='common',mode='read',flow='browser',contextScope='Process',loginHint,expectedId}={}) {
    return this.call('connect',{audience,mode,flow,contextScope,loginHint,expectedId},135000);
  }
  status() { return this.call('status'); }
  request(request) {
    // Do not pass arbitrary caller headers, secrets, timeouts or PowerShell source.
    const {url,method,body,authAccount:account}=request;
    return this.call('request',{url,method,account,...(body===undefined?{}:{body})});
  }
  close(code='SDK_PROCESS_CLOSED') {
    if (this.#closed) return;
    this.#closed=true;
    for (const p of this.#pending.values()) { clearTimeout(p.timer); p.reject(Object.assign(new Error(code),{code})); }
    this.#pending.clear();
    this.#child.stdin.end(); this.#child.kill('SIGTERM');
    this.emit('closed',code);
  }
}

// Reuse the existing, tested validation, pagination and mail-result semantics.
export function createMailAdapter(session,{sourcePath=new URL('../outlook-mail/main.js',import.meta.url)}={}) {
  let snapshot=null;
  let busy=false;
  const context=vm.createContext({URL,URLSearchParams,TextEncoder,TextDecoder,btoa,atob,
    fetch: async path => ({ok:true,json:async()=>path==='/kv'?{default_cloud:'global'}:[
      {key:'outlook_global',clientConfigured:true,accounts:snapshot?.account?[{id:snapshot.account.id,label:snapshot.account.login,status:'connected',isDefault:true}]:[]},
      {key:'outlook_china',clientConfigured:false,accounts:[]},
    ]}),
  });
  vm.runInContext(fs.readFileSync(sourcePath,'utf8'),context);
  return {validate:args=>context.OutlookMail.validate(args),async run(args) {
    context.OutlookMail.validate(args);
    if (busy) throw Object.assign(new Error('SDK_BUSY'),{code:'SDK_BUSY',execution:'not_executed'});
    busy=true;
    try {
      snapshot=await session.status();
      if (!['search','read','list_folders'].includes(args.action) && snapshot.mode!=='full') {
        throw Object.assign(new Error('WRITE_SCOPE_REQUIRED'),{code:'WRITE_SCOPE_REQUIRED',execution:'not_executed'});
      }
      return await context.OutlookMail.run(args,'sdk-prototype',r=>session.request(r));
    } finally { busy=false; }
  }};
}
