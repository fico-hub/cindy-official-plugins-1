'use strict';
(async () => {
  const $ = id => document.getElementById(id);
  let strings = {}, entries = [], selected = 'global', config = {}, busy = false, loaded = false;
  let sdkState=null,initialized=false;
  const channel=new BroadcastChannel('outlook-sdk-settings'),pending=new Map();
  channel.onmessage=event=>{
    const m=event.data;if(m?.type!=='settings-response')return;
    const p=pending.get(m.reqId);if(!p)return;clearTimeout(p.timer);clearInterval(p.retry);pending.delete(m.reqId);
    if(m.ok)p.resolve(m.result);else p.reject(new Error(m.message||t('connectFailed')));
  };
  function sdk(action,payload={}){
    const reqId=crypto.randomUUID();
    return new Promise((resolve,reject)=>{
      const p={resolve,reject,retry:null};
      p.timer=setTimeout(()=>{clearInterval(p.retry);pending.delete(reqId);reject(new Error(t('TIMEOUT')));},action==='status'?12000:155000);
      pending.set(reqId,p);
      const post=()=>channel.postMessage({type:'settings-request',reqId,action,payload});
      // Settings do not automatically start the on-demand plugin main.js.
      void fetch('/wake').then(()=>{
        if(!pending.has(reqId))return;
        post();p.retry=setInterval(post,400);
      }).catch(()=>{clearTimeout(p.timer);pending.delete(reqId);reject(new Error(t('loadFailed')));});
    });
  }
  const useSdk=()=>selected==='global'&&config.global_connection!=='direct';
  const clouds = {global:{key:'outlook_global',port:53687,portal:'https://portal.azure.com'},china:{key:'outlook_china',port:53688,portal:'https://portal.azure.cn'}};
  const t = key => strings[key] || key;
  async function request(path, init) {
    const r = await fetch(path, init);
    if (!r.ok) throw new Error(t('requestFailed')+' (HTTP '+r.status+')');
    return r.status === 204 ? null : r.json();
  }
  async function locale() {
    let lang = 'en';
    try { const c = await request('/app-context'); if (['en','zh-CN'].includes(c?.context?.locale)) lang = c.context.locale; } catch (_) { /* English fallback */ }
    try { strings = await request('/ui/'+lang+'.json'); }
    catch (_) { lang = 'en'; strings = await request('/ui/en.json'); }
    document.documentElement.lang = lang;
    document.querySelectorAll('[data-i18n]').forEach(el => { el.textContent = t(el.dataset.i18n); });
  }
  function status(message) { $('status').textContent = message; }
  function errorText(r) {
    const code = typeof r?.error === 'string' ? r.error : '';
    return t(code) === code ? t('connectFailed') : t(code);
  }
  function addButton(row,label,fn) {
    const b = document.createElement('button'); b.type = 'button'; b.textContent = label; b.disabled = busy;
    b.onclick = () => void perform(fn); row.appendChild(b);
  }
  function render() {
    const c = clouds[selected], e = entries.find(e => e.key === c.key);
    const isSdk=useSdk();
    $('connection-options').hidden=selected!=='global';
    $('connection-mode').value=config.global_connection==='direct'?'direct':'sdk';
    $('connection-mode').disabled=busy||!loaded;
    $('sdk-options').hidden=!isSdk;
    $('sdk-email').disabled=busy||!!config.sdk_account;
    $('sdk-mode').disabled=busy||!loaded;
    $('advanced').hidden=isSdk;
    $('cloud').value = selected;
    $('cloud').disabled = busy;
    $('redirect').textContent = 'http://127.0.0.1:'+c.port+'/callback';
    $('portal').href = c.portal;
    $('default-label').textContent = t('defaultService')+': '+t(config.default_cloud === 'china' ? 'china' : 'global');
    $('setup-note').textContent = !loaded ? t('loadFailed') : isSdk ? t(!sdkState?.supported?'sdkUnsupported':!sdkState?.runtime_ready?'sdkUnavailable':'sdkReady') : e?.clientConfigured ? t(e.clientCustom ? 'customReady' : 'builtInReady') : t('appMissing');
    $('advanced').open = loaded && !e?.clientConfigured;
    for (const id of ['default-cloud','save-client','reset-client','client-id']) $(id).disabled = busy || !loaded;
    $('connect').disabled = busy || !loaded || (isSdk?!sdkState?.supported||!sdkState?.runtime_ready:!e?.clientConfigured);
    $('reset-client').disabled = busy || !e?.clientCustom;
    $('accounts').textContent = '';
    if(isSdk){
      const a=sdkState?.account||config.sdk_account;
      if(a){
        const row=document.createElement('div');row.className='account';
        const name=document.createElement('span');name.className='identity';name.textContent=a.login;
        const tag=document.createElement('span');tag.className='tag';tag.textContent=t(sdkState?.connected?'connected':'savedLogin');
        name.appendChild(tag);row.appendChild(name);
        addButton(row,t('disconnect'),async()=>{await sdk('disconnect');status(t('disconnected'));$('sdk-email').value='';});
        $('accounts').appendChild(row);
      }else{const p=document.createElement('p');p.className='hint';p.textContent=t('noAccounts');$('accounts').appendChild(p);}
      return;
    }
    if (loaded && !e?.accounts?.length) { const p = document.createElement('p'); p.className='hint'; p.textContent=t('noAccounts'); $('accounts').appendChild(p); }
    for (const a of e?.accounts || []) {
      const row = document.createElement('div'); row.className='account';
      const name = document.createElement('span'); name.className='identity'; name.textContent=a.label || a.id;
      const tag = document.createElement('span'); tag.className='tag';
      tag.textContent = a.status !== 'connected' ? t('expired') : a.scopeStale ? t('scopeStale') : a.isDefault ? t('defaultAccount') : t('connected');
      name.appendChild(tag); row.appendChild(name);
      if (a.status !== 'connected' || a.scopeStale) addButton(row,t('reconnect'),connect);
      if (!a.isDefault && a.status === 'connected') addButton(row,t('setDefault'),async () => {
        await request('/oauth/'+c.key+'/default',{method:'POST',body:JSON.stringify({accountId:a.id})}); status(t('saved'));
      });
      addButton(row,t('disconnect'),async () => {
        await request('/oauth/'+c.key+'/accounts/'+encodeURIComponent(a.id),{method:'DELETE'}); status(t('disconnected'));
      });
      $('accounts').appendChild(row);
    }
  }
  async function load() {
    try {
      const values = await Promise.all([request('/oauth'),request('/kv')]);
      if (!Array.isArray(values[0])) throw new Error(t('loadFailed'));
      entries=values[0]; config=values[1];
      if(!initialized){selected=config.default_cloud==='china'?'china':'global';initialized=true;}
      if(useSdk()){
        try{sdkState=await sdk('status');}
        catch{sdkState={supported:true,runtime_ready:false};status(t('sdkUnavailable'));}
      }
      if(config.sdk_account)$('sdk-email').value=config.sdk_account.login;
      $('sdk-mode').value=config.sdk_mode==='read'?'read':'full';
      loaded=true;
    } catch (_) { loaded=false; entries=[]; status(t('loadFailed')); }
    render();
  }
  async function perform(fn) {
    if (busy) return;
    busy=true; render();
    try { await fn(); } catch (e) { status(e.message || t('requestFailed')); }
    finally { await load(); busy=false; render(); }
  }
  async function connect() {
    status(t('connecting'));
    if(useSdk()){
      const s=await sdk('connect',{loginHint:$('sdk-email').value.trim(),mode:$('sdk-mode').value});
      status(t('connected')+': '+(s.account?.login||''));return;
    }
    const r = await request('/oauth/'+clouds[selected].key+'/connect',{method:'POST'});
    if (!r.ok) { status(errorText(r)); return; }
    status(t('connected')+': '+(r.account?.label || ''));
  }
  $('cloud').onchange = () => { selected=$('cloud').value; $('client-id').value=''; status('');void perform(async()=>{}); };
  $('connection-mode').onchange=()=>{
    const chosen=$('connection-mode').value;
    void perform(async()=>{const current=await request('/kv');await request('/kv',{method:'PUT',body:JSON.stringify({...current,global_connection:chosen})});});
  };
  $('connect').onclick = () => void perform(connect);
  $('default-cloud').onclick = () => void perform(async () => {
    // Read-modify-write: preserve future unrelated preferences.
    const current = await request('/kv');
    await request('/kv',{method:'PUT',body:JSON.stringify({...current,default_cloud:selected})}); status(t('saved'));
  });
  $('save-client').onclick = () => void perform(async () => {
    const clientId=$('client-id').value.trim();
    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(clientId)) throw new Error(t('invalidClient'));
    await request('/oauth/'+clouds[selected].key+'/client',{method:'PUT',body:JSON.stringify({clientId})});
    $('client-id').value=''; status(t('clientSaved'));
  });
  $('reset-client').onclick = () => void perform(async () => {
    await request('/oauth/'+clouds[selected].key+'/client',{method:'DELETE'}); $('client-id').value=''; status(t('clientCleared'));
  });
  await locale(); await load();
  render();
})();
