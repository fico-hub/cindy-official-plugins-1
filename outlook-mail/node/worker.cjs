'use strict';
const readline=require('node:readline');
let activeService;
const ready=import('./service.mjs').then(m=>({service:activeService=new m.OutlookService(),safeError:m.safeError}));
const reply=x=>process.stdout.write(JSON.stringify(x)+'\n');
readline.createInterface({input:process.stdin}).on('line',async line=>{
  let r;
  try{r=JSON.parse(line);}catch{return;}
  if(!r||r.jsonrpc!=='2.0'||r.id===undefined)return;
  let heartbeat;
  try{
    const {service,safeError}=await ready;
    const methods={'account/status':()=>service.status(),'account/connect':()=>service.connect(r.params),
      'account/disconnect':()=>service.disconnect(),'mail/action':()=>service.mail(r.params)};
    if(!Object.hasOwn(methods,r.method)){reply({jsonrpc:'2.0',id:r.id,error:{code:-32601,message:'Method not found'}});return;}
    if(['account/connect','mail/action'].includes(r.method))heartbeat=setInterval(()=>reply({jsonrpc:'2.0',method:'progress',params:{stage:'working'}}),15000);
    try{reply({jsonrpc:'2.0',id:r.id,result:{ok:true,value:await methods[r.method]()}});}
    catch(e){reply({jsonrpc:'2.0',id:r.id,result:safeError(e)});}
  }catch{reply({jsonrpc:'2.0',id:r.id,result:{ok:false,code:'SDK_RUNTIME_UNAVAILABLE',execution_status:'not_executed'}});}
  finally{clearInterval(heartbeat);}
}).on('close',()=>{ready.then(({service})=>service.close());});
for(const sig of ['SIGTERM','SIGINT'])process.once(sig,()=>{ready.then(({service})=>{service.close();process.exit(0);});});
process.once('exit',()=>activeService?.close());
