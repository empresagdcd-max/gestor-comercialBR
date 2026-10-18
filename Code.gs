// GESTOR COMERCIAL - servidor (Google Apps Script). Pegar en Código.gs
const NEG={gp:'GERARDO POLIRUBRO',cd:'CANDU DESPENSA',ci:'CANDU INGENIO'},TZ='America/Argentina/Tucuman',MODS=['compras','boletas','stock'];
const P_=()=>PropertiesService.getScriptProperties();
const base_=()=>open_('BASE','01 - Base de Productos'),ops_=()=>open_('OPS','02 - Registros de Operaciones');
const sha_=s=>Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,P_().getProperty('SALT')+s).map(b=>('0'+(b&255).toString(16)).slice(-2)).join('');
const norm_=s=>String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim();

function open_(k,name){const p=P_();let id=p.getProperty(k);if(id)return SpreadsheetApp.openById(id);
  const f=DriveApp.getFoldersByName('Gestor Comercial'),fo=f.hasNext()?f.next():DriveApp.createFolder('Gestor Comercial'),it=fo.getFilesByName(name);let s;
  if(it.hasNext())s=SpreadsheetApp.open(it.next());else{s=SpreadsheetApp.create(name);DriveApp.getFileById(s.getId()).moveTo(fo);}
  p.setProperty(k,s.getId());return s;}
function tab_(ss,n,h,hide){let s=ss.getSheetByName(n);if(!s){s=ss.insertSheet(n);s.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold');s.setFrozenRows(1);if(hide)s.hideSheet();}return s;}
function txt_(s,row){s.getRange(s.getLastRow()+1,1,1,row.length).setNumberFormat('@').setValues([row]);}

// EJECUTAR UNA VEZ (se puede repetir sin borrar nada)
function setup(){
  if(!P_().getProperty('SALT'))P_().setProperty('SALT',Utilities.getUuid());
  const b=base_(),o=ops_();
  Object.keys(NEG).forEach(k=>{const K=k.toUpperCase();
    tab_(b,K+' Productos',['Codigo','Descripcion','Cód.Barras','CONTADO','Stock']);
    tab_(o,K+' Compras',['id','fecha','usuario','codigo','descripcion','cantidad','unidad','comentario']);
    tab_(o,K+' Boletas',['id','fecha','usuario','proveedor','factura','codigo','descripcion','cantidad','precio unitario','precio total']);
    tab_(o,K+' Stock',['id','fecha','usuario','codigo','descripcion','cantidad']);});
  tab_(o,'Productos pendientes',['fecha','negocio','codigo de barras','descripcion','estado','usuario']);
  tab_(o,'Documentos',['id','neg','tipo','estado','rev','actualizado','usuario','json'],1);
  const c=tab_(o,'Config',['clave','valor'],1),e=tab_(o,'Empleados',['id','nombre','hash','activo','creado'],1);
  if(c.getLastRow()<2)[['compras','3'],['boletas','4'],['stock','5'],['admin','1234']].forEach(r=>txt_(c,['pin_'+r[0],sha_(r[1])]));
  if(e.getLastRow()<2)Logger.log('CÓDIGO DEL ADMINISTRADOR (anotalo): '+addEmp_('Administrador'));
  Logger.log('Listo. Planillas conectadas.');
}
function nuevoAdmin(){Logger.log('NUEVO CÓDIGO DE ADMINISTRADOR: '+addEmp_('Administrador '+Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd')));}
function respaldo(){const fo=DriveApp.getFoldersByName('Gestor Comercial').next(),d=Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd');[base_(),ops_()].forEach(s=>DriveApp.getFileById(s.getId()).makeCopy(s.getName()+' RESPALDO '+d,fo));}

function addEmp_(n){const id=Utilities.getUuid().slice(0,6),c=Utilities.getUuid().replace(/-/g,'').slice(0,10);txt_(ops_().getSheetByName('Empleados'),[id,n,sha_(c),'SI',new Date().toISOString()]);return id+'.'+c;}
function auth_(t){const[id,c]=String(t||'').split('.');if(!id||!c)throw 'Acceso no autorizado';
  const r=ops_().getSheetByName('Empleados').getDataRange().getValues().find(x=>x[0]===id);
  if(!r||r[3]!=='SI'||r[2]!==sha_(c))throw 'Acceso no autorizado';return{id,name:r[1]};}
function cfg_(k){const r=ops_().getSheetByName('Config').getDataRange().getValues().find(x=>x[0]===k);return r&&r[1];}
function need_(u,q,m){if(!q.g||!q.g[m]||CacheService.getScriptCache().get('g'+u.id+m)!==q.g[m])throw 'Ingresá el PIN';}

function ix_(v){const h=v[0].map(x=>norm_(x).replace(/[^a-z]/g,'')),i=['codigo','descripcion','codbarras','contado','stock'].map(x=>h.indexOf(x));
  if(i.includes(-1))throw 'Faltan encabezados en la hoja de productos. Deben ser: Codigo, Descripcion, Cód.Barras, CONTADO, Stock';return i;}
function clear_(n){const c=CacheService.getScriptCache(),k='cat'+n,x=+c.get(k+'c')||0,ks=[k+'c'];for(let i=0;i<x;i++)ks.push(k+i);c.removeAll(ks);}
function cat_(n){const c=CacheService.getScriptCache(),k='cat'+n,cnt=+c.get(k+'c');
  if(cnt){let s='';for(let i=0;i<cnt;i++){const x=c.get(k+i);if(x==null){s=null;break;}s+=x;}if(s)return JSON.parse(s);}
  const v=base_().getSheetByName(n.toUpperCase()+' Productos').getDataRange().getDisplayValues(),ix=ix_(v);
  const d=v.slice(1).filter(r=>r[ix[0]]).map(r=>ix.map(i=>r[i])),s=JSON.stringify(d),parts=Math.ceil(s.length/30000);
  for(let i=0;i<parts;i++)c.put(k+i,s.substr(i*30000,30000),21600);c.put(k+'c',String(parts),21600);return d;}
function search_(n,q){q=norm_(q);if(!q)return[];const d=cat_(n),ex=d.filter(r=>r[0].toLowerCase()===q||r[2]===q);if(ex.length)return ex.slice(0,30);
  const w=q.split(/\s+/);return d.filter(r=>{const t=norm_(r[1]+' '+r[0]);return w.every(x=>t.includes(x));}).slice(0,30);}

function latest_(){const m={};ops_().getSheetByName('Documentos').getDataRange().getValues().slice(1).forEach(r=>{if(!m[r[0]]||r[4]>m[r[0]][4])m[r[0]]=r;});return m;}
function check_(d){if(!d.fecha)throw 'Falta la fecha';if(!d.items.length)throw 'No hay productos cargados';
  d.items.forEach(i=>{if(!i.cod)throw 'Hay productos sin código interno (pendientes): '+i.desc;if(!i.desc)throw 'Falta la descripción de '+i.cod;
    if(d.tipo==='stock'?!(i.q>=0):!(i.q>0))throw 'Cantidad inválida en '+i.cod;});
  if(d.tipo==='boletas'){if(!d.prov||!d.fact)throw 'Falta el proveedor o el número de factura';const s=d.items.reduce((a,i)=>a+i.imp,0);
    if(Math.abs(s-d.total)>0.009)throw 'El total cargado ($'+s.toFixed(2)+') no coincide con el declarado ($'+(+d.total).toFixed(2)+')';}}
function flat_(d,u){const K=d.neg.toUpperCase(),s=ops_().getSheetByName(K+' '+{compras:'Compras',boletas:'Boletas',stock:'Stock'}[d.tipo]);
  const r=d.items.map(i=>d.tipo==='compras'?[d.id,d.fecha,u.name,i.cod,i.desc,i.q,i.u||'',i.com||'']:d.tipo==='boletas'?[d.id,d.fecha,u.name,d.prov,d.fact,i.cod,i.desc,i.q,i.imp/i.q,i.imp]:[d.id,d.fecha,u.name,i.cod,i.desc,i.q]);
  const a=s.getLastRow()+1;s.getRange(a,d.tipo==='boletas'?6:4,r.length,1).setNumberFormat('@');s.getRange(a,1,r.length,r[0].length).setValues(r);}
function save_(u,d){if(!NEG[d.neg]||MODS.indexOf(d.tipo)<0||!Array.isArray(d.items))throw 'Documento inválido';
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{const sh=ops_().getSheetByName('Documentos'),cur=d.id&&latest_()[d.id];
    if(cur){if(cur[3]==='final')throw 'Este documento ya está finalizado y no se puede modificar';if(cur[4]!==d.rev)return{conflict:true,doc:JSON.parse(cur[7])};}
    else d.id=(d.tipo[0]+d.neg).toUpperCase()+'-'+Utilities.formatDate(new Date(),TZ,'yyyyMMdd-HHmmss');
    if(d.estado==='final')check_(d);
    d.rev=(cur?cur[4]:0)+1;d.user=u.name;
    sh.appendRow([d.id,d.neg,d.tipo,d.estado,d.rev,new Date().toISOString(),u.name,JSON.stringify(d)]);
    if(d.estado==='final')flat_(d,u);
    return{doc:d};}finally{lock.releaseLock();}}

function out_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function doGet(){return out_({ok:1});}
function doPost(e){let r;try{const q=JSON.parse(e.postData.contents);r=run_(q,auth_(q.t));}catch(x){const m=String(x&&x.message||x);r={err:m,auth:m==='Acceso no autorizado'};}return out_(r);}
function run_(q,u){const a=q.a,n=q.n;if(a==='ping')return{name:u.name};if(n&&!NEG[n])throw 'Negocio inválido';
  if(a==='pin'){if(['compras','boletas','stock','admin'].indexOf(q.m)<0)throw 'Módulo inválido';if(sha_(String(q.pin))!==cfg_('pin_'+q.m))throw 'PIN incorrecto';
    const g=Utilities.getUuid();CacheService.getScriptCache().put('g'+u.id+q.m,g,21600);return{g};}
  if(a==='search')return{r:search_(n,q.q)};
  if(a==='pend'){txt_(ops_().getSheetByName('Productos pendientes'),[new Date().toISOString(),NEG[n],String(q.bar||''),String(q.desc||''),'pendiente',u.name]);return{ok:1};}
  if(a==='list'){need_(u,q,q.m);const l=Object.values(latest_()).filter(r=>r[1]===n&&r[2]===q.m).sort((x,y)=>x[5]<y[5]?1:-1).slice(0,60)
    .map(r=>({id:r[0],estado:r[3],user:r[6],upd:r[5],n:JSON.parse(r[7]).items.length}));return{l};}
  if(a==='get'){need_(u,q,q.m);const r=latest_()[q.id];if(!r||r[1]!==n||r[2]!==q.m)throw 'Documento no encontrado';return{doc:JSON.parse(r[7])};}
  if(a==='save'){need_(u,q,q.d.tipo);return save_(u,q.d);}
  need_(u,q,'admin');const E=ops_().getSheetByName('Empleados');
  if(a==='emps')return{l:E.getDataRange().getValues().slice(1).map(r=>({id:r[0],name:r[1],on:r[3]==='SI'}))};
  if(a==='empAdd'){if(!q.name)throw 'Falta el nombre';return{code:addEmp_(String(q.name))};}
  if(a==='empOff'){const v=E.getDataRange().getValues(),i=v.findIndex(r=>r[0]===q.id);if(i<1)throw 'Empleado no encontrado';if(q.id===u.id)throw 'No podés revocar tu propio acceso';E.getRange(i+1,4).setValue('NO');return{ok:1};}
  if(a==='setPin'){if(!/^\d{1,8}$/.test(String(q.pin)))throw 'El PIN debe tener entre 1 y 8 números';const C=ops_().getSheetByName('Config'),v=C.getDataRange().getValues(),i=v.findIndex(r=>r[0]==='pin_'+q.m);if(i<0)throw 'Módulo inválido';C.getRange(i+1,2).setNumberFormat('@').setValue(sha_(String(q.pin)));return{ok:1};}
  if(a==='reload'){clear_(n);cat_(n);return{ok:1};}
  if(a==='bar'){(q.all?Object.keys(NEG):[n]).forEach(k=>{const s=base_().getSheetByName(k.toUpperCase()+' Productos'),v=s.getDataRange().getDisplayValues(),ix=ix_(v);
    v.forEach((r,i)=>{if(i&&r[ix[0]]===String(q.cod))s.getRange(i+1,ix[2]+1).setNumberFormat('@').setValue(String(q.bar));});clear_(k);});return{ok:1};}
  throw 'Acción desconocida';}
