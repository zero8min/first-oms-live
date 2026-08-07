const http=require('http'),fs=require('fs'),path=require('path'),url=require('url'),crypto=require('crypto'),XLSX=require('xlsx');
const sseClients=new Map();
function broadcastCustomers(list,tenantCode){
 const payload=`event: customers\ndata: ${JSON.stringify(list)}\n\n`;
 for(const [res,code] of [...sseClients.entries()]){
  if(tenantCode&&code!==tenantCode)continue;
  try{res.write(payload)}catch(e){sseClients.delete(res)}
 }
}
const ROOT=__dirname;
// Render Persistent Disk mount path. DATA_DIR can be overridden for local tests.
const DATA_ROOT=process.env.DATA_DIR||path.join(ROOT,'data');
const DATA=path.join(DATA_ROOT,'customers.json'), CUSTOMER_BACKUP=path.join(DATA_ROOT,'customers-backup.json'), CUSTOMER_XLSX=path.join(DATA_ROOT,'customers.xlsx'), INTEGRATIONS=path.join(DATA_ROOT,'integrations.json'), BACKUP_DIR=path.join(DATA_ROOT,'backups'), SEND_HISTORY=path.join(DATA_ROOT,'send-history.json'), YT_AUTH=path.join(DATA_ROOT,'youtube-auth.json');
const STATE_DATA=path.join(DATA_ROOT,'server-state.json'), STATE_BACKUP=path.join(DATA_ROOT,'server-state-backup.json'), STATE_XLSX=path.join(DATA_ROOT,'sales-list.xlsx'), SALES_ARCHIVE_DIR=path.join(DATA_ROOT,'sales-archives'), STATE_BACKUP_DIR=path.join(DATA_ROOT,'state-backups');
const ACCOUNTS=path.join(DATA_ROOT,'accounts.json'), ACCOUNTS_BACKUP=path.join(DATA_ROOT,'accounts-backup.json'), ACCOUNTS_XLSX=path.join(DATA_ROOT,'accounts.xlsx'), ACCOUNT_BACKUP_DIR=path.join(DATA_ROOT,'account-backups'), TENANTS_DIR=path.join(DATA_ROOT,'tenants');
const sessions=new Map();
if(!fs.existsSync(DATA_ROOT))fs.mkdirSync(DATA_ROOT,{recursive:true});
if(!fs.existsSync(DATA))fs.writeFileSync(DATA,'[]','utf8');
if(!fs.existsSync(YT_AUTH))fs.writeFileSync(YT_AUTH,'{}','utf8');
if(!fs.existsSync(INTEGRATIONS))fs.writeFileSync(INTEGRATIONS,'{}','utf8');
if(!fs.existsSync(SEND_HISTORY))fs.writeFileSync(SEND_HISTORY,'[]','utf8');
if(!fs.existsSync(STATE_DATA))fs.writeFileSync(STATE_DATA,JSON.stringify({orders:[],customers:readCustomers(),payments:[],settings:{},csRecords:[],shippingRecords:[]},null,2),'utf8');
for(const d of [SALES_ARCHIVE_DIR,STATE_BACKUP_DIR,ACCOUNT_BACKUP_DIR,TENANTS_DIR])if(!fs.existsSync(d))fs.mkdirSync(d,{recursive:true});
const mime={'.html':'text/html; charset=utf-8','.js':'application/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.json':'application/json; charset=utf-8','.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.ico':'image/x-icon'};
function readCustomers(){try{return JSON.parse(fs.readFileSync(DATA,'utf8'))}catch(e){return[]}}
function readIntegrations(){try{return JSON.parse(fs.readFileSync(INTEGRATIONS,'utf8'))}catch(e){return{}}}
function saveIntegrations(v){fs.writeFileSync(INTEGRATIONS,JSON.stringify(v,null,2),'utf8')}
function writeCustomerExcel(list){
 try{
  const rows=(list||[]).map(c=>({
   '등록일시':c.joinedAt||'', '등록경로':c.source||'', '성명':c.name||'', '닉네임':c.nickname||c.nick||'',
   '전화번호':c.phone||'', '우편번호':c.postalCode||'', '기본주소':c.address||'', '상세주소':c.detailAddress||'',
   '배송요청사항':c.memo||'', '고객ID':c.id||''
  }));
  const wb=XLSX.utils.book_new(), ws=XLSX.utils.json_to_sheet(rows);
  ws['!cols']=[{wch:22},{wch:12},{wch:12},{wch:18},{wch:16},{wch:10},{wch:36},{wch:28},{wch:28},{wch:28}];
  XLSX.utils.book_append_sheet(wb,ws,'고객DB');XLSX.writeFile(wb,CUSTOMER_XLSX);return true
 }catch(e){console.error('고객 엑셀 저장 실패',e);return false}
}

function backupCustomers(){
 try{
  if(!fs.existsSync(BACKUP_DIR))fs.mkdirSync(BACKUP_DIR,{recursive:true});
  const current=readCustomers();
  const text=JSON.stringify(current,null,2);
  fs.writeFileSync(CUSTOMER_BACKUP,text,'utf8');
  const stamp=new Date().toISOString().replace(/[:.]/g,'-');
  fs.writeFileSync(path.join(BACKUP_DIR,`customers-${stamp}.json`),text,'utf8');
  const files=fs.readdirSync(BACKUP_DIR).filter(x=>/^customers-.*\.json$/.test(x)).sort();
  while(files.length>30){const old=files.shift();try{fs.unlinkSync(path.join(BACKUP_DIR,old))}catch(e){}}
  return {count:current.length,file:'customers-backup.json'};
 }catch(e){return {count:0,error:e.message}}
}
function saveCustomers(v){
 backupCustomers();
 fs.writeFileSync(DATA,JSON.stringify(v,null,2),'utf8');
 fs.writeFileSync(CUSTOMER_BACKUP,JSON.stringify(v,null,2),'utf8');
 writeCustomerExcel(v);
 broadcastCustomers(v);
}

function readState(){try{return JSON.parse(fs.readFileSync(STATE_DATA,'utf8'))}catch(e){return {orders:[],customers:readCustomers(),payments:[],settings:{},csRecords:[],shippingRecords:[]}}}
function writeStateExcel(st){
 try{
  const wb=XLSX.utils.book_new();
  const orders=(st.orders||[]).map(o=>({'방송일':o.date||'','닉네임':o.nick||'','상품번호':o.productNo||'','상품명':o.item||'','수량':o.qty||0,'단가':o.unit||0,'금액':o.amount||0,'배송비':o.fee||0,'원본파일':o.source||''}));
  const ws=XLSX.utils.json_to_sheet(orders);XLSX.utils.book_append_sheet(wb,ws,'판매리스트');
  XLSX.writeFile(wb,STATE_XLSX);return true
 }catch(e){console.error('판매리스트 엑셀 저장 실패',e);return false}
}
function backupState(st){
 try{
  const text=JSON.stringify(st,null,2), stamp=new Date().toISOString().replace(/[:.]/g,'-');
  fs.writeFileSync(STATE_BACKUP,text,'utf8');
  fs.writeFileSync(path.join(STATE_BACKUP_DIR,`state-${stamp}.json`),text,'utf8');
  const files=fs.readdirSync(STATE_BACKUP_DIR).filter(x=>/^state-.*\.json$/.test(x)).sort();
  while(files.length>100){const old=files.shift();try{fs.unlinkSync(path.join(STATE_BACKUP_DIR,old))}catch(e){}}
 }catch(e){console.error('전체 상태 백업 실패',e)}
}
function archiveSalesByDate(st){
 try{
  const by={};for(const o of (st.orders||[])){const d=o.date||'날짜없음';(by[d]||(by[d]=[])).push(o)}
  for(const [d,rows] of Object.entries(by)){
   fs.writeFileSync(path.join(SALES_ARCHIVE_DIR,`${d}.json`),JSON.stringify({date:d,count:rows.length,orders:rows},null,2),'utf8');
   const wb=XLSX.utils.book_new(),ws=XLSX.utils.json_to_sheet(rows);XLSX.utils.book_append_sheet(wb,ws,'판매리스트');XLSX.writeFile(wb,path.join(SALES_ARCHIVE_DIR,`${d}.xlsx`));
  }
 }catch(e){console.error('날짜별 판매리스트 보존 실패',e)}
}
function saveState(st){
 const next={...readState(),...st,customers:Array.isArray(st.customers)?st.customers:readCustomers(),updatedAt:new Date().toISOString()};
 backupState(readState());
 fs.writeFileSync(STATE_DATA,JSON.stringify(next,null,2),'utf8');
 fs.writeFileSync(STATE_BACKUP,JSON.stringify(next,null,2),'utf8');
 writeStateExcel(next);archiveSalesByDate(next);
 if(Array.isArray(next.customers))saveCustomers(next.customers);
 return next
}
function listSalesArchives(){try{return fs.readdirSync(SALES_ARCHIVE_DIR).filter(x=>x.endsWith('.json')).sort().reverse().map(x=>{const d=JSON.parse(fs.readFileSync(path.join(SALES_ARCHIVE_DIR,x),'utf8'));return {date:d.date,count:d.count,file:x}})}catch(e){return[]}}


function passwordHash(password,salt=crypto.randomBytes(16).toString('hex')){
 const hash=crypto.scryptSync(String(password),salt,64).toString('hex');return `${salt}:${hash}`
}
function verifyPassword(password,stored){
 try{const [salt,hex]=String(stored||'').split(':');if(!salt||!hex)return false;const a=Buffer.from(hex,'hex'),b=crypto.scryptSync(String(password),salt,64);return a.length===b.length&&crypto.timingSafeEqual(a,b)}catch(e){return false}
}
function readJsonArraySafe(file){
 try{const v=JSON.parse(fs.readFileSync(file,'utf8'));return Array.isArray(v)?v:null}catch(e){return null}
}
function atomicWrite(file,text){
 const tmp=`${file}.tmp-${process.pid}-${Date.now()}`;fs.writeFileSync(tmp,text,'utf8');fs.renameSync(tmp,file)
}
function latestAccountBackup(){
 try{return fs.readdirSync(ACCOUNT_BACKUP_DIR).filter(x=>/^accounts-.*\.json$/.test(x)).sort().reverse().map(x=>path.join(ACCOUNT_BACKUP_DIR,x))[0]||null}catch(e){return null}
}
function readAccounts(){
 const primary=readJsonArraySafe(ACCOUNTS);if(primary)return primary;
 const fallback=readJsonArraySafe(ACCOUNTS_BACKUP);if(fallback){try{atomicWrite(ACCOUNTS,JSON.stringify(fallback,null,2))}catch(e){};return fallback}
 const latest=latestAccountBackup(), historical=latest&&readJsonArraySafe(latest);if(historical){try{atomicWrite(ACCOUNTS,JSON.stringify(historical,null,2));atomicWrite(ACCOUNTS_BACKUP,JSON.stringify(historical,null,2))}catch(e){};return historical}
 return []
}
function tenantDir(code){return path.join(TENANTS_DIR,String(code||'UNKNOWN').replace(/[^A-Za-z0-9_-]/g,''))}
function ensureTenantStorage(account){
 if(!account||!account.code)return;
 const dir=tenantDir(account.code);fs.mkdirSync(path.join(dir,'sales-archives'),{recursive:true});
 const defaults={
  'customers.json':'[]',
  'server-state.json':JSON.stringify({orders:[],customers:[],payments:[],settings:{},csRecords:[],shippingRecords:[]},null,2),
  'shipping.json':'[]',
  'payments.json':'[]',
  'settlements.json':'[]',
  'cs-history.json':'[]',
  'message-history.json':'[]',
  'solapi-settings.json':'{}',
  'youtube-auth.json':'{}',
  'tenant-settings.json':JSON.stringify({tenantCode:account.code,company:account.company||'',createdAt:new Date().toISOString()},null,2)
 };
 for(const [name,text] of Object.entries(defaults)){const f=path.join(dir,name);if(!fs.existsSync(f))atomicWrite(f,text)}
}
function writeAccountsExcel(list){
 try{const rows=(list||[]).map(a=>({'거래처코드':a.code||'','아이디':a.username||'','거래처명':a.company||'','대표자':a.ownerName||'','연락처':a.phone||'','권한':a.role||'tenant','상태':a.status||'pending','가입일':a.createdAt||'','최근로그인':a.lastLoginAt||''}));const wb=XLSX.utils.book_new(),ws=XLSX.utils.json_to_sheet(rows);XLSX.utils.book_append_sheet(wb,ws,'거래처계정');XLSX.writeFile(wb,ACCOUNTS_XLSX)}catch(e){console.error('계정 엑셀 실패',e)}
}
function saveAccounts(list){
 if(!Array.isArray(list))throw new Error('계정 데이터 형식 오류');
 const old=readAccounts(),stamp=new Date().toISOString().replace(/[:.]/g,'-'),nextText=JSON.stringify(list,null,2),oldText=JSON.stringify(old,null,2);
 // First preserve the last known-good copy. Never replace valid data with an empty accidental payload.
 if(old.length>0&&list.length===0)throw new Error('거래처 계정 전체 초기화가 차단되었습니다.');
 try{
  atomicWrite(ACCOUNTS_BACKUP,oldText);
  atomicWrite(path.join(ACCOUNT_BACKUP_DIR,`accounts-${stamp}.json`),oldText);
  const fsx=fs.readdirSync(ACCOUNT_BACKUP_DIR).filter(x=>/^accounts-.*\.json$/.test(x)).sort();while(fsx.length>100){try{fs.unlinkSync(path.join(ACCOUNT_BACKUP_DIR,fsx.shift()))}catch(e){}}
 }catch(e){console.error('계정 사전백업 실패',e);throw new Error('계정 백업에 실패하여 변경을 중단했습니다.')}
 atomicWrite(ACCOUNTS,nextText);atomicWrite(ACCOUNTS_BACKUP,nextText);writeAccountsExcel(list);
 for(const account of list)ensureTenantStorage(account);
 return true
}
const DEFAULT_ADMIN_ID='firstadmin',DEFAULT_ADMIN_PASSWORD='@31062224',OWNER_CREDENTIAL_VERSION='7.9';
function ensureOwnerAccount(){
 let list=readAccounts(),changed=false;
 let owner=list.find(a=>a.role==='superadmin'||a.code==='FIRST-MASTER');
 if(!owner){
  owner={id:crypto.randomUUID(),code:'FIRST-MASTER',username:DEFAULT_ADMIN_ID,passwordHash:passwordHash(DEFAULT_ADMIN_PASSWORD),company:'FIRST OMS',ownerName:'최고관리자',phone:'',role:'superadmin',status:'active',mustChangePassword:false,createdAt:new Date().toISOString(),bootstrapVersion:'7.9',credentialSetupVersion:OWNER_CREDENTIAL_VERSION};
  list.push(owner);changed=true;
 }else{
  if(owner.role!=='superadmin'){owner.role='superadmin';changed=true}
  if(owner.status!=='active'){owner.status='active';changed=true}
  if(owner.code!=='FIRST-MASTER'){owner.code='FIRST-MASTER';changed=true}
  // v7.9에서 주인님 최고관리자 계정을 최초 1회 확정한다.
  // credentialSetupVersion이 저장된 뒤에는 재배포/재시작해도 변경한 비밀번호를 덮어쓰지 않는다.
  if(owner.credentialSetupVersion!==OWNER_CREDENTIAL_VERSION){
   owner.username=DEFAULT_ADMIN_ID;
   owner.passwordHash=passwordHash(DEFAULT_ADMIN_PASSWORD);
   owner.mustChangePassword=false;
   owner.credentialSetupVersion=OWNER_CREDENTIAL_VERSION;
   owner.passwordChangedAt=new Date().toISOString();
   changed=true;
  }else{
   if(!owner.username){owner.username=DEFAULT_ADMIN_ID;changed=true}
   if(!owner.passwordHash){owner.passwordHash=passwordHash(DEFAULT_ADMIN_PASSWORD);owner.mustChangePassword=false;changed=true}
  }
 }
 if(changed)saveAccounts(list);else for(const a of list)ensureTenantStorage(a);
 console.log('[LOGIN] 최고관리자 계정 준비 완료');
}
function cookies(req){return Object.fromEntries(String(req.headers.cookie||'').split(';').map(x=>x.trim()).filter(Boolean).map(x=>{const i=x.indexOf('=');return [decodeURIComponent(x.slice(0,i)),decodeURIComponent(x.slice(i+1))]}))}
function currentUser(req){const sid=cookies(req).ddaeng_session,ss=sessions.get(sid);if(!ss||ss.expiresAt<Date.now()){if(sid)sessions.delete(sid);return null}return readAccounts().find(a=>a.id===ss.userId&&a.status==='active')||null}
function issueSession(req,res,user){const sid=crypto.randomBytes(32).toString('hex');sessions.set(sid,{userId:user.id,activeTenantCode:user.role==='superadmin'?'FIRST-MASTER':user.code,expiresAt:Date.now()+1000*60*60*24*7});const secure=String(req.headers['x-forwarded-proto']||'').includes('https')?'; Secure':'';res.setHeader('Set-Cookie',`ddaeng_session=${sid}; Path=/; HttpOnly; SameSite=Lax; Max-Age=604800${secure}`)}
function clearSession(req,res){const sid=cookies(req).ddaeng_session;if(sid)sessions.delete(sid);res.setHeader('Set-Cookie','ddaeng_session=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0')}
function newTenantCode(list){let n=1;const used=new Set(list.map(a=>a.code));while(used.has(`FIRST-${String(n).padStart(4,'0')}`))n++;return `FIRST-${String(n).padStart(4,'0')}`}
ensureOwnerAccount();


// ===== v7.5 strict multi-tenant isolation =====
const DDAENG_TENANT_CODE=process.env.DDAENG_TENANT_CODE||'FIRST-0001';
function ensureDdaengTenantAccount(){
 let list=readAccounts(),changed=false;
 let tenant=list.find(a=>a.code===DDAENG_TENANT_CODE);
 if(!tenant){
  tenant={
   id:crypto.randomUUID(),code:DDAENG_TENANT_CODE,username:'01021842344',
   passwordHash:passwordHash('@21842344'),
   company:'땡라이브',ownerName:'땡라이브',phone:'',role:'tenant',status:'active',
   systemManaged:true,credentialSetupVersion:'7.8',createdAt:new Date().toISOString()
  };
  list.push(tenant);changed=true;
 }else{
  if(tenant.role!=='tenant'){tenant.role='tenant';changed=true}
  if(tenant.status!=='active'){tenant.status='active';changed=true}
  if(!tenant.company||tenant.company==='FIRST OMS'){tenant.company='땡라이브';changed=true}
  if(!tenant.ownerName){tenant.ownerName='땡라이브';changed=true}
  // 요청한 땡라이브 계정은 이번 버전에서 최초 1회 확정하고 이후 변경값은 보존한다.
  if(tenant.credentialSetupVersion!=='7.8'){
   tenant.username='01021842344';
   tenant.passwordHash=passwordHash('@21842344');
   tenant.mustChangePassword=false;
   tenant.credentialSetupVersion='7.8';
   tenant.passwordChangedAt=new Date().toISOString();
   changed=true;
  }
  tenant.systemManaged=true;
 }
 if(changed)saveAccounts(list);else ensureTenantStorage(tenant);
 return tenant;
}
ensureDdaengTenantAccount();
// v7.34: 두 번째 거래처 MD유통. 기존 계정/비밀번호/데이터가 있으면 절대 덮어쓰지 않는다.
const MD_TENANT_CODE=process.env.MD_TENANT_CODE||'MD-0002';
function ensureMdTenantAccount(){
 let list=readAccounts(),changed=false,tenant=list.find(a=>a.code===MD_TENANT_CODE||String(a.company||'').trim()==='MD유통');
 if(!tenant){
  tenant={id:crypto.randomUUID(),code:MD_TENANT_CODE,username:process.env.MD_TENANT_USERNAME||'md0002',passwordHash:passwordHash(process.env.MD_TENANT_PASSWORD||'MD@0002!'),company:'MD유통',ownerName:'MD유통',phone:'',role:'tenant',status:'active',mustChangePassword:true,createdAt:new Date().toISOString(),bootstrapVersion:'7.34'};
  list.push(tenant);changed=true;
 }else{
  if(!tenant.code){tenant.code=MD_TENANT_CODE;changed=true}
  if(!tenant.company){tenant.company='MD유통';changed=true}
  if(tenant.role!=='tenant'){tenant.role='tenant';changed=true}
  if(!tenant.status){tenant.status='active';changed=true}
 }
 if(changed)saveAccounts(list);else ensureTenantStorage(tenant);
 return tenant;
}
ensureMdTenantAccount();
function tenantFile(code,name){ensureTenantStorage({code});return path.join(tenantDir(code),name)}
function tenantBackupDir(code,name){const d=path.join(tenantDir(code),name);fs.mkdirSync(d,{recursive:true});return d}
function readJsonObject(file,fallback){try{const v=JSON.parse(fs.readFileSync(file,'utf8'));return v&&typeof v==='object'?v:fallback}catch(e){return fallback}}
function selectedTenantCode(req){
 const user=currentUser(req);if(!user)return null;
 if(user.role!=='superadmin')return user.code;
 const sid=cookies(req).ddaeng_session,ss=sessions.get(sid);
 return (ss&&ss.activeTenantCode)||'FIRST-MASTER';
}
function assertTenantAccess(req,requested){const user=currentUser(req);if(!user)throw new Error('로그인이 필요합니다.');const code=String(requested||selectedTenantCode(req)||'');if(user.role!=='superadmin'&&code!==user.code)throw new Error('다른 거래처 데이터에는 접근할 수 없습니다.');return code}
function tenantReadCustomers(code){return readJsonObject(tenantFile(code,'customers.json'),[])}
function tenantReadState(code){const st=readJsonObject(tenantFile(code,'server-state.json'),{orders:[],customers:[],payments:[],settings:{},csRecords:[],shippingRecords:[]});st.customers=tenantReadCustomers(code);return st}
function tenantWriteCustomers(code,list){if(!Array.isArray(list))throw new Error('고객 데이터 형식 오류');const file=tenantFile(code,'customers.json'),backup=tenantFile(code,'customers-backup.json'),dir=tenantBackupDir(code,'backups');const old=tenantReadCustomers(code),stamp=new Date().toISOString().replace(/[:.]/g,'-');atomicWrite(backup,JSON.stringify(old,null,2));atomicWrite(path.join(dir,`customers-${stamp}.json`),JSON.stringify(old,null,2));atomicWrite(file,JSON.stringify(list,null,2));broadcastCustomers(list,code);}
function tenantWriteState(code,patch){const old=tenantReadState(code),authoritativeCustomers=tenantReadCustomers(code),safePatch={...(patch||{})};delete safePatch.customers;const next={...old,...safePatch,customers:authoritativeCustomers,updatedAt:new Date().toISOString()};const backup=tenantFile(code,'server-state-backup.json'),dir=tenantBackupDir(code,'state-backups');atomicWrite(backup,JSON.stringify(old,null,2));atomicWrite(path.join(dir,`state-${new Date().toISOString().replace(/[:.]/g,'-')}.json`),JSON.stringify(old,null,2));atomicWrite(tenantFile(code,'server-state.json'),JSON.stringify(next,null,2));return next}
function tenantReadIntegrations(code){return readJsonObject(tenantFile(code,'solapi-settings.json'),{})}
function tenantSaveIntegrations(code,v){atomicWrite(tenantFile(code,'solapi-settings.json'),JSON.stringify(v,null,2))}
function tenantReadSendHistory(code){return readJsonObject(tenantFile(code,'message-history.json'),[])}
function tenantAppendSendHistory(code,v){const a=tenantReadSendHistory(code);a.unshift(v);atomicWrite(tenantFile(code,'message-history.json'),JSON.stringify(a.slice(0,5000),null,2))}
function tenantSalesArchiveDir(code){return tenantBackupDir(code,'sales-archives')}
function tenantListSalesArchives(code){try{return fs.readdirSync(tenantSalesArchiveDir(code)).filter(x=>x.endsWith('.json')).sort().reverse().map(x=>{const d=readJsonObject(path.join(tenantSalesArchiveDir(code),x),{});return {date:d.date,count:d.count,file:x}})}catch(e){return[]}}
function copyFileIfTargetEmpty(src,dst){
 try{
  if(!fs.existsSync(src))return false;
  let copy=!fs.existsSync(dst)||fs.statSync(dst).size<5;
  if(!copy){
   const a=readJsonObject(src,null),b=readJsonObject(dst,null);
   if(Array.isArray(a)&&Array.isArray(b))copy=b.length===0&&a.length>0;
   else if(a&&b&&typeof a==='object'&&typeof b==='object')copy=Object.keys(b).length===0&&Object.keys(a).length>0;
  }
  if(copy){fs.copyFileSync(src,dst);return true}
 }catch(e){console.error('자료 복원 실패',path.basename(dst),e.message)}
 return false
}
function migrateOwnerLegacyDataOnce(){
 const marker=path.join(DATA_ROOT,'.v78-owner-legacy-preserved');
 const owner=readAccounts().find(a=>a.role==='superadmin');if(!owner)return;
 ensureTenantStorage(owner);
 const pairs=[[DATA,'customers.json'],[CUSTOMER_BACKUP,'customers-backup.json'],[STATE_DATA,'server-state.json'],[STATE_BACKUP,'server-state-backup.json'],[INTEGRATIONS,'solapi-settings.json'],[SEND_HISTORY,'message-history.json'],[YT_AUTH,'youtube-auth.json']];
 for(const [src,name] of pairs)copyFileIfTargetEmpty(src,tenantFile(owner.code,name));
 if(!fs.existsSync(marker))atomicWrite(marker,JSON.stringify({tenantCode:owner.code,at:new Date().toISOString(),note:'기존 고객·정산·솔라피 보존'},null,2));
}
function restoreDdaengCustomersAndSeparateSolapiOnce(){
 const marker=path.join(DATA_ROOT,'.v78-ddaeng-customer-restored-solapi-separated');if(fs.existsSync(marker))return;
 const tenant=readAccounts().find(a=>a.code===DDAENG_TENANT_CODE),owner=readAccounts().find(a=>a.role==='superadmin');if(!tenant||!owner)return;
 ensureTenantStorage(tenant);ensureTenantStorage(owner);
 const sourceCustomers=tenantReadCustomers(owner.code).length?tenantReadCustomers(owner.code):readCustomers();
 if(sourceCustomers.length){
  const current=tenantReadCustomers(tenant.code);
  if(current.length===0)tenantWriteCustomers(tenant.code,sourceCustomers);
  const st=tenantReadState(tenant.code);if(!Array.isArray(st.customers)||st.customers.length===0){st.customers=sourceCustomers;tenantWriteState(tenant.code,st)}
 }
 // v7.34: 기존 거래처의 솔라피/카카오 설정은 절대 비우거나 덮어쓰지 않는다.
 const solapiFile=tenantFile(tenant.code,'solapi-settings.json');
 const old=readJsonObject(solapiFile,{});
 if(Object.keys(old).length){atomicWrite(tenantFile(tenant.code,'solapi-settings-preserved-v734.json'),JSON.stringify(old,null,2))}
 atomicWrite(path.join(tenantDir(tenant.code),'tenant-settings.json'),JSON.stringify({company:'땡라이브',tenantCode:tenant.code,customersRestoredAt:new Date().toISOString(),solapiMode:'tenant-own-setting'},null,2));
 atomicWrite(marker,JSON.stringify({tenantCode:tenant.code,customerCount:sourceCustomers.length,at:new Date().toISOString()},null,2));
}
migrateOwnerLegacyDataOnce();
restoreDdaengCustomersAndSeparateSolapiOnce();
function restoreBundledDdaengDataIfEmpty(){
 try{
  const current=tenantReadState(DDAENG_TENANT_CODE);
  if(Array.isArray(current.orders)&&current.orders.length>0)return;
  const bundled=path.join(ROOT,'data','initial-backup.json');
  if(!fs.existsSync(bundled))return;
  const backup=readJsonObject(bundled,{}),source=backup.state||backup;
  if(!Array.isArray(source.orders)||source.orders.length===0)return;
  const restored={...current,...source,shippingScans:current.shippingScans||source.shippingScans||{},updatedAt:new Date().toISOString()};
  if(Array.isArray(restored.customers))tenantWriteCustomers(DDAENG_TENANT_CODE,restored.customers);
  tenantWriteState(DDAENG_TENANT_CODE,restored);
  console.log(`[DATA RECOVERY] ${DDAENG_TENANT_CODE}: orders ${restored.orders.length}, customers ${(restored.customers||[]).length}, payments ${(restored.payments||[]).length}`);
 }catch(e){console.error('[DATA RECOVERY] 실패:',e.message)}
}
restoreBundledDdaengDataIfEmpty();

function readBody(req,max=1024*1024){
 return new Promise((resolve,reject)=>{
  let body='';req.on('data',d=>{body+=d;if(body.length>max){reject(new Error('요청이 너무 큽니다'));req.destroy()}});
  req.on('end',()=>resolve(body));req.on('error',reject)
 })
}
function onlyDigits(v){return String(v||'').replace(/[^0-9]/g,'')}
function solapiAuth(apiKey,apiSecret){
 const date=new Date().toISOString(),salt=crypto.randomBytes(16).toString('hex');
 const signature=crypto.createHmac('sha256',apiSecret).update(date+salt).digest('hex');
 return `HMAC-SHA256 apiKey=${apiKey}, date=${date}, salt=${salt}, signature=${signature}`
}
async function sendSolapiSms(to,text,tenantCode){
 const cfg=solapiConfig(tenantCode),apiKey=cfg.apiKey,apiSecret=cfg.apiSecret,sender=cfg.sender;
 if(!apiKey||!apiSecret||!sender)throw new Error('Render 환경변수 SOLAPI_API_KEY, SOLAPI_API_SECRET, SOLAPI_SENDER를 확인해 주세요.');
 const receiver=onlyDigits(to);
 if(receiver.length<10)throw new Error('수신번호가 올바르지 않습니다.');
 const payload={messages:[{to:receiver,from:sender,text:String(text||'').slice(0,1900),autoTypeDetect:true}],showMessageList:true};
 const r=await fetch('https://api.solapi.com/messages/v4/send-many/detail',{
  method:'POST',headers:{Authorization:solapiAuth(apiKey,apiSecret),'Content-Type':'application/json'},body:JSON.stringify(payload)
 });
 const raw=await r.text();let data={};try{data=JSON.parse(raw)}catch(e){data={raw}}
 if(!r.ok)throw new Error(data.errorMessage||data.message||data.errorCode||('SOLAPI HTTP '+r.status));
 if(Array.isArray(data.failedMessageList)&&data.failedMessageList.length){
  const f=data.failedMessageList[0];throw new Error(f.statusMessage||f.errorMessage||'SOLAPI 접수 실패')
 }
 return data
}


function solapiConfig(tenantCode){
 const f=tenantCode?tenantReadIntegrations(tenantCode):readIntegrations();
 if(tenantCode)return {apiKey:f.apiKey||'',apiSecret:f.apiSecret||'',sender:onlyDigits(f.sender),pfId:f.pfId||'',templateId:f.templateId||''};
 return {
  apiKey:f.apiKey||process.env.SOLAPI_API_KEY||'',
  apiSecret:f.apiSecret||process.env.SOLAPI_API_SECRET||'',
  sender:onlyDigits(f.sender||process.env.SOLAPI_SENDER),
  pfId:f.pfId||process.env.SOLAPI_KAKAO_PF_ID||process.env.SOLAPI_PF_ID||'',
  templateId:f.templateId||process.env.SOLAPI_KAKAO_TEMPLATE_ID||process.env.SOLAPI_TEMPLATE_ID||''
 }
}
async function sendSolapiKakao(to,variables,text,tenantCode){
 const cfg=solapiConfig(tenantCode);
 if(!cfg.apiKey||!cfg.apiSecret||!cfg.sender)throw new Error('SOLAPI API Key·Secret·발신번호 환경변수를 확인해 주세요.');
 if(!cfg.pfId||!cfg.templateId)throw new Error('SOLAPI_KAKAO_PF_ID와 SOLAPI_KAKAO_TEMPLATE_ID를 Render 환경변수에 등록해 주세요.');
 const receiver=onlyDigits(to);
 if(receiver.length<10)throw new Error('수신번호가 올바르지 않습니다.');
 const vars={};
 Object.entries(variables||{}).forEach(([k,v])=>{vars[String(k)]=String(v??'')});
 const payload={messages:[{
  to:receiver,
  from:cfg.sender,
  text:String(text||'').slice(0,1900),
  kakaoOptions:{pfId:cfg.pfId,templateId:cfg.templateId,variables:vars,disableSms:false}
 }],showMessageList:true};
 const r=await fetch('https://api.solapi.com/messages/v4/send-many/detail',{
  method:'POST',headers:{Authorization:solapiAuth(cfg.apiKey,cfg.apiSecret),'Content-Type':'application/json'},body:JSON.stringify(payload)
 });
 const raw=await r.text();let data={};try{data=JSON.parse(raw)}catch(e){data={raw}}
 if(!r.ok)throw new Error(data.errorMessage||data.message||data.errorCode||('SOLAPI HTTP '+r.status));
 if(Array.isArray(data.failedMessageList)&&data.failedMessageList.length){
  const f=data.failedMessageList[0];throw new Error(f.statusMessage||f.errorMessage||'알림톡 접수 실패')
 }
 return data
}


function readSendHistory(){try{return JSON.parse(fs.readFileSync(SEND_HISTORY,'utf8'))}catch(e){return[]}}
function appendSendHistory(v){const a=readSendHistory();a.unshift(v);fs.writeFileSync(SEND_HISTORY,JSON.stringify(a.slice(0,5000),null,2),'utf8')}
async function sendSolapiMms(to,imageBase64,subject,text,tenantCode){
 const cfg=solapiConfig(tenantCode);if(!cfg.apiKey||!cfg.apiSecret||!cfg.sender)throw new Error('SOLAPI API Key·Secret·승인 발신번호를 확인해 주세요.');
 const receiver=onlyDigits(to);if(receiver.length<10)throw new Error('수신번호가 올바르지 않습니다.');
 const clean=String(imageBase64||'').replace(/^data:image\/jpeg;base64,/,'');if(!clean)throw new Error('정산서 이미지가 없습니다.');
 const bytes=Buffer.from(clean,'base64');if(bytes.length>200*1024)throw new Error(`이미지 용량이 200KB를 넘습니다 (${Math.ceil(bytes.length/1024)}KB).`);
 const upload=await fetch('https://api.solapi.com/storage/v1/files',{method:'POST',headers:{Authorization:solapiAuth(cfg.apiKey,cfg.apiSecret),'Content-Type':'application/json'},body:JSON.stringify({file:clean,type:'MMS',name:'FIRST_OMS_receipt.jpg'})});
 const uraw=await upload.text();let ud={};try{ud=JSON.parse(uraw)}catch(e){ud={raw:uraw}}if(!upload.ok||!ud.fileId)throw new Error(ud.errorMessage||ud.message||`이미지 업로드 실패 ${upload.status}`);
 const payload={messages:[{to:receiver,from:cfg.sender,text:String(text||'정산서 이미지입니다.').slice(0,1900),subject:String(subject||'땡라이브 정산서').slice(0,40),imageId:ud.fileId,autoTypeDetect:true}],showMessageList:true};
 const r=await fetch('https://api.solapi.com/messages/v4/send-many/detail',{method:'POST',headers:{Authorization:solapiAuth(cfg.apiKey,cfg.apiSecret),'Content-Type':'application/json'},body:JSON.stringify(payload)});
 const raw=await r.text();let data={};try{data=JSON.parse(raw)}catch(e){data={raw}}if(!r.ok)throw new Error(data.errorMessage||data.message||data.errorCode||('SOLAPI HTTP '+r.status));if(Array.isArray(data.failedMessageList)&&data.failedMessageList.length){const f=data.failedMessageList[0];throw new Error(f.statusMessage||f.errorMessage||'MMS 접수 실패')}return {upload:ud,result:data}
}

function readYoutubeAuth(){try{return JSON.parse(fs.readFileSync(YT_AUTH,'utf8'))}catch(e){return{}}}
function saveYoutubeAuth(v){fs.writeFileSync(YT_AUTH,JSON.stringify(v,null,2),'utf8')}
function youtubeConfig(){
 return {
  clientId:process.env.YOUTUBE_CLIENT_ID||'',
  clientSecret:process.env.YOUTUBE_CLIENT_SECRET||'',
  redirectUri:process.env.YOUTUBE_REDIRECT_URI||'',
  apiKey:process.env.YOUTUBE_API_KEY||''
 }
}
async function youtubeToken(){
 const cfg=youtubeConfig(),auth=readJsonObject(tenantFile(tenantCode,'youtube-auth.json'),{});
 if(auth.access_token&&auth.expires_at>Date.now()+60000)return auth.access_token;
 if(auth.refresh_token&&cfg.clientId&&cfg.clientSecret){
  const body=new URLSearchParams({client_id:cfg.clientId,client_secret:cfg.clientSecret,refresh_token:auth.refresh_token,grant_type:'refresh_token'});
  const r=await fetch('https://oauth2.googleapis.com/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body});
  const j=await r.json();
  if(!r.ok)throw new Error(j.error_description||j.error||'유튜브 토큰 갱신 실패');
  const next={...auth,...j,refresh_token:auth.refresh_token,expires_at:Date.now()+(j.expires_in||3600)*1000};
  saveYoutubeAuth(next);return next.access_token
 }
 return ''
}
async function youtubeApi(pathname,params={}){
 const cfg=youtubeConfig(),token=await youtubeToken();
 const q=new URLSearchParams(params);
 if(!token&&cfg.apiKey)q.set('key',cfg.apiKey);
 if(!token&&!cfg.apiKey)throw new Error('유튜브 계정 연결 또는 YOUTUBE_API_KEY가 필요합니다.');
 const r=await fetch('https://www.googleapis.com/youtube/v3/'+pathname+'?'+q.toString(),{
  headers:token?{Authorization:'Bearer '+token}:{}
 });
 const j=await r.json();
 if(!r.ok){
  const msg=j?.error?.message||j?.error?.errors?.[0]?.reason||('YouTube API '+r.status);
  const err=new Error(msg);err.status=r.status;throw err
 }
 return j
}
async function youtubeApiPost(pathname,params={},bodyData={}){
 const token=await youtubeToken();
 if(!token)throw new Error('유튜브 계정 연결이 필요합니다. API 키만으로는 댓글 작성이 불가능합니다.');
 const q=new URLSearchParams(params);
 const r=await fetch('https://www.googleapis.com/youtube/v3/'+pathname+'?'+q.toString(),{
  method:'POST',
  headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},
  body:JSON.stringify(bodyData)
 });
 const j=await r.json();
 if(!r.ok){
  const msg=j?.error?.message||j?.error?.errors?.[0]?.reason||('YouTube API '+r.status);
  const err=new Error(msg);err.status=r.status;throw err
 }
 return j
}

async function liveChatIdForVideo(videoId){
 const j=await youtubeApi('videos',{part:'liveStreamingDetails',id:videoId});
 const item=j.items&&j.items[0];
 const id=item?.liveStreamingDetails?.activeLiveChatId;
 if(!id)throw new Error('현재 방송의 실시간 채팅 ID를 찾지 못했습니다. 방송이 실제 LIVE 상태이고 채팅이 켜져 있는지 확인해 주세요.');
 return id
}

function json(res,code,data){res.writeHead(code,{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'});res.end(JSON.stringify(data))}
writeCustomerExcel(readCustomers());
const server=http.createServer((req,res)=>{
 res.setHeader('Access-Control-Allow-Origin','*');
 res.setHeader('Access-Control-Allow-Headers','Content-Type');
 res.setHeader('Access-Control-Allow-Methods','GET,POST,DELETE,OPTIONS');
 if(req.method==='OPTIONS'){res.writeHead(204);return res.end()}
 const u=url.parse(req.url,true);
 if(u.pathname==='/api/health')return json(res,200,{ok:true,time:new Date().toISOString()});
 if(u.pathname==='/api/auth/login'&&req.method==='POST')return readBody(req).then(body=>{try{const d=JSON.parse(body||'{}'),list=readAccounts(),a=list.find(x=>x.username===String(d.username||'').trim());if(!a||!verifyPassword(d.password,a.passwordHash))return json(res,401,{ok:false,error:'아이디 또는 비밀번호가 맞지 않습니다.'});if(a.status!=='active')return json(res,403,{ok:false,error:a.status==='pending'?'최고관리자 승인 대기 중입니다.':'사용이 정지된 계정입니다.'});a.lastLoginAt=new Date().toISOString();saveAccounts(list);issueSession(req,res,a);return json(res,200,{ok:true,user:{id:a.id,username:a.username,company:a.company,code:a.code,role:a.role,mustChangePassword:!!a.mustChangePassword}})}catch(e){return json(res,400,{ok:false,error:e.message})}});
 if(u.pathname==='/api/auth/signup'&&req.method==='POST')return readBody(req).then(body=>{try{const d=JSON.parse(body||'{}'),list=readAccounts(),username=String(d.username||'').trim();if(!username||String(d.password||'').length<8||!d.company||!d.ownerName||!d.phone)return json(res,400,{ok:false,error:'거래처명·대표자·연락처·아이디와 8자 이상 비밀번호를 입력해 주세요.'});if(list.some(x=>x.username===username))return json(res,409,{ok:false,error:'이미 사용 중인 아이디입니다.'});const a={id:crypto.randomUUID(),code:newTenantCode(list),username,passwordHash:passwordHash(d.password),company:String(d.company).trim(),ownerName:String(d.ownerName).trim(),phone:onlyDigits(d.phone),role:'tenant',status:'pending',createdAt:new Date().toISOString()};list.push(a);saveAccounts(list);return json(res,200,{ok:true,code:a.code,status:a.status})}catch(e){return json(res,400,{ok:false,error:e.message})}});
 if(u.pathname==='/api/auth/forgot-password'&&req.method==='POST')return readBody(req).then(body=>{try{
  const d=JSON.parse(body||'{}'),username=String(d.username||'').trim(),company=String(d.company||'').trim(),phone=onlyDigits(d.phone),next=String(d.newPassword||'');
  if(!username||!company||!phone||next.length<8)return json(res,400,{ok:false,error:'아이디·거래처명·등록 연락처와 8자 이상 새 비밀번호를 입력해 주세요.'});
  const list=readAccounts(),a=list.find(x=>x.role==='tenant'&&x.username===username&&String(x.company||'').trim()===company&&onlyDigits(x.phone)===phone);
  if(!a)return json(res,404,{ok:false,error:'등록된 거래처 정보와 일치하지 않습니다. 최고관리자에게 확인해 주세요.'});
  a.passwordHash=passwordHash(next);a.mustChangePassword=false;a.passwordChangedAt=new Date().toISOString();a.passwordResetMethod='registered-info';saveAccounts(list);
  return json(res,200,{ok:true,message:'비밀번호가 변경되었습니다. 새 비밀번호로 로그인해 주세요.'});
 }catch(e){return json(res,400,{ok:false,error:e.message})}});
 if(u.pathname==='/api/auth/logout'&&req.method==='POST'){clearSession(req,res);return json(res,200,{ok:true})}
 if(u.pathname==='/api/auth/me'&&req.method==='GET'){const a=currentUser(req);return a?json(res,200,{ok:true,user:{id:a.id,username:a.username,company:a.company,code:a.code,role:a.role,mustChangePassword:!!a.mustChangePassword,activeTenantCode:selectedTenantCode(req)}}):json(res,401,{ok:false})}
 if(u.pathname==='/api/auth/change-password'&&req.method==='POST'){const a=currentUser(req);if(!a)return json(res,401,{ok:false,error:'로그인이 필요합니다.'});return readBody(req).then(body=>{try{const d=JSON.parse(body||'{}'),current=String(d.currentPassword||''),next=String(d.newPassword||'');if(!verifyPassword(current,a.passwordHash))return json(res,400,{ok:false,error:'현재 비밀번호가 맞지 않습니다.'});if(next.length<8)return json(res,400,{ok:false,error:'새 비밀번호는 8자 이상 입력해 주세요.'});if(current===next)return json(res,400,{ok:false,error:'현재 비밀번호와 다른 비밀번호를 입력해 주세요.'});const list=readAccounts(),target=list.find(x=>x.id===a.id);if(!target)return json(res,404,{ok:false,error:'계정을 찾을 수 없습니다.'});target.passwordHash=passwordHash(next);target.mustChangePassword=false;target.passwordChangedAt=new Date().toISOString();saveAccounts(list);return json(res,200,{ok:true})}catch(e){return json(res,400,{ok:false,error:e.message})}})}
 const publicPaths=new Set(['/login.html','/signup.html','/join.html','/packing.html','/forgot-password.html','/favicon.ico']);
 const publicApi=(u.pathname==='/api/health'||u.pathname.startsWith('/api/auth/')||u.pathname.startsWith('/api/public/packing')||(u.pathname==='/api/customers'&&req.method==='POST'&&!!u.query.tenant));
 const user=currentUser(req);
 if(!publicPaths.has(u.pathname)&&!publicApi&&!user){if(u.pathname.startsWith('/api/'))return json(res,401,{ok:false,error:'로그인이 필요합니다.'});res.writeHead(302,{Location:'/login.html'});return res.end()}
 if(u.pathname==='/api/admin/accounts'&&req.method==='GET'){if(user.role!=='superadmin')return json(res,403,{ok:false,error:'최고관리자 권한이 필요합니다.'});return json(res,200,{ok:true,accounts:readAccounts().map(({passwordHash,...a})=>a)})}
 if(u.pathname==='/api/admin/select-tenant'&&req.method==='POST'){if(user.role!=='superadmin')return json(res,403,{ok:false,error:'최고관리자 권한이 필요합니다.'});return readBody(req).then(body=>{try{const d=JSON.parse(body||'{}');if(String(d.code||'')===DDAENG_TENANT_CODE)ensureDdaengTenantAccount();const a=readAccounts().find(x=>x.code===d.code&&(x.role==='tenant'||x.role==='superadmin'));if(!a)return json(res,404,{ok:false,error:'관리페이지를 찾을 수 없습니다.'});const sid=cookies(req).ddaeng_session,ss=sessions.get(sid);if(!ss)return json(res,401,{ok:false,error:'세션이 만료되었습니다.'});ss.activeTenantCode=a.code;ss.activeTenantAt=Date.now();sessions.set(sid,ss);return json(res,200,{ok:true,tenant:{code:a.code,company:a.company,ownerName:a.ownerName}})}catch(e){return json(res,400,{ok:false,error:e.message})}})}
 if(u.pathname==='/api/admin/current-tenant'&&req.method==='GET'){const code=selectedTenantCode(req),a=readAccounts().find(x=>x.code===code);return json(res,200,{ok:true,tenant:a?{code:a.code,company:a.company,ownerName:a.ownerName}:{code,company:'거래처'},actingAs:user.role==='superadmin'})}
 if(u.pathname==='/api/admin/accounts/status'&&req.method==='POST'){if(user.role!=='superadmin')return json(res,403,{ok:false,error:'최고관리자 권한이 필요합니다.'});return readBody(req).then(body=>{try{const d=JSON.parse(body||'{}'),list=readAccounts(),a=list.find(x=>x.id===d.id);if(!a)return json(res,404,{ok:false,error:'계정을 찾을 수 없습니다.'});if(!['active','pending','suspended'].includes(d.status))return json(res,400,{ok:false,error:'상태값 오류'});a.status=d.status;saveAccounts(list);return json(res,200,{ok:true})}catch(e){return json(res,400,{ok:false,error:e.message})}})}
 if(u.pathname==='/api/admin/accounts/backup'&&req.method==='POST'){if(user.role!=='superadmin')return json(res,403,{ok:false,error:'최고관리자 권한이 필요합니다.'});try{const list=readAccounts(),stamp=new Date().toISOString().replace(/[:.]/g,'-'),file=path.join(ACCOUNT_BACKUP_DIR,`accounts-manual-${stamp}.json`);atomicWrite(file,JSON.stringify(list,null,2));return json(res,200,{ok:true,count:list.length,file:path.basename(file)})}catch(e){return json(res,500,{ok:false,error:e.message})}}
 if(u.pathname==='/api/admin/accounts/backups'&&req.method==='GET'){if(user.role!=='superadmin')return json(res,403,{ok:false,error:'최고관리자 권한이 필요합니다.'});try{const files=fs.readdirSync(ACCOUNT_BACKUP_DIR).filter(x=>x.endsWith('.json')).sort().reverse();return json(res,200,{ok:true,files})}catch(e){return json(res,500,{ok:false,error:e.message})}}

 const tenantCode=user?selectedTenantCode(req):String(u.query.tenant||'');

 if(u.pathname==='/api/youtube/status'&&req.method==='GET'){
  const cfg=youtubeConfig(),auth=readJsonObject(tenantFile(tenantCode,'youtube-auth.json'),{});
  return json(res,200,{ok:true,connected:!!(auth.refresh_token||auth.access_token||cfg.apiKey),oauth:!!(auth.refresh_token||auth.access_token),apiKey:!!cfg.apiKey})
 }
 if(u.pathname==='/api/youtube/oauth/start'&&req.method==='GET'){
  const cfg=youtubeConfig();
  if(!cfg.clientId||!cfg.redirectUri){res.writeHead(500,{'Content-Type':'text/plain; charset=utf-8'});return res.end('Render 환경변수 YOUTUBE_CLIENT_ID, YOUTUBE_REDIRECT_URI를 확인해 주세요.')}
  const state=crypto.randomBytes(20).toString('hex');
  atomicWrite(tenantFile(tenantCode,'youtube-auth.json'),JSON.stringify({...readJsonObject(tenantFile(tenantCode,'youtube-auth.json'),{}),oauth_state:state},null,2));
  const q=new URLSearchParams({
   client_id:cfg.clientId,redirect_uri:cfg.redirectUri,response_type:'code',
   scope:'https://www.googleapis.com/auth/youtube.force-ssl',
   access_type:'offline',prompt:'consent',include_granted_scopes:'true',state
  });
  res.writeHead(302,{Location:'https://accounts.google.com/o/oauth2/v2/auth?'+q.toString()});return res.end()
 }
 if(u.pathname==='/api/youtube/oauth/callback'&&req.method==='GET'){
  return (async()=>{
   try{
    const cfg=youtubeConfig(),auth=readYoutubeAuth();
    if(u.query.error)throw new Error(String(u.query.error));
    if(!u.query.code||!u.query.state||u.query.state!==auth.oauth_state)throw new Error('OAuth 상태값이 올바르지 않습니다.');
    const body=new URLSearchParams({code:String(u.query.code),client_id:cfg.clientId,client_secret:cfg.clientSecret,redirect_uri:cfg.redirectUri,grant_type:'authorization_code'});
    const r=await fetch('https://oauth2.googleapis.com/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body});
    const j=await r.json();if(!r.ok)throw new Error(j.error_description||j.error||'토큰 발급 실패');
    atomicWrite(tenantFile(tenantCode,'youtube-auth.json'),JSON.stringify({...j,refresh_token:j.refresh_token||auth.refresh_token||'',expires_at:Date.now()+(j.expires_in||3600)*1000},null,2));
    res.writeHead(200,{'Content-Type':'text/html; charset=utf-8'});
    return res.end('<!doctype html><meta charset="utf-8"><title>연결 완료</title><body style="font-family:sans-serif;padding:30px"><h2>유튜브 계정 연결 완료</h2><p>이 창은 자동으로 닫힙니다.</p><script>if(window.opener)window.opener.postMessage("youtube-connected","*");setTimeout(()=>window.close(),1000)</script></body>')
   }catch(e){res.writeHead(400,{'Content-Type':'text/html; charset=utf-8'});return res.end('<meta charset="utf-8"><h2>유튜브 연결 실패</h2><pre>'+String(e.message).replace(/[&<>]/g,s=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[s]))+'</pre>')}
  })()
 }

 if(u.pathname==='/api/youtube/message'&&req.method==='POST'){
  return readBody(req).then(async body=>{
   try{
    const data=JSON.parse(body||'{}');
    const videoId=String(data.videoId||'').trim();
    const text=String(data.text||'').trim();
    if(!videoId||!text)return json(res,400,{ok:false,error:'방송 ID 또는 댓글 내용이 없습니다.'});
    const liveChatId=await liveChatIdForVideo(videoId);
    const result=await youtubeApiPost('liveChat/messages',{part:'snippet'},{
      snippet:{
        liveChatId,
        type:'textMessageEvent',
        textMessageDetails:{messageText:text.slice(0,200)}
      }
    });
    return json(res,200,{ok:true,id:result.id||'',text});
   }catch(e){
    const status=e.status===401||/계정 연결/.test(e.message)?401:500;
    return json(res,status,{ok:false,error:e.message});
   }
  }).catch(e=>json(res,400,{ok:false,error:e.message}))
 }

 if(u.pathname==='/api/youtube/comments'&&req.method==='GET'){
  return (async()=>{
   try{
    const videoId=String(u.query.videoId||'').trim();
    if(!videoId)return json(res,400,{ok:false,error:'videoId가 없습니다.'});
    const liveChatId=await liveChatIdForVideo(videoId);
    const params={liveChatId,part:'id,snippet,authorDetails',maxResults:'200'};
    if(u.query.pageToken)params.pageToken=String(u.query.pageToken);
    const j=await youtubeApi('liveChat/messages',params);
    const messages=(j.items||[]).filter(x=>x.snippet?.type==='textMessageEvent').map(x=>({
      id:x.id,author:x.authorDetails?.displayName||'유튜브고객',
      authorChannelId:x.authorDetails?.channelId||'',
      text:x.snippet?.displayMessage||x.snippet?.textMessageDetails?.messageText||'',
      publishedAt:x.snippet?.publishedAt||'',
      isOwner:!!x.authorDetails?.isChatOwner,
      isModerator:!!x.authorDetails?.isChatModerator
    }));
    return json(res,200,{ok:true,liveChatId,messages,nextPageToken:j.nextPageToken||'',pollingIntervalMillis:j.pollingIntervalMillis||3500})
   }catch(e){
    const status=e.status===401||/연결|API_KEY/.test(e.message)?401:500;
    return json(res,status,{ok:false,error:e.message})
   }
  })()
 }

 if(u.pathname==='/api/solapi/config'&&req.method==='GET'){
  const cfg=solapiConfig(tenantCode);
  return json(res,200,{ok:true,configured:!!(cfg.apiKey&&cfg.apiSecret&&cfg.sender),apiKey:cfg.apiKey?cfg.apiKey.slice(0,4)+'••••••':'',sender:cfg.sender||'',pfId:cfg.pfId||'',templateId:cfg.templateId||'',hasSecret:!!cfg.apiSecret});
 }
 if(u.pathname==='/api/solapi/config'&&req.method==='POST'){
  return readBody(req).then(body=>{try{
   const d=JSON.parse(body||'{}'), old=tenantReadIntegrations(tenantCode);
   const next={apiKey:String(d.apiKey||old.apiKey||'').trim(),apiSecret:String(d.apiSecret||old.apiSecret||'').trim(),sender:onlyDigits(d.sender||old.sender),pfId:String(d.pfId||old.pfId||'').trim(),templateId:String(d.templateId||old.templateId||'').trim()};
   if(!next.apiKey||!next.apiSecret||!next.sender)return json(res,400,{ok:false,error:'API Key, API Secret, 승인 발신번호를 모두 입력해 주세요.'});
   tenantSaveIntegrations(tenantCode,next);return json(res,200,{ok:true,configured:true,sender:next.sender});
  }catch(e){return json(res,400,{ok:false,error:e.message})}});
 }
 if(u.pathname==='/api/customers/export.xlsx'&&req.method==='GET'){
  const list=tenantReadCustomers(tenantCode);const tf=tenantFile(tenantCode,'customers.xlsx');try{const wb=XLSX.utils.book_new(),ws=XLSX.utils.json_to_sheet(list);XLSX.utils.book_append_sheet(wb,ws,'고객DB');XLSX.writeFile(wb,tf)}catch(e){}
  if(!fs.existsSync(tf))return json(res,500,{ok:false,error:'고객 엑셀 생성 실패'});
  res.writeHead(200,{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','Content-Disposition':'attachment; filename="FIRST_OMS_customers.xlsx"','Cache-Control':'no-store'});return fs.createReadStream(tf).pipe(res);
 }
 if(u.pathname==='/api/kakao/status'&&req.method==='GET'){
  const cfg=solapiConfig(tenantCode);
  return json(res,200,{ok:true,ready:!!(cfg.apiKey&&cfg.apiSecret&&cfg.sender&&cfg.pfId&&cfg.templateId),sender:cfg.sender?cfg.sender.slice(0,3)+'****'+cfg.sender.slice(-4):'',pfId:cfg.pfId?cfg.pfId.slice(0,6)+'…':'',templateId:cfg.templateId?cfg.templateId.slice(0,6)+'…':'',missing:[!cfg.apiKey&&'SOLAPI_API_KEY',!cfg.apiSecret&&'SOLAPI_API_SECRET',!cfg.sender&&'SOLAPI_SENDER',!cfg.pfId&&'SOLAPI_KAKAO_PF_ID',!cfg.templateId&&'SOLAPI_KAKAO_TEMPLATE_ID'].filter(Boolean)})
 }
 if(u.pathname==='/api/kakao/send'&&req.method==='POST'){
  return readBody(req).then(async body=>{
   try{
    const data=JSON.parse(body||'{}');
    if(!data.to)return json(res,400,{ok:false,error:'수신번호가 없습니다.'});
    const result=await sendSolapiKakao(data.to,data.variables||{},data.text||'',tenantCode);
    json(res,200,{ok:true,result})
   }catch(e){json(res,500,{ok:false,error:e.message})}
  }).catch(e=>json(res,400,{ok:false,error:e.message}))
 }


 if(u.pathname==='/api/public/packing'&&req.method==='GET'){
  const code=String(u.query.code||'').trim().toUpperCase();let found=null;const tenantCodes=readAccounts().filter(a=>a.role==='tenant'||a.role==='superadmin').map(a=>a.code);
  for(const tc of tenantCodes){const st=tenantReadState(tc),customers=Array.isArray(st.customers)?st.customers:tenantReadCustomers(tc),groups=new Map();for(const o of (st.orders||[])){const k=(o.customerId||o.nick)+'|'+(o.date||'');if(!groups.has(k))groups.set(k,{name:'',nick:o.nick||'',phone:'',address:'',dates:new Set(),items:[],subtotal:0,fee:Number(o.fee)||0,total:0,code:'',tenantCode:tc});const g=groups.get(k),c=customers.find(x=>x.id===o.customerId)||{};g.name=c.name||'';g.phone=c.phone||'';g.address=[c.postalCode,c.address,c.detailAddress].filter(Boolean).join(' ');g.dates.add(o.date||'');g.items.push({number:o.number||g.items.length+1,item:o.item||'',qty:Number(o.qty)||0,unit:Number(o.unit)||0,amount:Number(o.amount)||0});g.subtotal+=Number(o.amount)||0}function sc(g){const raw=[g.name,g.nick,g.phone,g.address].join('|');let h=2166136261;for(let i=0;i<raw.length;i++){h^=raw.charCodeAt(i);h=Math.imul(h,16777619)}return 'FIRST-'+(h>>>0).toString(36).toUpperCase().padStart(7,'0')}found=[...groups.values()].map(g=>({...g,dates:[...g.dates],total:g.subtotal+g.fee,code:sc(g)})).find(g=>g.code===code);if(found){found.completedAt=st.shippingScans?.[code]?.at||null;found.worker=st.shippingScans?.[code]?.worker||'';found.trackingNumber=st.shippingScans?.[code]?.trackingNumber||'';found.trackingUpdatedAt=st.shippingScans?.[code]?.trackingUpdatedAt||null;break}}
  if(!found)return json(res,404,{ok:false,error:'포장리스트를 찾을 수 없습니다.'});return json(res,200,{ok:true,job:found})
 }
 if(u.pathname==='/api/public/packing/complete'&&req.method==='POST')return readBody(req).then(body=>{try{const d=JSON.parse(body||'{}'),code=String(d.code||'').trim().toUpperCase(),tc=String(d.tenantCode||'');if(!code||!tc)return json(res,400,{ok:false,error:'QR 코드 또는 거래처 코드가 없습니다.'});const st=tenantReadState(tc);st.shippingScans=st.shippingScans||{};const completedAt=new Date().toISOString();const prev=st.shippingScans[code]||{};st.shippingScans[code]={...prev,at:completedAt,source:'mobile-public',worker:String(d.worker||'').trim()};tenantWriteState(tc,st);return json(res,200,{ok:true,completedAt})}catch(e){return json(res,400,{ok:false,error:e.message})}})
 if(u.pathname==='/api/public/packing/tracking'&&req.method==='POST')return readBody(req).then(body=>{try{const d=JSON.parse(body||'{}'),code=String(d.code||'').trim().toUpperCase(),tc=String(d.tenantCode||''),trackingNumber=String(d.trackingNumber||'').replace(/\s+/g,'').trim();if(!code||!tc)return json(res,400,{ok:false,error:'QR 코드 또는 거래처 코드가 없습니다.'});if(!trackingNumber)return json(res,400,{ok:false,error:'송장번호를 입력하거나 스캔해 주세요.'});if(trackingNumber.length<6||trackingNumber.length>40)return json(res,400,{ok:false,error:'송장번호 형식을 확인해 주세요.'});const st=tenantReadState(tc);st.shippingScans=st.shippingScans||{};const prev=st.shippingScans[code]||{};const trackingUpdatedAt=new Date().toISOString();st.shippingScans[code]={...prev,trackingNumber,trackingUpdatedAt};tenantWriteState(tc,st);return json(res,200,{ok:true,trackingNumber,trackingUpdatedAt})}catch(e){return json(res,400,{ok:false,error:e.message})}})


 if(u.pathname==='/api/packing/tracking/bulk'&&req.method==='POST')return readBody(req,5*1024*1024).then(body=>{try{const d=JSON.parse(body||'{}'),updates=Array.isArray(d.updates)?d.updates:[];if(!updates.length)return json(res,400,{ok:false,error:'등록할 송장정보가 없습니다.'});const st=tenantReadState(tenantCode);st.shippingScans=st.shippingScans||{};let updated=0,skipped=0;for(const u of updates){const code=String(u.code||'').trim().toUpperCase(),trackingNumber=String(u.trackingNumber||'').replace(/\s+/g,'').trim(),courier=String(u.courier||'파일접수').trim();if(!code||!trackingNumber){skipped++;continue}const prev=st.shippingScans[code]||{};if(prev.trackingNumber&&prev.trackingNumber!==trackingNumber){skipped++;continue}st.shippingScans[code]={...prev,trackingNumber,courier,trackingUpdatedAt:new Date().toISOString(),trackingSource:'file-upload'};updated++}tenantWriteState(tenantCode,st);return json(res,200,{ok:true,updated,skipped})}catch(e){return json(res,400,{ok:false,error:e.message})}})
 if(u.pathname==='/api/packing/scans'&&req.method==='GET'){const st=tenantReadState(tenantCode);return json(res,200,{ok:true,shippingScans:st.shippingScans||{}})}
 if(u.pathname==='/api/packing/status'&&req.method==='POST')return readBody(req).then(body=>{try{const d=JSON.parse(body||'{}'),code=String(d.code||'').trim().toUpperCase();if(!code)return json(res,400,{ok:false,error:'택배코드가 없습니다.'});const st=tenantReadState(tenantCode);st.shippingScans=st.shippingScans||{};const prev=st.shippingScans[code]||{};st.shippingScans[code]={...prev,at:new Date().toISOString(),source:String(d.source||'admin-manual'),worker:String(d.worker||'관리자')};tenantWriteState(tenantCode,st);return json(res,200,{ok:true,scan:st.shippingScans[code]})}catch(e){return json(res,400,{ok:false,error:e.message})}})
 if(u.pathname==='/api/packing/status'&&req.method==='DELETE'){const code=String(u.query.code||'').trim().toUpperCase();if(!code)return json(res,400,{ok:false,error:'택배코드가 없습니다.'});const st=tenantReadState(tenantCode);st.shippingScans=st.shippingScans||{};delete st.shippingScans[code];tenantWriteState(tenantCode,st);return json(res,200,{ok:true})}
 if(u.pathname==='/api/state'&&req.method==='GET')return json(res,200,{ok:true,state:tenantReadState(tenantCode),archives:tenantListSalesArchives(tenantCode),tenantCode});
 if(u.pathname==='/api/state'&&req.method==='POST'){
  return readBody(req,20*1024*1024).then(body=>{try{
   const st=JSON.parse(body||'{}'),current=tenantReadState(tenantCode);
   // 전체복원/자동저장에 고객 배열이 포함되면 customers.json을 먼저 갱신한다.
   // SOLAPI 설정(solapi-settings.json)은 이 경로에서 읽거나 쓰지 않는다.
   if(Array.isArray(st.customers))tenantWriteCustomers(tenantCode,st.customers);
   st.shippingScans=current.shippingScans||st.shippingScans||{};
   const saved=tenantWriteState(tenantCode,st);
   return json(res,200,{ok:true,updatedAt:saved.updatedAt,orders:(saved.orders||[]).length,customers:(saved.customers||[]).length})
  }catch(e){return json(res,400,{ok:false,error:e.message})}})
 }
 if(u.pathname==='/api/state/backup'&&req.method==='GET'){
  const st=tenantReadState(tenantCode);const payload={version:7.5,tenantCode,exportedAt:new Date().toISOString(),state:st};
  res.writeHead(200,{'Content-Type':'application/json; charset=utf-8','Content-Disposition':`attachment; filename=ddaenglive_full_backup_${new Date().toISOString().slice(0,10)}.json`,'Cache-Control':'no-store'});return res.end(JSON.stringify(payload,null,2));
 }
 if(u.pathname==='/api/sales/archives'&&req.method==='GET')return json(res,200,{ok:true,archives:tenantListSalesArchives(tenantCode)});
 if(u.pathname.startsWith('/api/sales/archive/')&&req.method==='GET'){
  const date=decodeURIComponent(u.pathname.split('/').pop()).replace(/[^0-9-]/g,'');const f=path.join(tenantSalesArchiveDir(tenantCode),date+'.json');
  if(!fs.existsSync(f))return json(res,404,{ok:false,error:'해당 날짜 판매리스트가 없습니다.'});return json(res,200,{ok:true,...JSON.parse(fs.readFileSync(f,'utf8'))});
 }
 if(u.pathname==='/api/send-history'&&req.method==='GET'){
  const date=String(u.query.date||'');const history=tenantReadSendHistory(tenantCode).filter(x=>!date||x.date===date);return json(res,200,{ok:true,history});
 }
 if(u.pathname==='/api/mms/send'&&req.method==='POST'){
  return readBody(req,1024*1024).then(async body=>{let meta={};try{const d=JSON.parse(body||'{}');meta={sentAt:new Date().toISOString(),date:String(d.date||''),nickname:String(d.nickname||''),name:String(d.name||''),toMasked:onlyDigits(d.to).replace(/^(\d{3})\d+(\d{4})$/,'$1****$2'),total:Number(d.total)||0};if(!d.to||!d.imageBase64)return json(res,400,{ok:false,error:'수신번호 또는 정산서 이미지가 없습니다.'});const result=await sendSolapiMms(d.to,d.imageBase64,d.subject,d.text,tenantCode);tenantAppendSendHistory(tenantCode,{...meta,ok:true});return json(res,200,{ok:true,result})}catch(e){tenantAppendSendHistory(tenantCode,{...meta,ok:false,error:e.message});return json(res,500,{ok:false,error:e.message})}}).catch(e=>json(res,400,{ok:false,error:e.message}))
 }
 if(u.pathname==='/api/sms/send'&&req.method==='POST'){
  return readBody(req).then(async body=>{
   try{
    const data=JSON.parse(body||'{}');
    if(!data.to||!data.text)return json(res,400,{ok:false,error:'수신번호 또는 문자내용이 없습니다.'});
    const result=await sendSolapiSms(data.to,data.text,tenantCode);
    json(res,200,{ok:true,result})
   }catch(e){json(res,500,{ok:false,error:e.message})}
  }).catch(e=>json(res,400,{ok:false,error:e.message}))
 }

 if(u.pathname==='/api/customers/backup'&&req.method==='GET'){
  const list=tenantReadCustomers(tenantCode);const payload={version:2,tenantCode,exportedAt:new Date().toISOString(),count:list.length,customers:list};
  res.writeHead(200,{
   'Content-Type':'application/json; charset=utf-8',
   'Content-Disposition':`attachment; filename=FIRST_OMS_customers_backup_${new Date().toISOString().slice(0,10)}.json`,
   'Cache-Control':'no-store',
   'X-Backup-Count':String(list.length)
  });
  return res.end(JSON.stringify(payload,null,2));
 }
 if(u.pathname==='/api/customers/backup/status'&&req.method==='GET'){
  return json(res,200,{ok:true,count:tenantReadCustomers(tenantCode).length,backupExists:fs.existsSync(tenantFile(tenantCode,'customers-backup.json'))});
 }

 if(u.pathname==='/api/customers/stream'&&req.method==='GET'){
  res.writeHead(200,{
   'Content-Type':'text/event-stream; charset=utf-8',
   'Cache-Control':'no-cache, no-transform',
   'Connection':'keep-alive',
   'X-Accel-Buffering':'no'
  });
  res.write(`event: customers\ndata: ${JSON.stringify(tenantReadCustomers(tenantCode))}\n\n`);
  sseClients.set(res,tenantCode);
  const keep=setInterval(()=>{try{res.write(': keepalive\n\n')}catch(e){}},15000);
  req.on('close',()=>{clearInterval(keep);sseClients.delete(res)});
  return;
 }
 if(u.pathname==='/api/customers'&&req.method==='GET')return json(res,200,tenantReadCustomers(tenantCode));

 if(u.pathname.startsWith('/api/customers/')&&req.method==='DELETE'){
  const nickname=decodeURIComponent(u.pathname.split('/').pop());
  let list=tenantReadCustomers(tenantCode).map(x=>x.nickname===nickname?{...x,active:false,archivedAt:new Date().toISOString()}:x);
  tenantWriteCustomers(tenantCode,list);const st=tenantReadState(tenantCode);st.customers=list;tenantWriteState(tenantCode,st);return json(res,200,{ok:true,softDeleted:true})
 }

 if(u.pathname==='/api/customers'&&req.method==='POST'){
  let body='';req.on('data',d=>{body+=d;if(body.length>1e6)req.destroy()});
  return req.on('end',()=>{try{
   const c=JSON.parse(body||'{}');
   if(!c.name||!c.nickname||!c.phone)return json(res,400,{error:'필수값 누락'});
   let list=tenantReadCustomers(tenantCode),i=list.findIndex(x=>x.nickname===c.nickname||x.phone===c.phone);
   const next={id:c.id||Date.now().toString(36),joinedAt:c.joinedAt||new Date().toLocaleString('ko-KR'),source:c.source||'가입폼',...c};
   if(i>=0)list[i]={...list[i],...next};else list.push(next);
   tenantWriteCustomers(tenantCode,list);const st=tenantReadState(tenantCode);st.customers=list;atomicWrite(tenantFile(tenantCode,'server-state.json'),JSON.stringify({...st,customers:list,updatedAt:new Date().toISOString()},null,2));json(res,200,{...next,totalCustomers:list.length})
  }catch(e){json(res,400,{error:'잘못된 요청'})}})
 }
 let p=u.pathname==='/'?'/index.html':u.pathname;
 p=path.normalize(p).replace(/^(\.\.[\/\\])+/, '');
 let f=path.join(ROOT,p);
 if(!f.startsWith(ROOT))return res.end('forbidden');
 fs.readFile(f,(e,data)=>{if(e){res.writeHead(404);return res.end('Not found')}res.writeHead(200,{'Content-Type':mime[path.extname(f).toLowerCase()]||'application/octet-stream'});res.end(data)})
});
const PORT=process.env.PORT||3010;
server.listen(PORT,'0.0.0.0',()=>console.log(`FIRST OMS emergency server: http://localhost:${PORT}`));