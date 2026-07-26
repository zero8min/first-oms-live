const http=require('http'),fs=require('fs'),path=require('path'),url=require('url'),crypto=require('crypto'),XLSX=require('xlsx');
const sseClients=new Set();
function broadcastCustomers(list){
 const payload=`event: customers\ndata: ${JSON.stringify(list)}\n\n`;
 for(const res of [...sseClients]){try{res.write(payload)}catch(e){sseClients.delete(res)}}
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
  'cs-history.json':'[]'
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
const DEFAULT_ADMIN_ID='firstadmin',DEFAULT_ADMIN_PASSWORD='12345678';
function ensureOwnerAccount(){
 let list=readAccounts(),changed=false;
 let owner=list.find(a=>a.role==='superadmin'||a.code==='FIRST-MASTER');
 if(!owner){
  owner={id:crypto.randomUUID(),code:'FIRST-MASTER',username:DEFAULT_ADMIN_ID,passwordHash:passwordHash(DEFAULT_ADMIN_PASSWORD),company:'FIRST OMS',ownerName:'최고관리자',phone:'',role:'superadmin',status:'active',mustChangePassword:true,createdAt:new Date().toISOString(),bootstrapVersion:'7.3'};
  list.push(owner);changed=true;
 }else{
  // 아직 최초 비밀번호를 변경하지 않은 관리자 계정은 배포 후에도 기본 로그인값으로 확실히 복구한다.
  // 사용자가 비밀번호를 변경해 passwordChangedAt이 생긴 뒤에는 절대 덮어쓰지 않는다.
  if(!owner.passwordChangedAt){
   if(owner.username!==DEFAULT_ADMIN_ID){owner.username=DEFAULT_ADMIN_ID;changed=true}
   if(!verifyPassword(DEFAULT_ADMIN_PASSWORD,owner.passwordHash)){owner.passwordHash=passwordHash(DEFAULT_ADMIN_PASSWORD);changed=true}
   if(owner.mustChangePassword!==true){owner.mustChangePassword=true;changed=true}
  }
  if(owner.role!=='superadmin'){owner.role='superadmin';changed=true}
  if(owner.status!=='active'){owner.status='active';changed=true}
  if(owner.code!=='FIRST-MASTER'){owner.code='FIRST-MASTER';changed=true}
 }
 if(changed)saveAccounts(list);else for(const a of list)ensureTenantStorage(a);
 console.log('[LOGIN] 최고관리자 계정 준비 완료');
}
function cookies(req){return Object.fromEntries(String(req.headers.cookie||'').split(';').map(x=>x.trim()).filter(Boolean).map(x=>{const i=x.indexOf('=');return [decodeURIComponent(x.slice(0,i)),decodeURIComponent(x.slice(i+1))]}))}
function currentUser(req){const sid=cookies(req).ddaeng_session,ss=sessions.get(sid);if(!ss||ss.expiresAt<Date.now()){if(sid)sessions.delete(sid);return null}return readAccounts().find(a=>a.id===ss.userId&&a.status==='active')||null}
function issueSession(req,res,user){const sid=crypto.randomBytes(32).toString('hex');sessions.set(sid,{userId:user.id,expiresAt:Date.now()+1000*60*60*24*7});const secure=String(req.headers['x-forwarded-proto']||'').includes('https')?'; Secure':'';res.setHeader('Set-Cookie',`ddaeng_session=${sid}; Path=/; HttpOnly; SameSite=Lax; Max-Age=604800${secure}`)}
function clearSession(req,res){const sid=cookies(req).ddaeng_session;if(sid)sessions.delete(sid);res.setHeader('Set-Cookie','ddaeng_session=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0')}
function newTenantCode(list){let n=1;const used=new Set(list.map(a=>a.code));while(used.has(`FIRST-${String(n).padStart(4,'0')}`))n++;return `FIRST-${String(n).padStart(4,'0')}`}
ensureOwnerAccount();

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
async function sendSolapiSms(to,text){
 const cfg=solapiConfig(),apiKey=cfg.apiKey,apiSecret=cfg.apiSecret,sender=cfg.sender;
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


function solapiConfig(){
 const f=readIntegrations();
 return {
  apiKey:f.apiKey||process.env.SOLAPI_API_KEY||'',
  apiSecret:f.apiSecret||process.env.SOLAPI_API_SECRET||'',
  sender:onlyDigits(f.sender||process.env.SOLAPI_SENDER),
  pfId:f.pfId||process.env.SOLAPI_KAKAO_PF_ID||process.env.SOLAPI_PF_ID||'',
  templateId:f.templateId||process.env.SOLAPI_KAKAO_TEMPLATE_ID||process.env.SOLAPI_TEMPLATE_ID||''
 }
}
async function sendSolapiKakao(to,variables,text){
 const cfg=solapiConfig();
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
async function sendSolapiMms(to,imageBase64,subject,text){
 const cfg=solapiConfig();if(!cfg.apiKey||!cfg.apiSecret||!cfg.sender)throw new Error('SOLAPI API Key·Secret·승인 발신번호를 확인해 주세요.');
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
 const cfg=youtubeConfig(),auth=readYoutubeAuth();
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
 if(u.pathname==='/api/auth/logout'&&req.method==='POST'){clearSession(req,res);return json(res,200,{ok:true})}
 if(u.pathname==='/api/auth/me'&&req.method==='GET'){const a=currentUser(req);return a?json(res,200,{ok:true,user:{id:a.id,username:a.username,company:a.company,code:a.code,role:a.role,mustChangePassword:!!a.mustChangePassword}}):json(res,401,{ok:false})}
 if(u.pathname==='/api/auth/change-password'&&req.method==='POST'){const a=currentUser(req);if(!a)return json(res,401,{ok:false,error:'로그인이 필요합니다.'});return readBody(req).then(body=>{try{const d=JSON.parse(body||'{}'),current=String(d.currentPassword||''),next=String(d.newPassword||'');if(!verifyPassword(current,a.passwordHash))return json(res,400,{ok:false,error:'현재 비밀번호가 맞지 않습니다.'});if(next.length<8)return json(res,400,{ok:false,error:'새 비밀번호는 8자 이상 입력해 주세요.'});if(current===next)return json(res,400,{ok:false,error:'현재 비밀번호와 다른 비밀번호를 입력해 주세요.'});const list=readAccounts(),target=list.find(x=>x.id===a.id);if(!target)return json(res,404,{ok:false,error:'계정을 찾을 수 없습니다.'});target.passwordHash=passwordHash(next);target.mustChangePassword=false;target.passwordChangedAt=new Date().toISOString();saveAccounts(list);return json(res,200,{ok:true})}catch(e){return json(res,400,{ok:false,error:e.message})}})}
 const publicPaths=new Set(['/login.html','/signup.html','/join.html','/favicon.ico']);
 const publicApi=(u.pathname==='/api/health'||u.pathname.startsWith('/api/auth/')||(u.pathname==='/api/customers'&&req.method==='POST'));
 const user=currentUser(req);
 if(!publicPaths.has(u.pathname)&&!publicApi&&!user){if(u.pathname.startsWith('/api/'))return json(res,401,{ok:false,error:'로그인이 필요합니다.'});res.writeHead(302,{Location:'/login.html'});return res.end()}
 if(u.pathname==='/api/admin/accounts'&&req.method==='GET'){if(user.role!=='superadmin')return json(res,403,{ok:false,error:'최고관리자 권한이 필요합니다.'});return json(res,200,{ok:true,accounts:readAccounts().map(({passwordHash,...a})=>a)})}
 if(u.pathname==='/api/admin/accounts/status'&&req.method==='POST'){if(user.role!=='superadmin')return json(res,403,{ok:false,error:'최고관리자 권한이 필요합니다.'});return readBody(req).then(body=>{try{const d=JSON.parse(body||'{}'),list=readAccounts(),a=list.find(x=>x.id===d.id);if(!a)return json(res,404,{ok:false,error:'계정을 찾을 수 없습니다.'});if(!['active','pending','suspended'].includes(d.status))return json(res,400,{ok:false,error:'상태값 오류'});a.status=d.status;saveAccounts(list);return json(res,200,{ok:true})}catch(e){return json(res,400,{ok:false,error:e.message})}})}
 if(u.pathname==='/api/admin/accounts/backup'&&req.method==='POST'){if(user.role!=='superadmin')return json(res,403,{ok:false,error:'최고관리자 권한이 필요합니다.'});try{const list=readAccounts(),stamp=new Date().toISOString().replace(/[:.]/g,'-'),file=path.join(ACCOUNT_BACKUP_DIR,`accounts-manual-${stamp}.json`);atomicWrite(file,JSON.stringify(list,null,2));return json(res,200,{ok:true,count:list.length,file:path.basename(file)})}catch(e){return json(res,500,{ok:false,error:e.message})}}
 if(u.pathname==='/api/admin/accounts/backups'&&req.method==='GET'){if(user.role!=='superadmin')return json(res,403,{ok:false,error:'최고관리자 권한이 필요합니다.'});try{const files=fs.readdirSync(ACCOUNT_BACKUP_DIR).filter(x=>x.endsWith('.json')).sort().reverse();return json(res,200,{ok:true,files})}catch(e){return json(res,500,{ok:false,error:e.message})}}


 if(u.pathname==='/api/youtube/status'&&req.method==='GET'){
  const cfg=youtubeConfig(),auth=readYoutubeAuth();
  return json(res,200,{ok:true,connected:!!(auth.refresh_token||auth.access_token||cfg.apiKey),oauth:!!(auth.refresh_token||auth.access_token),apiKey:!!cfg.apiKey})
 }
 if(u.pathname==='/api/youtube/oauth/start'&&req.method==='GET'){
  const cfg=youtubeConfig();
  if(!cfg.clientId||!cfg.redirectUri){res.writeHead(500,{'Content-Type':'text/plain; charset=utf-8'});return res.end('Render 환경변수 YOUTUBE_CLIENT_ID, YOUTUBE_REDIRECT_URI를 확인해 주세요.')}
  const state=crypto.randomBytes(20).toString('hex');
  saveYoutubeAuth({...readYoutubeAuth(),oauth_state:state});
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
    saveYoutubeAuth({...j,refresh_token:j.refresh_token||auth.refresh_token||'',expires_at:Date.now()+(j.expires_in||3600)*1000});
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
  const cfg=solapiConfig();
  return json(res,200,{ok:true,configured:!!(cfg.apiKey&&cfg.apiSecret&&cfg.sender),apiKey:cfg.apiKey?cfg.apiKey.slice(0,4)+'••••••':'',sender:cfg.sender||'',pfId:cfg.pfId||'',templateId:cfg.templateId||'',hasSecret:!!cfg.apiSecret});
 }
 if(u.pathname==='/api/solapi/config'&&req.method==='POST'){
  return readBody(req).then(body=>{try{
   const d=JSON.parse(body||'{}'), old=readIntegrations();
   const next={apiKey:String(d.apiKey||old.apiKey||'').trim(),apiSecret:String(d.apiSecret||old.apiSecret||'').trim(),sender:onlyDigits(d.sender||old.sender),pfId:String(d.pfId||old.pfId||'').trim(),templateId:String(d.templateId||old.templateId||'').trim()};
   if(!next.apiKey||!next.apiSecret||!next.sender)return json(res,400,{ok:false,error:'API Key, API Secret, 승인 발신번호를 모두 입력해 주세요.'});
   saveIntegrations(next);return json(res,200,{ok:true,configured:true,sender:next.sender});
  }catch(e){return json(res,400,{ok:false,error:e.message})}});
 }
 if(u.pathname==='/api/customers/export.xlsx'&&req.method==='GET'){
  writeCustomerExcel(readCustomers());
  if(!fs.existsSync(CUSTOMER_XLSX))return json(res,500,{ok:false,error:'고객 엑셀 생성 실패'});
  res.writeHead(200,{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','Content-Disposition':'attachment; filename="FIRST_OMS_customers.xlsx"','Cache-Control':'no-store'});return fs.createReadStream(CUSTOMER_XLSX).pipe(res);
 }
 if(u.pathname==='/api/kakao/status'&&req.method==='GET'){
  const cfg=solapiConfig();
  return json(res,200,{ok:true,ready:!!(cfg.apiKey&&cfg.apiSecret&&cfg.sender&&cfg.pfId&&cfg.templateId),sender:cfg.sender?cfg.sender.slice(0,3)+'****'+cfg.sender.slice(-4):'',pfId:cfg.pfId?cfg.pfId.slice(0,6)+'…':'',templateId:cfg.templateId?cfg.templateId.slice(0,6)+'…':'',missing:[!cfg.apiKey&&'SOLAPI_API_KEY',!cfg.apiSecret&&'SOLAPI_API_SECRET',!cfg.sender&&'SOLAPI_SENDER',!cfg.pfId&&'SOLAPI_KAKAO_PF_ID',!cfg.templateId&&'SOLAPI_KAKAO_TEMPLATE_ID'].filter(Boolean)})
 }
 if(u.pathname==='/api/kakao/send'&&req.method==='POST'){
  return readBody(req).then(async body=>{
   try{
    const data=JSON.parse(body||'{}');
    if(!data.to)return json(res,400,{ok:false,error:'수신번호가 없습니다.'});
    const result=await sendSolapiKakao(data.to,data.variables||{},data.text||'');
    json(res,200,{ok:true,result})
   }catch(e){json(res,500,{ok:false,error:e.message})}
  }).catch(e=>json(res,400,{ok:false,error:e.message}))
 }

 if(u.pathname==='/api/state'&&req.method==='GET')return json(res,200,{ok:true,state:readState(),archives:listSalesArchives()});
 if(u.pathname==='/api/state'&&req.method==='POST'){
  return readBody(req,20*1024*1024).then(body=>{try{const st=JSON.parse(body||'{}');const saved=saveState(st);return json(res,200,{ok:true,updatedAt:saved.updatedAt,orders:(saved.orders||[]).length,customers:(saved.customers||[]).length})}catch(e){return json(res,400,{ok:false,error:e.message})}})
 }
 if(u.pathname==='/api/state/backup'&&req.method==='GET'){
  const st=readState();const payload={version:6,exportedAt:new Date().toISOString(),state:st};
  res.writeHead(200,{'Content-Type':'application/json; charset=utf-8','Content-Disposition':`attachment; filename=ddaenglive_full_backup_${new Date().toISOString().slice(0,10)}.json`,'Cache-Control':'no-store'});return res.end(JSON.stringify(payload,null,2));
 }
 if(u.pathname==='/api/sales/archives'&&req.method==='GET')return json(res,200,{ok:true,archives:listSalesArchives()});
 if(u.pathname.startsWith('/api/sales/archive/')&&req.method==='GET'){
  const date=decodeURIComponent(u.pathname.split('/').pop()).replace(/[^0-9-]/g,'');const f=path.join(SALES_ARCHIVE_DIR,date+'.json');
  if(!fs.existsSync(f))return json(res,404,{ok:false,error:'해당 날짜 판매리스트가 없습니다.'});return json(res,200,{ok:true,...JSON.parse(fs.readFileSync(f,'utf8'))});
 }
 if(u.pathname==='/api/send-history'&&req.method==='GET'){
  const date=String(u.query.date||'');const history=readSendHistory().filter(x=>!date||x.date===date);return json(res,200,{ok:true,history});
 }
 if(u.pathname==='/api/mms/send'&&req.method==='POST'){
  return readBody(req,1024*1024).then(async body=>{let meta={};try{const d=JSON.parse(body||'{}');meta={sentAt:new Date().toISOString(),date:String(d.date||''),nickname:String(d.nickname||''),name:String(d.name||''),toMasked:onlyDigits(d.to).replace(/^(\d{3})\d+(\d{4})$/,'$1****$2'),total:Number(d.total)||0};if(!d.to||!d.imageBase64)return json(res,400,{ok:false,error:'수신번호 또는 정산서 이미지가 없습니다.'});const result=await sendSolapiMms(d.to,d.imageBase64,d.subject,d.text);appendSendHistory({...meta,ok:true});return json(res,200,{ok:true,result})}catch(e){appendSendHistory({...meta,ok:false,error:e.message});return json(res,500,{ok:false,error:e.message})}}).catch(e=>json(res,400,{ok:false,error:e.message}))
 }
 if(u.pathname==='/api/sms/send'&&req.method==='POST'){
  return readBody(req).then(async body=>{
   try{
    const data=JSON.parse(body||'{}');
    if(!data.to||!data.text)return json(res,400,{ok:false,error:'수신번호 또는 문자내용이 없습니다.'});
    const result=await sendSolapiSms(data.to,data.text);
    json(res,200,{ok:true,result})
   }catch(e){json(res,500,{ok:false,error:e.message})}
  }).catch(e=>json(res,400,{ok:false,error:e.message}))
 }

 if(u.pathname==='/api/customers/backup'&&req.method==='GET'){
  const info=backupCustomers();
  const payload={version:1,exportedAt:new Date().toISOString(),count:readCustomers().length,customers:readCustomers()};
  res.writeHead(200,{
   'Content-Type':'application/json; charset=utf-8',
   'Content-Disposition':`attachment; filename=FIRST_OMS_customers_backup_${new Date().toISOString().slice(0,10)}.json`,
   'Cache-Control':'no-store',
   'X-Backup-Count':String(info.count||0)
  });
  return res.end(JSON.stringify(payload,null,2));
 }
 if(u.pathname==='/api/customers/backup/status'&&req.method==='GET'){
  return json(res,200,{ok:true,count:readCustomers().length,backupExists:fs.existsSync(CUSTOMER_BACKUP)});
 }

 if(u.pathname==='/api/customers/stream'&&req.method==='GET'){
  res.writeHead(200,{
   'Content-Type':'text/event-stream; charset=utf-8',
   'Cache-Control':'no-cache, no-transform',
   'Connection':'keep-alive',
   'X-Accel-Buffering':'no'
  });
  res.write(`event: customers\ndata: ${JSON.stringify(readCustomers())}\n\n`);
  sseClients.add(res);
  const keep=setInterval(()=>{try{res.write(': keepalive\n\n')}catch(e){}},15000);
  req.on('close',()=>{clearInterval(keep);sseClients.delete(res)});
  return;
 }
 if(u.pathname==='/api/customers'&&req.method==='GET')return json(res,200,readCustomers());

 if(u.pathname.startsWith('/api/customers/')&&req.method==='DELETE'){
  const nickname=decodeURIComponent(u.pathname.split('/').pop());
  let list=readCustomers().map(x=>x.nickname===nickname?{...x,active:false,archivedAt:new Date().toISOString()}:x);
  saveCustomers(list);const st=readState();st.customers=list;saveState(st);return json(res,200,{ok:true,softDeleted:true})
 }

 if(u.pathname==='/api/customers'&&req.method==='POST'){
  let body='';req.on('data',d=>{body+=d;if(body.length>1e6)req.destroy()});
  return req.on('end',()=>{try{
   const c=JSON.parse(body||'{}');
   if(!c.name||!c.nickname||!c.phone)return json(res,400,{error:'필수값 누락'});
   let list=readCustomers(),i=list.findIndex(x=>x.nickname===c.nickname||x.phone===c.phone);
   const next={id:c.id||Date.now().toString(36),joinedAt:c.joinedAt||new Date().toLocaleString('ko-KR'),source:c.source||'가입폼',...c};
   if(i>=0)list[i]={...list[i],...next};else list.push(next);
   saveCustomers(list);json(res,200,next)
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