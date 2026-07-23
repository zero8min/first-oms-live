const http=require('http'),fs=require('fs'),path=require('path'),url=require('url'),crypto=require('crypto');
const ROOT=__dirname, DATA=path.join(ROOT,'data','customers.json');
if(!fs.existsSync(path.dirname(DATA)))fs.mkdirSync(path.dirname(DATA),{recursive:true});
if(!fs.existsSync(DATA))fs.writeFileSync(DATA,'[]','utf8');
const mime={'.html':'text/html; charset=utf-8','.js':'application/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.json':'application/json; charset=utf-8','.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.ico':'image/x-icon'};
function readCustomers(){try{return JSON.parse(fs.readFileSync(DATA,'utf8'))}catch(e){return[]}}
function saveCustomers(v){fs.writeFileSync(DATA,JSON.stringify(v,null,2),'utf8')}

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
 const apiKey=process.env.SOLAPI_API_KEY,apiSecret=process.env.SOLAPI_API_SECRET,sender=onlyDigits(process.env.SOLAPI_SENDER);
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

function json(res,code,data){res.writeHead(code,{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'});res.end(JSON.stringify(data))}
const server=http.createServer((req,res)=>{
 res.setHeader('Access-Control-Allow-Origin','*');
 res.setHeader('Access-Control-Allow-Headers','Content-Type');
 res.setHeader('Access-Control-Allow-Methods','GET,POST,DELETE,OPTIONS');
 if(req.method==='OPTIONS'){res.writeHead(204);return res.end()}
 const u=url.parse(req.url,true);
 if(u.pathname==='/api/health')return json(res,200,{ok:true,time:new Date().toISOString()});

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

 if(u.pathname==='/api/customers'&&req.method==='GET')return json(res,200,readCustomers());

 if(u.pathname.startsWith('/api/customers/')&&req.method==='DELETE'){
  const nickname=decodeURIComponent(u.pathname.split('/').pop());
  let list=readCustomers().filter(x=>x.nickname!==nickname);
  saveCustomers(list);return json(res,200,{ok:true})
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