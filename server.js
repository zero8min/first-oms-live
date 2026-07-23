const http=require('http'),fs=require('fs'),path=require('path'),url=require('url'),crypto=require('crypto');
const ROOT=__dirname, DATA=path.join(ROOT,'data','customers.json'), YT_AUTH=path.join(ROOT,'data','youtube-auth.json');
if(!fs.existsSync(path.dirname(DATA)))fs.mkdirSync(path.dirname(DATA),{recursive:true});
if(!fs.existsSync(DATA))fs.writeFileSync(DATA,'[]','utf8');
if(!fs.existsSync(YT_AUTH))fs.writeFileSync(YT_AUTH,'{}','utf8');
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
async function liveChatIdForVideo(videoId){
 const j=await youtubeApi('videos',{part:'liveStreamingDetails',id:videoId});
 const item=j.items&&j.items[0];
 const id=item?.liveStreamingDetails?.activeLiveChatId;
 if(!id)throw new Error('현재 방송의 실시간 채팅 ID를 찾지 못했습니다. 방송이 실제 LIVE 상태이고 채팅이 켜져 있는지 확인해 주세요.');
 return id
}

function json(res,code,data){res.writeHead(code,{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'});res.end(JSON.stringify(data))}
const server=http.createServer((req,res)=>{
 res.setHeader('Access-Control-Allow-Origin','*');
 res.setHeader('Access-Control-Allow-Headers','Content-Type');
 res.setHeader('Access-Control-Allow-Methods','GET,POST,DELETE,OPTIONS');
 if(req.method==='OPTIONS'){res.writeHead(204);return res.end()}
 const u=url.parse(req.url,true);
 if(u.pathname==='/api/health')return json(res,200,{ok:true,time:new Date().toISOString()});


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
      publishedAt:x.snippet?.publishedAt||''
    }));
    return json(res,200,{ok:true,liveChatId,messages,nextPageToken:j.nextPageToken||'',pollingIntervalMillis:j.pollingIntervalMillis||3500})
   }catch(e){
    const status=e.status===401||/연결|API_KEY/.test(e.message)?401:500;
    return json(res,status,{ok:false,error:e.message})
   }
  })()
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