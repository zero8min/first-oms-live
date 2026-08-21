
const qs=new URLSearchParams(location.search);
const tenant=qs.get('tenant')||'FIRST-0001';
const fixedCode=(qs.get('code')||'').trim().toUpperCase();
let worker=localStorage.getItem('FIRST_OMS_PACKER_NAME')||'';
let job=null,stream=null,scanTimer=null,zxingReader=null,locked=false;
const esc=v=>String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
const totalQty=items=>(items||[]).reduce((a,x)=>a+(Number(x.qty)||0),0);
function askWorker(force=false){
 if(worker&&!force)return true;
 const v=prompt('포장자 이름을 입력해주세요',worker||'');
 if(!v||!v.trim())return false;
 worker=v.trim();localStorage.setItem('FIRST_OMS_PACKER_NAME',worker);return true;
}
async function loadJob(){
 if(!askWorker()){document.getElementById('app').innerHTML='<div class="empty">포장자 이름을 입력해야 작업을 시작할 수 있습니다.<button class="btn primary" onclick="loadJob()">이름 입력</button></div>';return}
 stopScan();locked=false;
 try{
  let url=fixedCode?'/api/public/packing?code='+encodeURIComponent(fixedCode):'/api/public/packing/next?tenant='+encodeURIComponent(tenant)+'&worker='+encodeURIComponent(worker);
  const r=await fetch(url,{cache:'no-store'}),d=await r.json();if(!r.ok||!d.ok)throw new Error(d.error||'조회 실패');
  job=d.job||null;if(job&&!job.tenantCode)job.tenantCode=tenant;
  renderJob();
 }catch(e){document.getElementById('app').innerHTML='<div class="notice warn">'+esc(e.message)+'</div><button class="btn muted" onclick="loadJob()">다시 불러오기</button>'}
}
function renderJob(){
 const app=document.getElementById('app');
 if(!job){app.innerHTML=`<div class="top"><div class="worker">${esc(worker)}님</div><button class="change" onclick="changeWorker()">포장자 변경</button></div><div class="empty"><div style="font-size:30px;font-weight:950">🎉 포장할 주문이 없습니다</div><p>현재 출고준비가 필요한 주문을 모두 처리했습니다.</p><button class="btn muted" onclick="loadJob()">새로고침</button></div>`;return}
 if(job.status==='packed'||job.completedAt){return showScanner()}
 const normal=(job.items||[]).filter(x=>!x.service),services=(job.items||[]).filter(x=>x.service);
 app.innerHTML=`<div class="top"><div class="worker">${esc(worker)}님 <span class="pill">자동 저장됨</span></div><button class="change" onclick="changeWorker()">포장자 변경</button></div>
 <div class="job">${esc(job.jobNo||'')}번 / ${esc(job.name||job.nick||'고객')}</div>
 <div class="count">📦 총 ${totalQty(job.items)}개 넣으세요</div>
 <div class="items">${normal.map(x=>`<div class="item"><div class="name">#${esc(x.number)} ${esc(x.item)}</div><div class="qty">×${Number(x.qty)||0}</div></div>`).join('')}</div>
 ${services.length?`<div class="service">🩵 🎁 서비스<br>${services.map(x=>`#${esc(x.number)} ${esc(x.item)} ×${Number(x.qty)||0}`).join('<br>')}</div>`:''}
 <button class="btn good" onclick="packed()">다 넣었어요</button>
 <div class="small">이 버튼을 누르면 관리자 택배리스트에 즉시 포장완료로 기록됩니다. 이후 송장스캔이 끝날 때까지 다음 주문으로 넘어갈 수 없습니다.</div>`;
}
async function packed(){
 if(locked)return;locked=true;
 try{
  const r=await fetch('/api/public/packing/packed',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({tenantCode:job.tenantCode||tenant,code:job.code,worker})}),d=await r.json();
  if(!r.ok||!d.ok)throw new Error(d.error||'포장완료 처리 실패');
  job.status='packed';job.completedAt=d.packingCompletedAt;showScanner();
 }catch(e){alert(e.message);locked=false}
}
function showScanner(){
 const app=document.getElementById('app');
 app.innerHTML=`<div class="top"><div class="worker">${esc(worker)}님</div><button class="change" onclick="changeWorker()" disabled>포장 진행중</button></div>
 <div class="scan-screen show">
   <div class="scan-title">📷 송장을 스캔해주세요</div>
   <div class="scan-sub"><b>${esc(job.jobNo||'')}번 / ${esc(job.name||job.nick||'')}</b><br>CJ대한통운 송장 바코드를 카메라에 맞춰주세요.</div>
   <div class="camera"><video id="video" playsinline muted></video></div>
   <div id="scanStatus" class="notice ok">카메라를 준비하는 중입니다.</div>
   <button class="btn muted" onclick="restartScan()">카메라 다시 열기</button>
 </div>`;
 locked=false;startScan();
}
async function startScan(){
 stopScan();
 const status=document.getElementById('scanStatus'),video=document.getElementById('video');if(!video)return;
 try{
  stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:{ideal:'environment'}},audio:false});video.srcObject=stream;await video.play();
  if('BarcodeDetector' in window){
   const det=new BarcodeDetector({formats:['code_128','code_39','ean_13','ean_8','itf','codabar']});
   const loop=async()=>{if(!stream)return;try{const arr=await det.detect(video);if(arr?.length&&arr[0].rawValue){await finalScan(arr[0].rawValue);return}}catch(e){}scanTimer=setTimeout(loop,250)};loop();status.textContent='송장 바코드를 화면 중앙에 맞춰주세요.';return
  }
  if(window.ZXing&&window.ZXing.BrowserMultiFormatReader){
   zxingReader=new ZXing.BrowserMultiFormatReader();status.textContent='송장 바코드를 화면 중앙에 맞춰주세요.';
   zxingReader.decodeFromVideoElementContinuously(video,(result,err)=>{if(result&&!locked)finalScan(result.getText())});return
  }
  throw new Error('이 브라우저에서 바코드 인식을 시작할 수 없습니다. Chrome 또는 Safari 최신 버전으로 열어주세요.');
 }catch(e){if(status)status.innerHTML='<b>카메라 오류</b><br>'+esc(e.message)}
}
async function finalScan(raw){
 if(locked)return;locked=true;const tracking=String(raw||'').replace(/\s+/g,'').trim();const status=document.getElementById('scanStatus');
 if(!tracking){locked=false;return}
 if(status)status.textContent='송장을 확인하는 중입니다...';
 stopScan();
 try{
  const r=await fetch('/api/public/packing/scan-final',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({tenantCode:job.tenantCode||tenant,code:job.code,worker,trackingNumber:tracking})}),d=await r.json();
  if(!r.ok||!d.ok){
   if(d.reason==='WRONG_LABEL'){document.getElementById('app').innerHTML=`<div class="notice warn" style="font-size:24px;text-align:center"><div style="font-size:42px">🚨</div><b>잘못된 송장입니다</b><br><br>이 송장을 붙이면 안 됩니다.<br>현재 고객의 송장을 다시 스캔해주세요.</div><button class="btn primary" onclick="showScanner()">현재 고객 송장 다시 스캔</button>`;locked=false;return}
   throw new Error(d.error||'송장 확인 실패')
  }
  const msg=d.smsFailed?`<div class="notice smsfail">⚠️ 배송문자 발송실패<br>포장과 송장 확인은 정상 완료되었습니다. 관리자가 재발송할 수 있습니다.</div>`:`<div class="notice ok">문자까지 발송되었습니다.</div>`;
  document.getElementById('app').innerHTML=`<div style="text-align:center;padding:28px 6px"><div class="done">✅ 완료!</div><p style="font-size:20px"><b>${esc(job.name||job.nick)}</b><br>송장번호 ${esc(d.trackingNumber)}<br>CJ대한통운</p>${msg}<div class="small">잠시 후 다음 포장 주문이 자동으로 표시됩니다.</div></div>`;
  setTimeout(()=>{job=null;loadJob()},1400);
 }catch(e){document.getElementById('app').innerHTML=`<div class="notice warn"><b>송장 처리 실패</b><br>${esc(e.message)}</div><button class="btn primary" onclick="showScanner()">다시 스캔</button>`;locked=false}
}
function stopScan(){
 if(scanTimer){clearTimeout(scanTimer);scanTimer=null}
 if(zxingReader){try{zxingReader.reset()}catch(e){}zxingReader=null}
 if(stream){stream.getTracks().forEach(t=>t.stop());stream=null}
}
function restartScan(){locked=false;startScan()}
function changeWorker(){stopScan();if(askWorker(true))loadJob()}
window.addEventListener('pagehide',stopScan);
loadJob();
