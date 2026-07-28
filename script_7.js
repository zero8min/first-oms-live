<\/script><style>*{box-sizing:border-box}body{margin:0;font-family:Arial,'Noto Sans KR',sans-serif;color:#111}.page{width:210mm;min-height:297mm;padding:14mm;page-break-after:always;position:relative}.page:last-child{page-break-after:auto}header{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:3px solid #111;padding-bottom:12px}h1{font-size:25px;margin:0 0 8px}.qr img,.qr canvas{width:110px!important;height:110px!important}.meta{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:18px 0;font-size:14px}.address{grid-column:1/-1;line-height:1.5}table{width:100%;border-collapse:collapse;table-layout:fixed}th,td{border:1px solid #555;padding:11px 8px;text-align:center;word-break:keep-all}th:nth-child(2),td:nth-child(2){text-align:left;width:42%}.totals{margin:18px 0 0 auto;width:330px;font-size:16px;line-height:2;text-align:right}.grand{font-size:23px;border-top:2px solid #111;margin-top:4px;padding-top:5px}footer{position:absolute;bottom:12mm;left:14mm;right:14mm;text-align:center;color:#555}@page{size:A4 portrait;margin:0}</style></head><body>${pages}<script>const d=${data};window.onload=()=>{d.forEach((x,i)=>new QRCode(document.getElementById('qr-'+i),{text:x.url,width:180,height:180}));setTimeout(()=>window.print(),700)};<\/script></body></html>`);w.document.close()}
function downloadReceiptList(){
 const rows=getReceipts().map((r,i)=>({'정산번호':r.date.replaceAll('-','')+'-'+String(i+1).padStart(3,'0'),'방송일':r.date,'실명':r.customer?.name||'','닉네임':r.nick,'전화번호':r.customer?.phone||'','주소':[r.customer?.postalCode,r.customer?.address,r.customer?.detailAddress].filter(Boolean).join(' '),'상품내역':r.items.map(x=>x.item+' '+x.qty+'개').join(' / '),'상품합계':r.subtotal,'배송비':r.fee,'결제금액':r.total,'고객매칭':r.customer?'완료':'확인필요','입금상태':r.payment.status==='paid'?'입금완료':'미입금/확인필요'}));
 exportXlsx(rows,'FIRST_OMS_정산서목록.xlsx','정산서');
}
function downloadCustomerTemplate(){exportXlsx([{'실명':'홍길동','닉네임':'길동맘','전화번호':'010-0000-0000','주소':'서울시 ...'}],'FIRST_OMS_고객DB_양식.xlsx','고객DB')}
function exportXlsx(rows,name,sheet){
 if(!rows.length){alert('출력할 자료가 없습니다.');return}
 const ws=XLSX.utils.json_to_sheet(rows),wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,sheet);XLSX.writeFile(wb,name);
}

function closeAdminPostcode(){const layer=$('adminPostcodeLayer');layer.classList.remove('show');$('adminPostcodeEmbed').innerHTML=''}
function findAdminAddress(){
 if(!window.daum||!window.daum.Postcode)return alert('주소검색을 불러오지 못했습니다. 기본주소를 직접 입력해 주세요.');
 const layer=$('adminPostcodeLayer'),embed=$('adminPostcodeEmbed');layer.classList.add('show');embed.innerHTML='';
 try{new window.daum.Postcode({oncomplete:function(d){$('cPostalCode').value=d.zonecode||'';$('cAddress').value=d.roadAddress||d.jibunAddress||'';$('cAddressStatus').textContent='주소 선택 완료 · 상세주소를 입력해 주세요.';closeAdminPostcode();setTimeout(()=>$('cDetailAddress').focus(),100)},onresize:function(size){embed.style.height=Math.min(size.height||600,620)+'px'},width:'100%',height:'100%'}).embed(embed)}catch(e){closeAdminPostcode();alert('주소검색 실행 중 오류가 발생했습니다. 기본주소를 직접 입력해 주세요.')}
}
function openCustomerModal(){ $('customerModal').dataset.receiptKey='';$('cEditIndex').value='';$('cName').value='';$('cNick').value='';$('cPhone').value='';$('cPostalCode').value='';$('cAddress').value='';$('cDetailAddress').value='';$('cMemo').value='';$('cAddressStatus').textContent='';$('customerModal').classList.add('show');setTimeout(()=>$('cName').focus(),50)}
function closeCustomerModal(){$('customerModal').classList.remove('show')}

function openCustomerForReceipt(i){
 const r=getReceipts()[i];
 if(!r)return;
 openCustomerModal();
 $('cNick').value=r.nick||'';
 $('cName').value='';
 $('cPhone').value='';
 $('cPostalCode').value='';
 $('cAddress').value='';
 $('cDetailAddress').value='';
 $('cMemo').value='';
 $('cAddressStatus').textContent='정산서의 닉네임을 자동 입력했습니다. 고객 정보를 채우고 저장하세요.';
 $('customerModal').dataset.receiptKey=r.key||'';
}

function editCustomerById(id){const i=state.customers.findIndex(c=>String(c.id)===String(id));if(i<0)return alert('고객정보를 찾을 수 없습니다.');editCustomer(i)}
function editCustomer(i){const c=state.customers[i];$('cEditIndex').value=i;$('cName').value=c.name||'';$('cNick').value=c.nick||c.nickname||'';$('cPhone').value=c.phone||'';$('cPostalCode').value=c.postalCode||'';$('cAddress').value=c.address||'';$('cDetailAddress').value=c.detailAddress||'';$('cMemo').value=c.memo||'';$('cAddressStatus').textContent='';$('customerModal').classList.add('show')}
async function saveCustomer(){
 const oldIndex=$('cEditIndex').value;
 const old=oldIndex===''?null:state.customers[Number(oldIndex)];
 const c={id:old?.id||crypto.randomUUID(),name:$('cName').value.trim(),nickname:$('cNick').value.trim(),nick:$('cNick').value.trim(),phone:$('cPhone').value.trim(),postalCode:$('cPostalCode').value.trim(),address:$('cAddress').value.trim(),detailAddress:$('cDetailAddress').value.trim(),memo:$('cMemo').value.trim(),joinedAt:old?.joinedAt||new Date().toLocaleString('ko-KR'),source:'관리자폼'};
 if(!c.name||!c.nickname||!c.phone||!c.address){alert('성함, 닉네임, 연락처, 기본주소를 모두 입력해 주세요.');return}
 try{
  const r=await fetch('/api/customers',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(c)});
  const data=await r.json(); if(!r.ok)throw new Error(data.error||'고객 저장 실패');
  await syncServerCustomers(false); autoMatchAll(); saveSilently(); renderAll(); closeCustomerModal(); alert('고객정보가 저장되어 정산서와 택배실에 바로 반영되었습니다.');
 }catch(e){alert('고객정보 저장 실패: '+e.message)}
}
async function deleteCustomer(i){const c=state.customers[i];if(!confirm('고객정보는 완전 삭제되지 않습니다. 목록에서 보관처리할까요?'))return;try{const key=encodeURIComponent(c.nickname||c.nick||'');const r=await fetch('/api/customers/'+key,{method:'DELETE'});if(!r.ok)throw new Error('보관 실패');await syncServerCustomers(false);alert('고객정보를 삭제하지 않고 보관 처리했습니다.')}catch(e){alert('고객 보관 실패: '+e.message)}}
function deleteOrder(i){state.orders.splice(i,1);saveSilently();renderAll()}
function clearOrders(){if(confirm('모든 주문을 초기화할까요?')){state.orders=[];saveSilently();renderAll()}}
function clearPayments(){if(confirm('입금내역을 초기화할까요?')){state.payments=[];saveSilently();renderAll()}}
let stateSaveTimer=null;
function saveSilently(){
 localStorage.setItem(KEY,JSON.stringify(state));
 clearTimeout(stateSaveTimer);
 stateSaveTimer=setTimeout(async()=>{try{
  const r=await fetch('/api/state',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(state)});
  if(!r.ok)throw new Error('서버 저장 실패 '+r.status);
  const d=await r.json();const el=document.getElementById('serverSaveStatus');if(el)el.textContent=`자동보존 완료 · 주문 ${d.orders}건 · 고객 ${d.customers}명 · ${new Date().toLocaleTimeString('ko-KR')}`;
 }catch(e){console.error(e);const el=document.getElementById('serverSaveStatus');if(el)el.textContent='⚠ 서버 자동보존 재시도 중 — 브라우저에는 저장됨';}},350);
}
function updateKpi(){const rs=getReceipts();$('kpiOrderRows').textContent=state.orders.length;$('kpiReceipts').textContent=rs.length;$('kpiMatched').textContent=rs.filter(r=>r.customer).length;$('kpiPaid').textContent=rs.filter(r=>r.payment.status==='paid').length;refreshSolapiKpi()}
async function refreshSolapiKpi(){const el=$('kpiSolapi');if(!el)return;try{const r=await fetch('/api/solapi/config',{cache:'no-store'}),d=await r.json();el.textContent=d.configured?'문자 준비':'미연동'}catch(e){el.textContent='확인필요'}}
function renderAll(){syncShippingFeeLockUI();renderMessageSettings();renderOrders();renderCustomers();renderReceipts();renderPayments();renderShipping();updateKpi()}
function esc(v){return String(v??'').replace(/[&<>"']/g,s=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]))}

async function downloadCustomerBackup(){
 try{
  const r=await fetch('/api/customers/backup',{cache:'no-store'});
  if(!r.ok)throw new Error('HTTP '+r.status);
  const blob=await r.blob();
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download='FIRST_OMS_고객정보백업_'+today()+'.json';
  document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(a.href);
  alert('고객정보 백업 파일을 저장했습니다. 서버에도 customers-backup.json으로 자동 보관됩니다.');
 }catch(e){
  const blob=new Blob([JSON.stringify({version:1,exportedAt:new Date().toISOString(),customers:state.customers},null,2)],{type:'application/json'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='FIRST_OMS_고객정보백업_'+today()+'.json';a.click();URL.revokeObjectURL(a.href);
  alert('서버 연결 없이 현재 화면의 고객정보를 백업했습니다.');
 }
}

async function downloadBackup(){
 try{
  const r=await fetch('/api/state/backup',{cache:'no-store'});if(!r.ok)throw new Error('HTTP '+r.status);
  const blob=await r.blob(),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='땡라이브서버_전체백업_'+today()+'.json';a.click();URL.revokeObjectURL(a.href);
 }catch(e){const blob=new Blob([JSON.stringify({version:6,exportedAt:new Date().toISOString(),state},null,2)],{type:'application/json'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='땡라이브서버_전체백업_'+today()+'.json';a.click();URL.revokeObjectURL(a.href)}
}
$('backupFile').onchange=e=>{
 const f=e.target.files[0]; if(!f)return;
 const r=new FileReader();
 r.onload=ev=>{try{
   const obj=JSON.parse(ev.target.result), restored=obj.state||obj;
   if(!restored||!Array.isArray(restored.customers)) throw new Error('올바른 백업 파일이 아닙니다.');
   state={orders:Array.isArray(restored.orders)?restored.orders:[],customers:restored.customers,payments:Array.isArray(restored.payments)?restored.payments:[],settings:{...state.settings,...(restored.settings||{})}};
   autoMatchAll(); saveSilently(); renderAll(); alert('백업을 복원했습니다.');
 }catch(err){alert('백업 복원 실패: '+err.message)}};
 r.readAsText(f);
};
function recoverOldData(){
 let found=[];
 for(let i=0;i<localStorage.length;i++){
   const key=localStorage.key(i); if(key===KEY)continue;
   try{
     const obj=JSON.parse(localStorage.getItem(key)), cand=obj?.state||obj;
     if(cand && Array.isArray(cand.customers) && cand.customers.length){
       found.push({key,customers:cand.customers,orders:Array.isArray(cand.orders)?cand.orders:[],payments:Array.isArray(cand.payments)?cand.payments:[]});
     }
   }catch(e){}
 }
 if(!found.length){alert('이 브라우저에서 복구 가능한 이전 고객DB를 찾지 못했습니다. 기존 주소와 같은 브라우저에서 다시 시도해 주세요.');return}
 found.sort((a,b)=>b.customers.length-a.customers.length);
 const pick=found[0];
 if(confirm(`이전 저장자료 "${pick.key}"에서 고객 ${pick.customers.length}명을 발견했습니다. 현재 자료에 합칠까요?`)){
   state.customers.push(...pick.customers); dedupeCustomers();
   if(!state.orders.length) state.orders=pick.orders;
   if(!state.payments.length) state.payments=pick.payments;
   autoMatchAll(); saveSilently(); renderAll(); alert('이전 고객DB를 복구했습니다. 바로 전체 백업을 내려받아 주세요.');
 }
}

async function syncServerCustomers(showMessage=false){
 try{
  const r=await fetch('/api/customers?ts='+Date.now(),{cache:'no-store'}); if(!r.ok)throw new Error('HTTP '+r.status);
  const list=await r.json();
  if(Array.isArray(list)){
   const serverCustomers=list.map(c=>({
    ...c,
    id:c.id||crypto.randomUUID(),
    name:cleanBrokenText(c.name||''),
    nickname:cleanBrokenText(c.nickname||c.nick||''),
    nick:cleanBrokenText(c.nickname||c.nick||''),
    phone:cleanBrokenText(c.phone||''),
    postalCode:cleanBrokenText(c.postalCode||''),
    address:cleanBrokenText(c.address||''),
    detailAddress:cleanBrokenText(c.detailAddress||''),
    memo:cleanBrokenText(c.memo||'')
   })).filter(c=>c.name||c.nick);
   state.customers=serverCustomers; dedupeCustomers(); autoMatchAll(); saveSilently(); renderAll();
   if(showMessage)alert(`고객폼에서 등록된 고객 ${state.customers.length}명을 불러왔습니다.`);
  }
 }catch(e){console.warn('고객DB 서버 동기화 실패',e);if(showMessage)alert('고객폼 정보 불러오기 실패: '+e.message)}
}
function applyServerCustomers(list){
 if(!Array.isArray(list))return;
 const serverCustomers=list.map(c=>({
  ...c,id:c.id||crypto.randomUUID(),name:cleanBrokenText(c.name||''),
  nickname:cleanBrokenText(c.nickname||c.nick||''),nick:cleanBrokenText(c.nickname||c.nick||''),
  phone:cleanBrokenText(c.phone||''),postalCode:cleanBrokenText(c.postalCode||''),
  address:cleanBrokenText(c.address||''),detailAddress:cleanBrokenText(c.detailAddress||''),memo:cleanBrokenText(c.memo||'')
 })).filter(c=>c.name||c.nick);
 state.customers=serverCustomers;dedupeCustomers();autoMatchAll();saveSilently();renderAll();
}
function startCustomerRealtime(){
 if(!window.EventSource)return false;
 const es=new EventSource('/api/customers/stream');
 es.addEventListener('customers',ev=>{try{applyServerCustomers(JSON.parse(ev.data))}catch(e){console.warn('고객 실시간 반영 실패',e)}});
 es.onerror=()=>console.warn('고객 실시간 연결 재시도 중');
 window.customerEventSource=es;
 return true;
}
async function bootFirstOms(){
 try{
  const r=await fetch('/api/state?ts='+Date.now(),{cache:'no-store'});
  if(r.ok){const d=await r.json(),remote=d.state;if(remote&&Array.isArray(remote.orders)&&Array.isArray(remote.customers)){
   state={...state,...remote,settings:{...state.settings,...(remote.settings||{})}};
   localStorage.setItem(KEY,JSON.stringify(state));renderMessageSettings();renderAll();
  }else throw new Error('서버 상태 없음');
  }else throw new Error('서버 응답 오류');
 }catch(e){
  const saved=localStorage.getItem(KEY);
  if(saved)loadAll();else{try{const r=await fetch('/data/initial-backup.json?ts='+Date.now(),{cache:'no-store'});const obj=await r.json(),restored=obj.state||obj;state={...state,...restored,settings:{...state.settings,...(restored.settings||{})}};saveSilently();renderAll()}catch(x){loadAll()}}
 }
 await syncServerCustomers(false);startCustomerRealtime();refreshSolapiKpi();
}
bootFirstOms();
setInterval(()=>{if(!document.hidden)syncServerCustomers(false)},10000);
setInterval(()=>{if(!document.hidden&&document.querySelector('[data-page=shipping]')?.classList.contains('active'))syncShippingScans(true)},3000);
syncShippingScans(true);

async function loadLoginUser(){try{const r=await fetch('/api/auth/me');if(!r.ok){location.href='/login.html';return}const d=await r.json();const u=d.user||{};if(u.mustChangePassword){location.href='/change-password.html';return}const el=document.getElementById('loginUser');if(el)el.textContent=`${u.company||u.username} · ${u.code||''}`;const b=document.getElementById('accountAdminBtn');if(b&&u.role==='superadmin')b.style.display='inline-block'}catch(e){location.href='/login.html'}}
async function logoutNow(){try{await fetch('/api/auth/logout',{method:'POST'})}finally{location.href='/login.html'}}
loadLoginUser();

async function runFunctionCheck(){
 const tests=[['로그인 상태','/api/auth/me'],['서버 상태','/api/health'],['판매·정산 데이터','/api/state'],['고객DB','/api/customers'],['고객DB 백업','/api/customers/backup/status'],['솔라피 설정','/api/solapi/config']];
 const result=[];for(const [name,path] of tests){try{const r=await fetch(path,{cache:'no-store'});result.push(`${r.ok?'✅':'⚠️'} ${name} (${r.status})`)}catch(e){result.push(`❌ ${name}: ${e.message}`)}}
 const required=['go','renderOrders','renderCustomers','renderReceipts','renderPayments','renderShipping','saveAll','downloadBackup','openCustomerModal','sendAllMms','logoutNow'];
 const missing=required.filter(n=>typeof window[n]!=='function');result.push(missing.length?`❌ 버튼함수 누락: ${missing.join(', ')}`:'✅ 주요 버튼함수 연결 정상');alert('땡라이브 기능점검 결과\n\n'+result.join('\n'))
}
