
const KEY='firstOmsEmergencySettlementV1';
let state={orders:[],customers:[],payments:[],paymentOverrides:{},shippingScans:{},settings:{bank:'카카오뱅크',holder:'김미숙',account:'',contact:'문의는 땡라이브 카카오채널로 부탁드립니다.',smsTemplate:'[땡라이브 정산서]\n{고객명}님\n{상품내역}\n상품합계 {상품합계}\n배송비 {배송비}\n결제금액 {결제금액}\n\n입금계좌: {은행} {계좌번호} {예금주}\n{문의안내}\n감사합니다.'}};
const $=id=>document.getElementById(id);
const money=n=>(Number(n)||0).toLocaleString('ko-KR')+'원';
const cleanBrokenText=v=>String(v??'').replace(/\uFFFD/g,' ').replace(/[�]+/g,' ').replace(/\?{2,}/g,' ').replace(/\s+/g,' ').trim();
const hasBrokenText=v=>/[�]|\uFFFD|\?{2,}/.test(String(v??''));
const safeCustomerText=(v,fallback='미등록')=>{const x=cleanBrokenText(v);return x||fallback};
const norm=v=>cleanBrokenText(v).replace(/[^0-9a-zA-Z가-힣]/g,'').toLowerCase();
const phoneNorm=v=>String(v??'').replace(/\D/g,'');
const today=()=>new Date().toISOString().slice(0,10);
$('broadcastDate').value=today();

document.querySelectorAll('.tab').forEach(b=>b.addEventListener('click',e=>{e.preventDefault();go(b.dataset.page)}));
function go(id){
 const target=document.getElementById(id);
 if(!target){console.warn('페이지를 찾을 수 없습니다:',id);return;}
 document.querySelectorAll('.tab').forEach(b=>b.classList.toggle('active',b.dataset.page===id));
 document.querySelectorAll('.page').forEach(p=>p.classList.toggle('active',p.id===id));
 if(id==='dashboard') updateKpi();
 if(id==='orders') renderOrders();
 if(id==='customers') renderCustomers();
 if(id==='receipts') renderReceipts();
 if(id==='payments') renderPayments();
 if(id==='shipping') renderShipping();
 window.scrollTo({top:0,behavior:'auto'});
}
function saveAll(){localStorage.setItem(KEY,JSON.stringify(state));alert('현재 자료를 이 기기에 저장했습니다.');}
function loadAll(){try{const saved=JSON.parse(localStorage.getItem(KEY));if(saved)state={...state,...saved,settings:{...state.settings,...(saved.settings||{})}}}catch(e){} renderMessageSettings(true);initSmsTemplateGuard();renderAll();}
function readWorkbook(file, cb){
 const r=new FileReader(); r.onload=e=>{try{
  const wb=XLSX.read(e.target.result,{type:'array',cellDates:true});
  // 땡라이브 양식은 F열 수식이 100만 행까지 미리 들어가 있어 !ref가 크게 잡힙니다.
  // 전체 시트를 무작정 읽지 않고, 실제 데이터가 있는 시트/행만 찾아서 읽습니다.
  const aliases = cb===importOrders ? orderAliases : (cb===importCustomers ? customerAliases : paymentAliases);
  const preferred = cb===importOrders ? ['판매리스트','판매 리스트','주문리스트','주문'] : [];
  const names=[...wb.SheetNames].sort((a,b)=>{
    const ap=preferred.some(x=>norm(a).includes(norm(x)))?0:1;
    const bp=preferred.some(x=>norm(b).includes(norm(x)))?0:1;
    return ap-bp;
  });
  let bestRows=null,bestScore=-1;
  for(const sn of names){
    const ws=wb.Sheets[sn]; if(!ws)continue;
    // 헤더 탐색은 위쪽 60행, A~Z만 읽기
    const probe=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',blankrows:true,raw:false,range:'A1:Z60'});
    let hi=-1,score=-1;
    for(let i=0;i<probe.length;i++){
      const row=(probe[i]||[]).map(norm); let sc=0;
      Object.values(aliases).forEach(list=>{if(row.some(v=>list.some(a=>v.includes(norm(a)))))sc++});
      if(sc>score){score=sc;hi=i}
    }
    if(score<2)continue;
    // 실제 주문행은 최대 50,000행까지만 읽고, 연속 빈 행 200개면 종료
    const endRow=Math.min(50000, Math.max(hi+5000, 5000));
    const rawRows=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',blankrows:true,raw:false,range:`A${hi+1}:Z${endRow}`});
    const rows=[]; let emptyRun=0;
    for(const row of rawRows){
      const meaningful=(row||[]).slice(0,12).some(v=>String(v??'').trim()!=='' && !String(v).trim().startsWith('='));
      if(meaningful){rows.push(row);emptyRun=0}else{emptyRun++; if(emptyRun>=200 && rows.length>1)break}
    }
    if(score>bestScore){bestScore=score;bestRows=rows}
    if(preferred.some(x=>norm(sn).includes(norm(x))))break;
  }
  if(!bestRows)throw new Error('번호·고객명·품목·수량·단가·금액 제목행을 찾지 못했습니다.');
  cb(bestRows,file.name);
 }catch(err){alert('파일을 읽지 못했습니다: '+err.message)}}; r.readAsArrayBuffer(file);
}
function findHeader(rows, aliases){
 for(let i=0;i<Math.min(rows.length,30);i++){
  const row=rows[i].map(norm); let score=0;
  Object.values(aliases).forEach(list=>{if(row.some(v=>list.some(a=>v.includes(norm(a)))))score++});
  if(score>=2)return i;
 } return 0;
}
function colMap(header, aliases){
 const h=header.map(norm), m={};
 for(const [key,list] of Object.entries(aliases)){
  m[key]=h.findIndex(v=>list.some(a=>v.includes(norm(a))));
 }
 return m;
}
const orderAliases={
 number:['번호','상품번호','순번','#'],
 nick:['닉네임','고객닉네임','구매자','고객명','성명'],
 item:['품목','상품명','상품','제품명'],
 qty:['수량','개수','주문수량'],
 unit:['단가','가격','판매가'],
 amount:['금액','합계','총액','판매금액'],
 date:['방송일','방송일자','주문일','날짜']
};
function importOrders(rows,filename){
 const hi=findHeader(rows,orderAliases), map=colMap(rows[hi]||[],orderAliases);
 const bdate=$('broadcastDate').value||today(), fee=Number($('shippingFee').value)||0;
 let parsed=[];
 rows.slice(hi+1).forEach((r,idx)=>{
  const nick=map.nick>=0?r[map.nick]:'', item=map.item>=0?r[map.item]:'', number=map.number>=0?r[map.number]:(idx+1);
  if(!String(nick).trim()||!String(item).trim())return;
  const qty=Number(String(map.qty>=0?r[map.qty]:1).replace(/[^0-9.-]/g,''))||1;
  const unit=Number(String(map.unit>=0?r[map.unit]:0).replace(/[^0-9.-]/g,''))||0;
  const amount=Number(String(map.amount>=0?r[map.amount]:0).replace(/[^0-9.-]/g,''))||qty*unit;
  const dateVal=map.date>=0&&r[map.date]?formatDate(r[map.date]):bdate;
  parsed.push({id:crypto.randomUUID(),number:String(number||idx+1).replace(/^#/,'').trim(),date:dateVal,nick:String(nick).trim(),item:String(item).trim(),qty,unit,amount,fee,source:filename});
 });
 if($('duplicateMode').value==='replaceDate'){
  const dates=[...new Set(parsed.map(x=>x.date))]; state.orders=state.orders.filter(o=>!dates.includes(o.date));
 }
 state.orders.push(...parsed); autoMatchAll(); saveSilently(); renderAll(); go('orders');
 alert(parsed.length+'개 주문행을 불러왔습니다.');
}
function formatDate(v){
 if(v instanceof Date&&!isNaN(v))return v.toISOString().slice(0,10);
 const s=String(v).trim(); const d=new Date(s); return isNaN(d)?($('broadcastDate').value||today()):d.toISOString().slice(0,10);
}
$('orderFile').onchange=e=>[...e.target.files].forEach(f=>readWorkbook(f,importOrders));
$('quickOrderFile').onchange=e=>[...e.target.files].forEach(f=>readWorkbook(f,importOrders));

const customerAliases={name:['실명','이름','고객명','성명'],nick:['닉네임','방송닉네임','별명'],phone:['전화번호','연락처','휴대폰','핸드폰'],address:['주소','배송지','배송주소']};
function importCustomers(rows){
 const hi=findHeader(rows,customerAliases), map=colMap(rows[hi]||[],customerAliases); let n=0;
 rows.slice(hi+1).forEach(r=>{
  const name=map.name>=0?r[map.name]:'', nick=map.nick>=0?r[map.nick]:'';
  if(!String(name).trim()&&!String(nick).trim())return;
  state.customers.push({id:crypto.randomUUID(),name:String(name).trim(),nick:String(nick).trim(),phone:String(map.phone>=0?r[map.phone]:'').trim(),address:String(map.address>=0?r[map.address]:'').trim()}); n++;
 });
 dedupeCustomers(); autoMatchAll(); saveSilently(); renderAll(); alert(n+'명 고객을 불러왔습니다.');
}
$('customerFile').onchange=e=>readWorkbook(e.target.files[0],importCustomers);
function dedupeCustomers(){
 const seen=new Map(); state.customers=state.customers.filter(c=>{const k=norm(c.nick)+'|'+norm(c.name)+'|'+phoneNorm(c.phone);if(seen.has(k))return false;seen.set(k,1);return true});
}
function findCustomerByNick(nick){
 const n=norm(nick);if(!n)return {customer:null,status:'unmatched'};
 const active=state.customers.filter(c=>c.active!==false);
 const exact=active.filter(c=>[c.nick,c.nickname,c.name].some(v=>norm(v)===n));
 if(exact.length===1)return {customer:exact[0],status:'matched'};
 if(exact.length>1)return {customer:null,status:'duplicate'};
 // 깨진 문자가 들어간 기존 자료도 정상 닉네임의 한글·숫자 부분으로 다시 연결합니다.
 const fuzzy=active.filter(c=>{const vals=[c.nick,c.nickname,c.name].map(norm).filter(Boolean);return vals.some(v=>(n.length>=2&&v.includes(n))||(v.length>=2&&n.includes(v)))});
 if(fuzzy.length===1)return {customer:fuzzy[0],status:'matched-fuzzy'};
 return {customer:null,status:'unmatched'};
}
function autoMatchAll(){state.orders.forEach(o=>{const m=findCustomerByNick(o.nick);o.customerId=m.customer?.id||null;o.matchStatus=m.status})}
function getReceipts(){
 const groups=new Map();
 state.orders.forEach(o=>{
  const k=o.date+'|'+norm(o.nick);
  if(!groups.has(k))groups.set(k,{key:k,date:o.date,nick:o.nick,items:[],customerId:o.customerId,matchStatus:o.matchStatus,fee:o.fee||0});
  groups.get(k).items.push(o);
 });
 return [...groups.values()].map(g=>{
  const c=state.customers.find(x=>x.id===g.customerId)||null;
  const subtotal=g.items.reduce((a,x)=>a+(Number(x.amount)||0),0), total=subtotal+(Number(g.fee)||0);
  const pay=matchPayment(g,c,total);
  return {...g,customer:c,subtotal,total,payment:pay,starred:g.items.some(x=>x.starred)};
 }).sort((a,b)=>b.date.localeCompare(a.date)||a.nick.localeCompare(b.nick,'ko'));
}
function matchPayment(g,c,total){
 const override=state.paymentOverrides?.[g.key];
 if(override&&override.status)return {status:override.status,payment:override.payment||null,reason:'관리자 수기변경',manual:true};
 const names=[g.nick,c?.nick,c?.name].filter(Boolean).map(norm);
 const candidates=state.payments.filter(p=>names.includes(norm(p.payer)));
 const exact=candidates.filter(p=>Number(p.amount)===Number(total));
 if(exact.length===1)return {status:'paid',payment:exact[0]};
 if(exact.length>1)return {status:'review',payment:exact[0],reason:'동일 금액 입금 여러 건'};
 if(candidates.length)return {status:'amount-mismatch',payment:candidates[0],reason:'이름 일치·금액 불일치'};
 return {status:'unpaid'};
}
function renderOrders(){
 if(!state.orders.length){$('ordersTable').innerHTML='<div class="empty">판매리스트 엑셀을 올려주세요.</div>';return}
 $('ordersTable').innerHTML='<div class="scroll"><table class="orders-fixed-table"><colgroup><col><col><col><col><col><col><col><col><col></colgroup><thead><tr><th>방송일</th><th>닉네임</th><th>번호</th><th>상품</th><th>수량</th><th>단가</th><th>금액</th><th>배송비</th><th>고객매칭</th><th>관리</th></tr></thead><tbody>'+
 state.orders.map((o,i)=>`<tr><td><input type="date" value="${esc(o.date)}" onchange="updateOrder(${i},'date',this.value)"></td><td><input value="${esc(o.nick)}" onchange="updateOrder(${i},'nick',this.value)"></td><td><input value="${esc(o.number||i+1)}" onchange="updateOrder(${i},'number',this.value)"></td><td><input value="${esc(o.item)}" onchange="updateOrder(${i},'item',this.value)"></td><td><input type="number" min="0" value="${o.qty}" onchange="updateOrder(${i},'qty',this.value)"></td><td><input type="number" min="0" value="${o.unit}" onchange="updateOrder(${i},'unit',this.value)"></td><td><input type="number" min="0" value="${o.amount}" onchange="updateOrder(${i},'amount',this.value,true)"></td><td><input type="number" min="0" value="${o.fee||0}" onchange="updateOrder(${i},'fee',this.value)"></td><td>${matchBadge(o)}</td><td><button class="btn bad" onclick="deleteOrder(${i})">삭제</button></td></tr>`).join('')+
 '</tbody></table></div>';
}
function updateOrder(i,key,value,manualAmount=false){
 const o=state.orders[i]; if(!o)return;
 if(['qty','unit','amount','fee'].includes(key)) value=Number(value)||0;
 o[key]=value;
 if((key==='qty'||key==='unit')&&!manualAmount)o.amount=(Number(o.qty)||0)*(Number(o.unit)||0);
 autoMatchAll();saveSilently();renderAll();
}

function confirmShippingFee(){
 const fee=Number($('shippingFee').value)||0;if(fee<0)return alert('배송비를 확인해 주세요.');
 if(!confirm(`배송비 ${money(fee)}를 현재 판매리스트 전체에 적용하고 고정할까요?`))return;
 state.settings=state.settings||{};state.settings.shippingFeeLocked=true;state.settings.shippingFee=fee;applyShippingFeeToAll();syncShippingFeeLockUI();
}
function unlockShippingFee(){if(!confirm('배송비 고정을 풀고 수정할까요?'))return;state.settings.shippingFeeLocked=false;saveSilently();syncShippingFeeLockUI()}
function syncShippingFeeLockUI(){const locked=!!state.settings?.shippingFeeLocked,el=$('shippingFee');if(!el)return;el.disabled=locked;el.classList.toggle('locked-input',locked);$('shippingFeeConfirmBtn').style.display=locked?'none':'';$('shippingFeeEditBtn').style.display=locked?'':'none';if(state.settings?.shippingFee!=null)el.value=state.settings.shippingFee}
function applyShippingFeeToAll(){
 const fee=Number($('shippingFee').value)||0;
 if(!state.orders.length){alert('먼저 판매리스트를 업로드해 주세요.');return}
 state.orders.forEach(o=>o.fee=fee);
 saveSilently();renderAll();
 alert(`전체 주문에 배송비 ${money(fee)}를 적용했습니다.`);
}

function addOrder(){
 const date=$('broadcastDate').value||today();
 const nick=prompt('고객 닉네임을 입력해 주세요.'); if(!nick)return;
 const item=prompt('상품명을 입력해 주세요.'); if(!item)return;
 const qty=Number(prompt('수량을 입력해 주세요.','1'))||1;
 const unit=Number(prompt('단가를 입력해 주세요.','0'))||0;
 const fee=Number($('shippingFee').value)||0;
 state.orders.push({id:crypto.randomUUID(),date,nick:nick.trim(),item:item.trim(),qty,unit,amount:qty*unit,fee,source:'직접등록'});
 autoMatchAll();saveSilently();renderAll();go('orders');
}
function matchBadge(o){return o.customerId?'<span class="badge good">DB 매칭</span>':o.matchStatus==='duplicate'?'<span class="badge warn">중복 후보</span>':'<span class="badge bad">정보 확인 필요</span>'}
let smsTemplateDirty=false;
const SMS_DRAFT_KEY='firstOmsSmsTemplateDraftV1';
function renderMessageSettings(force=false){
 const x=state.settings||{};
 if($('setBank')&&(force||document.activeElement!==$('setBank')))$('setBank').value=x.bank||'카카오뱅크';
 if($('setHolder')&&(force||document.activeElement!==$('setHolder')))$('setHolder').value=x.holder||'김미숙';
 if($('setAccount')&&(force||document.activeElement!==$('setAccount')))$('setAccount').value=x.account||'';
 if($('setContact')&&(force||document.activeElement!==$('setContact')))$('setContact').value=x.contact||'';
 const ta=$('setSmsTemplate');
 if(ta&&!smsTemplateDirty&&document.activeElement!==ta){
   const draft=localStorage.getItem(SMS_DRAFT_KEY);
   ta.value=draft!==null?draft:(x.smsTemplate||'');
 }
}
function initSmsTemplateGuard(){
 const ta=$('setSmsTemplate');if(!ta||ta.dataset.guard)return;ta.dataset.guard='1';
 const draft=localStorage.getItem(SMS_DRAFT_KEY);if(draft!==null){ta.value=draft;smsTemplateDirty=true;}
 ta.addEventListener('input',()=>{smsTemplateDirty=true;localStorage.setItem(SMS_DRAFT_KEY,ta.value);if($('smsDraftStatus'))$('smsDraftStatus').textContent='입력 내용 임시보관 중 · 저장 버튼을 누르면 서버 설정에 반영됩니다.';});
}
function saveMessageSettings(){
 state.settings={bank:$('setBank').value.trim()||'카카오뱅크',holder:$('setHolder').value.trim()||'김미숙',account:$('setAccount').value.trim(),contact:$('setContact').value.trim(),smsTemplate:$('setSmsTemplate').value};
 smsTemplateDirty=false;localStorage.removeItem(SMS_DRAFT_KEY);saveSilently();renderReceipts();if($('smsDraftStatus'))$('smsDraftStatus').textContent='저장 완료';alert('문자내용과 입금계좌 설정을 저장했습니다.');
}
function formatSms(r){
 const c=r.customer||{}, x=state.settings||{};
 const items=r.items.map(v=>`${v.item} ${v.qty}개`).join(', ');
 const values={'고객명':c.name||r.nick,'닉네임':r.nick,'상품내역':items,'상품합계':money(r.subtotal),'배송비':money(r.fee),'결제금액':money(r.total),'은행':x.bank||'카카오뱅크','예금주':x.holder||'김미숙','계좌번호':x.account||'계좌번호 미설정','문의안내':x.contact||''};
 let text=x.smsTemplate||state.settings.smsTemplate;
 Object.entries(values).forEach(([k,v])=>text=text.split('{'+k+'}').join(v));
 return text;
}
function clearCustomerSearch(){$('customerSearch').value='';renderCustomers()}
function registerFromSearch(){
 const q=($('customerSearch').value||'').trim();openCustomerModal();
 if(!q)return;
 if(/^01\d[\d-]+$/.test(q))$('cPhone').value=q;
 else $('cName').value=q;
}
function renderCustomers(){
 const q=norm($('customerSearch')?.value||'');
 const list=state.customers.map((c,i)=>({c,i})).filter(({c})=>!q||[c.name,c.nick,c.phone,c.address,c.detailAddress,c.memo].some(v=>norm(v).includes(q)));
 if($('customerSearchSummary'))$('customerSearchSummary').textContent=q?`검색 결과 ${list.length}명 / 전체 ${state.customers.length}명`:`전체 고객 ${state.customers.length}명`;
 if(!list.length){$('customersTable').innerHTML=`<div class="empty">${q?'검색 결과가 없습니다. 위의 ‘검색어로 고객 등록’을 눌러 바로 등록하세요.':'고객DB를 업로드하거나 직접 등록해주세요.'}</div>`;return}
 $('customersTable').innerHTML='<div class="scroll"><table><thead><tr><th>실명</th><th>닉네임</th><th>전화번호</th><th>주소</th><th>요청사항</th><th>관리</th></tr></thead><tbody>'+ 
 list.map(({c,i})=>`<tr><td>${esc(c.name)}</td><td>${esc(c.nick)}</td><td>${esc(c.phone)}</td><td class="address-cell">${esc([c.postalCode,c.address,c.detailAddress].filter(Boolean).join(' '))}</td><td class="address-cell">${esc(c.memo||'')}</td><td class="action-cell"><button class="btn secondary" onclick="editCustomer(${i})">수정</button> <button class="btn bad" onclick="deleteCustomer(${i})">삭제</button></td></tr>`).join('')+
 '</tbody></table></div>';
}
function filteredReceipts(){
 const q=norm($('receiptSearch')?.value||''), d=$('receiptDate')?.value||'';
 return getReceipts().filter(r=>{
  const c=r.customer||{};
  return (!d||r.date===d)&&(!q||[r.nick,c.name,c.nickname,c.nick].some(v=>norm(v).includes(q)));
 });
}
function clearReceiptFilters(){if($('receiptSearch'))$('receiptSearch').value='';if($('receiptDate'))$('receiptDate').value='';renderReceipts()}
function renderReceipts(){
 initSmsTemplateGuard();
 const rs=filteredReceipts().sort((a,b)=>String(a.nick||'').localeCompare(String(b.nick||''),'ko'));
 const matched=rs.filter(r=>r.customer).length;
 $('receiptSummary').textContent=`검색된 정산서 ${rs.length}장 · 고객DB 매칭 ${matched}장 · 정보확인 필요 ${rs.length-matched}장`;
 if(!rs.length){$('receiptCards').innerHTML='<div class="empty">조건에 맞는 정산서가 없습니다.</div>';updateKpi();return;}
 $('receiptCards').innerHTML=`<div class="scroll" style="grid-column:1/-1;width:100%"><table class="receipt-list-table"><colgroup><col><col><col><col><col><col><col><col><col><col></colgroup><thead><tr><th>중요</th><th>닉네임</th><th>이름</th><th>연락처</th><th>주소</th><th>합계</th><th>입금</th><th>고객정보</th><th>문자전송</th><th>상세</th></tr></thead><tbody>${rs.map(r=>{
   const c=r.customer||{};const paid=r.payment?.status==='paid';
   const infoBtn=c&&c.id?`<button class="btn secondary" onclick="editCustomerById('${esc(c.id)}')">정보수정</button>`:`<button class="btn warn" onclick="openCustomerForKey('${esc(r.key)}')">정보등록</button>`;
   const smsBtn=c.phone?`<button class="btn" onclick="sendMmsByKey('${esc(r.key)}',this)">문자전송</button>`:`<button class="btn secondary" disabled>전화번호 없음</button>`;
   const cleanName=safeCustomerText(c.name);const cleanPhone=safeCustomerText(c.phone);const cleanAddress=safeCustomerText([c.postalCode,c.address,c.detailAddress].map(cleanBrokenText).filter(Boolean).join(' '));
   return `<tr class="${r.starred?'important-row':''}"><td><button class="btn star-btn ${r.starred?'active':''}" aria-pressed="${r.starred?'true':'false'}" title="중요 고객 표시" onclick="toggleReceiptStar('${esc(r.key)}',this)">${r.starred?'★':'☆'}</button></td><td class="nick" title="${esc(cleanBrokenText(r.nick))}">${esc(cleanBrokenText(r.nick))}</td><td>${esc(cleanName)}</td><td>${esc(cleanPhone)}</td><td class="address-cell" title="${esc(cleanAddress)}">${esc(cleanAddress)}</td><td><b>${money(r.total)}</b></td><td>${paid?'<span class="badge good">입금</span>':'<span class="badge bad">미입금</span>'}</td><td>${infoBtn}</td><td>${smsBtn}</td><td><button class="btn secondary" onclick="openReceiptDetail('${esc(r.key)}')">상세</button></td></tr>`;
 }).join('')}</tbody></table></div>`;
 updateKpi();
}
function toggleReceiptStar(key,btn){
 const receipt=getReceipts().find(r=>r.key===key);if(!receipt)return;
 const ids=receipt.items.map(x=>x.id),next=!receipt.starred;
 state.orders.forEach(o=>{if(ids.includes(o.id))o.starred=next});
 if(btn){btn.textContent=next?'★':'☆';btn.classList.toggle('active',next);btn.setAttribute('aria-pressed',next?'true':'false')}
 saveSilently();renderReceipts();
}
async function openReceiptDetail(key){
 const r=receiptByKey(key);if(!r)return alert('정산서를 찾을 수 없습니다.');
 let modal=document.getElementById('receiptDetailModal');
 if(!modal){modal=document.createElement('div');modal.id='receiptDetailModal';modal.className='receipt-detail-modal';modal.innerHTML='<div class="receipt-detail-box" style="max-width:920px"><div class="receipt-detail-actions no-print"><button class="btn warn" id="detailEditBtn">수정</button><button class="btn secondary" id="detailPrintBtn">인쇄</button><button class="btn bad" onclick="closeReceiptDetail()">닫기</button></div><div id="receiptDetailBody"><div class="empty">정산서 이미지를 만드는 중입니다.</div></div></div>';document.body.appendChild(modal);modal.addEventListener('click',e=>{if(e.target===modal)closeReceiptDetail()});}
 modal.classList.add('show');
 document.getElementById('detailEditBtn').onclick=()=>openReceiptEditByKey(key);
 document.getElementById('detailPrintBtn').onclick=()=>printByKey(key);
 try{const base64=await receiptImageBase64(r);document.getElementById('receiptDetailBody').innerHTML=`<img class="receipt-image-preview" alt="${esc(r.nick)} 정산서 이미지" src="data:image/jpeg;base64,${base64}">`;}catch(e){document.getElementById('receiptDetailBody').innerHTML=receiptHTML(r,indexOfReceiptKey(key));}
}
function closeReceiptDetail(){document.getElementById('receiptDetailModal')?.classList.remove('show');}

function receiptHTML(r,i){
 const c=r.customer, pay=r.payment;
 const payBadge=pay.status==='paid'?'<span class="badge good">입금확인</span>':pay.status==='amount-mismatch'?'<span class="badge warn">금액 확인</span>':'<span class="badge bad">미입금</span>';
 const account=state.settings?.account||'계좌번호를 설정해 주세요';
 return `<article class="receipt" data-key="${esc(r.key)}">
 <h3>땡라이브 정산서</h3>
 <div class="meta">정산번호 ${r.date.replaceAll('-','')}-${String(i+1).padStart(3,'0')}<br>방송일 ${r.date}</div>
 <div style="font-size:18px"><b>${esc(r.nick)}</b> ${payBadge}<br><span style="font-size:13px;color:#666">${c?esc(c.name||''):`<button class="btn bad no-print" onclick="openCustomerForKey('${esc(r.key)}')">고객정보 확인필요 · 폼 열기</button>`}</span></div>
 <div class="meta">${esc(c?.phone||'연락처 미등록')}<br>${esc([c?.address,c?.detailAddress].filter(Boolean).join(' ')||'주소 미등록')}</div>
 <table><thead><tr><th>상품</th><th>수량</th><th>단가</th><th>금액</th></tr></thead><tbody>
 ${r.items.map(x=>`<tr><td>#${esc(x.number||'')} ${esc(x.item)}</td><td>${x.qty}</td><td>${money(x.unit)}</td><td>${money(x.amount)}</td></tr>`).join('')}
 </tbody></table>
 <div style="margin-top:10px;font-size:13px">상품합계 <b>${money(r.subtotal)}</b><br>배송비 <b>${money(r.fee)}</b></div>
 <div class="total">결제금액 ${money(r.total)}</div>
 <div class="account-box"><b>입금계좌</b><br>${esc(state.settings?.bank||'카카오뱅크')} <span class="copy-account" onclick="copyAccount()" title="눌러서 계좌번호 복사">${esc(account)}</span> <button class="btn secondary no-print" style="padding:4px 8px" onclick="copyAccount()">복사</button><br>예금주 ${esc(state.settings?.holder||'김미숙')}</div>
 <div class="actions no-print" style="margin-top:12px">
  <button class="btn" onclick="${c?.phone?`sendMmsByKey('${esc(r.key)}',this)`:`openCustomerForKey('${esc(r.key)}')`}">${c?.phone?'정산서 이미지 전송':'고객정보 등록 후 이미지 전송'}</button>
  <button class="btn warn" onclick="openReceiptEditByKey('${esc(r.key)}')">정산서 수정</button>
  <button class="btn secondary" onclick="printByKey('${esc(r.key)}')">이 정산서 인쇄</button>
 </div>
 <div class="foot">FIRST OMS · 주문내역을 확인해 주세요.</div>
 </article>`;
}
function copyAccount(){const a=state.settings?.account||'';if(!a)return alert('계좌번호를 먼저 설정해 주세요.');navigator.clipboard.writeText(a).then(()=>alert('계좌번호를 복사했습니다.')).catch(()=>prompt('계좌번호를 복사하세요.',a))}
function indexOfReceiptKey(key){return getReceipts().findIndex(r=>r.key===key)}
function openCustomerForKey(key){const i=indexOfReceiptKey(key);if(i>=0)openCustomerForReceipt(i)}
function openReceiptEditByKey(key){const i=indexOfReceiptKey(key);if(i>=0)openReceiptEdit(i)}
function printByKey(key){const i=indexOfReceiptKey(key);if(i>=0)printOne(i)}

function receiptByKey(key){return getReceipts().find(r=>r.key===key)||null}
function openReceiptEdit(i){
 const r=getReceipts()[i]; if(!r)return;
 $('receiptEditModal').dataset.key=r.key;
 $('receiptEditTitle').textContent=`${r.date} · ${r.customer?.name||r.nick} (${r.nick})`;
 renderReceiptEditRows();
 $('receiptEditModal').classList.add('show');
}
function closeReceiptEdit(){$('receiptEditModal').classList.remove('show')}
function renderReceiptEditRows(){
 const key=$('receiptEditModal').dataset.key, r=receiptByKey(key); if(!r){closeReceiptEdit();return}
 $('receiptEditFee').value=Number(r.fee)||0;
 $('receiptEditRows').innerHTML=r.items.map(o=>`<div class="receipt-edit-row" data-id="${esc(o.id)}">
   <div class="item-field"><label>상품명</label><input data-field="item" value="${esc(o.item)}"></div>
   <div><label>수량</label><input data-field="qty" type="number" min="0" value="${Number(o.qty)||0}"></div>
   <div><label>단가</label><input data-field="unit" type="number" min="0" value="${Number(o.unit)||0}"></div>
   <div><label>금액</label><input data-field="amount" type="number" min="0" value="${Number(o.amount)||0}"></div>
   <div class="delete-field"><label>&nbsp;</label><button class="btn bad" onclick="deleteReceiptItem('${esc(o.id)}')">삭제</button></div>
  </div>`).join('');
}
function addReceiptItem(){
 const key=$('receiptEditModal').dataset.key, r=receiptByKey(key); if(!r)return;
 const fee=Number($('receiptEditFee').value)||0;
 state.orders.push({id:crypto.randomUUID(),date:r.date,nick:r.nick,item:'새 상품',qty:1,unit:0,amount:0,fee,customerId:r.customerId||null,matchStatus:r.matchStatus||'unmatched',source:'정산서 수정'});
 renderReceiptEditRows();
}
function deleteReceiptItem(id){
 const r=receiptByKey($('receiptEditModal').dataset.key); if(!r)return;
 if(r.items.length<=1){alert('정산서에는 상품이 최소 1개 있어야 합니다. 정산서 전체를 지우려면 판매리스트에서 주문을 삭제해 주세요.');return}
 if(!confirm('이 상품을 정산서에서 삭제할까요?'))return;
 state.orders=state.orders.filter(o=>o.id!==id);
 renderReceiptEditRows();
}
function saveReceiptEdit(){
 const key=$('receiptEditModal').dataset.key, r=receiptByKey(key); if(!r)return;
 const fee=Number($('receiptEditFee').value)||0;
 document.querySelectorAll('#receiptEditRows .receipt-edit-row').forEach(row=>{
  const o=state.orders.find(x=>x.id===row.dataset.id); if(!o)return;
  const item=row.querySelector('[data-field="item"]').value.trim();
  const qty=Math.max(0,Number(row.querySelector('[data-field="qty"]').value)||0);
  const unit=Math.max(0,Number(row.querySelector('[data-field="unit"]').value)||0);
  const amountRaw=row.querySelector('[data-field="amount"]').value;
  const parsedAmount=Number(amountRaw);
  o.item=item||'상품명 미입력';o.qty=qty;o.unit=unit;o.amount=Number.isFinite(parsedAmount)?Math.max(0,parsedAmount):qty*unit;o.fee=fee;
 });
 state.orders.filter(o=>o.date===r.date&&norm(o.nick)===norm(r.nick)).forEach(o=>o.fee=fee);
 autoMatchAll();saveSilently();renderAll();closeReceiptEdit();alert('정산서 수정내용을 저장했습니다.');
}

function kakaoVariables(r){
 const c=r.customer||{}, x=state.settings||{};
 return {
  '#{고객명}':c.name||r.nick,
  '#{닉네임}':r.nick,
  '#{상품내역}':r.items.map(v=>`${v.item} ${v.qty}개`).join(', '),
  '#{상품합계}':money(r.subtotal),
  '#{배송비}':money(r.fee),
  '#{결제금액}':money(r.total),
  '#{은행}':x.bank||'카카오뱅크',
  '#{예금주}':x.holder||'김미숙',
  '#{계좌번호}':x.account||'계좌번호 미설정',
  '#{문의안내}':x.contact||''
 };
}

function wrapText(ctx,text,x,y,maxWidth,lineHeight,maxLines=99){
 const words=String(text||'').split(/\s+/);let line='',lines=0;
 for(const w of words){const test=line?line+' '+w:w;if(ctx.measureText(test).width>maxWidth&&line){ctx.fillText(line,x,y);y+=lineHeight;lines++;line=w;if(lines>=maxLines)return y}else line=test}
 if(line&&lines<maxLines){ctx.fillText(line,x,y);y+=lineHeight}return y;
}
async function receiptImageBase64(r){
 const W=900,H=1350,c=document.createElement('canvas');c.width=W;c.height=H;const x=c.getContext('2d');
 x.fillStyle='#fff';x.fillRect(0,0,W,H);x.fillStyle='#111';x.textAlign='center';x.font='bold 34px sans-serif';x.fillText('땡라이브 정산서',W/2,55);
 x.textAlign='left';x.font='22px sans-serif';let y=100;x.fillText(`방송일  ${r.date}`,55,y);y+=45;x.font='bold 34px sans-serif';x.fillText(r.nick,55,y);y+=35;x.font='20px sans-serif';x.fillStyle='#555';x.fillText(r.customer?.name||'고객정보 확인필요',55,y);y+=38;x.fillStyle='#111';x.font='18px sans-serif';x.fillText(r.customer?.phone||'',55,y);y+=30;y=wrapText(x,[r.customer?.address,r.customer?.detailAddress].filter(Boolean).join(' '),55,y,790,25,2)+15;
 x.strokeStyle='#bbb';x.beginPath();x.moveTo(55,y);x.lineTo(845,y);x.stroke();y+=35;x.font='bold 20px sans-serif';x.fillText('상품',55,y);x.fillText('수량',590,y);x.fillText('금액',730,y);y+=28;x.font='18px sans-serif';
 for(const item of r.items.slice(0,16)){y=wrapText(x,item.item,55,y,500,24,2);x.fillText(String(item.qty),610,y-24);x.fillText(money(item.amount),730,y-24);y+=10;if(y>920)break}
 y=Math.max(y+15,980);x.strokeStyle='#bbb';x.beginPath();x.moveTo(55,y);x.lineTo(845,y);x.stroke();y+=42;x.font='20px sans-serif';x.fillText(`상품합계  ${money(r.subtotal)}`,55,y);y+=34;x.fillText(`배송비  ${money(r.fee)}`,55,y);y+=48;x.font='bold 32px sans-serif';x.fillText(`결제금액  ${money(r.total)}`,55,y);y+=60;x.fillStyle='#f4eee9';x.fillRect(45,y-30,810,125);x.fillStyle='#111';x.font='bold 22px sans-serif';x.fillText('입금계좌',70,y);y+=36;x.font='22px sans-serif';x.fillText(`${state.settings?.bank||'카카오뱅크'} ${state.settings?.account||'계좌번호 미설정'}`,70,y);y+=34;x.fillText(`예금주 ${state.settings?.holder||'김미숙'}`,70,y);
 x.textAlign='center';x.fillStyle='#777';x.font='16px sans-serif';x.fillText('FIRST OMS · 주문내역을 확인해 주세요.',W/2,1320);
 let q=.78,data=c.toDataURL('image/jpeg',q);while(data.length>265000&&q>.35){q-=.08;data=c.toDataURL('image/jpeg',q)}return data.split(',')[1];
}
async function sendMmsByKey(key,button,skipConfirm=false){
 const r=receiptByKey(key),c=r?.customer;if(!r||!c?.phone)throw new Error('고객 전화번호가 없습니다.');
 if(!skipConfirm&&!confirm(`${r.nick}님 (${c.phone})에게 정산서 이미지를 MMS로 전송할까요?`))return false;
 const old=button?.textContent||'정산서 이미지 전송';if(button){button.disabled=true;button.textContent='이미지 만드는 중...'}
 try{const imageBase64=await receiptImageBase64(r);if(button)button.textContent='전송 중...';const response=await fetch('/api/mms/send',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({to:c.phone,imageBase64,subject:'땡라이브 정산서',text:`${r.nick}님 정산서입니다. 총 결제금액 ${money(r.total)}`,date:r.date,nickname:r.nick,name:c.name||'',total:r.total})});const data=await response.json();if(!response.ok||!data.ok)throw new Error(data.error||`MMS API 오류 ${response.status}`);if(!skipConfirm)alert(`${r.nick}님에게 정산서 이미지 전송을 접수했습니다.`);return true}catch(e){if(!skipConfirm)alert('정산서 이미지 전송 실패: '+e.message);throw e}finally{if(button){button.disabled=false;button.textContent=old}}}
async function sendAllMms(button){const rs=filteredReceipts().filter(r=>r.customer?.phone);if(!rs.length)return alert('전화번호가 등록된 고객 정산서가 없습니다.');if(!confirm(`현재 검색·날짜 조건의 ${rs.length}명에게 정산서 이미지를 전체 전송할까요?`))return;const old=button.textContent;button.disabled=true;let ok=0,fail=[];for(let n=0;n<rs.length;n++){button.textContent=`전송 ${n+1}/${rs.length}`;try{await sendMmsByKey(rs[n].key,null,true);ok++}catch(e){fail.push(`${rs[n].nick}: ${e.message}`)}}button.disabled=false;button.textContent=old;alert(`이미지 전체전송 완료\n성공 ${ok}명 / 실패 ${fail.length}명${fail.length?'\n\n'+fail.slice(0,10).join('\n'):''}`);loadSendHistory()}
async function loadSendHistory(){const box=$('sendHistoryBox'),list=$('sendHistoryList');box.style.display='block';list.innerHTML='불러오는 중...';try{const d=$('receiptDate')?.value||'';const r=await fetch('/api/send-history'+(d?`?date=${encodeURIComponent(d)}`:''),{cache:'no-store'}),j=await r.json();const rows=j.history||[];list.innerHTML=rows.length?rows.map(h=>`<div class="history-row"><span>${esc(new Date(h.sentAt).toLocaleString('ko-KR'))}</span><b>${esc(h.nickname||h.name||'')}</b><span>${esc(h.toMasked||'')} · ${money(h.total||0)}</span><span>${esc(h.date||'')}</span><span class="badge ${h.ok?'good':'bad'}">${h.ok?'성공':'실패'}</span></div>`).join(''):'<div class="empty">전송기록이 없습니다.</div>'}catch(e){list.innerHTML=`<div class="empty">기록 조회 실패: ${esc(e.message)}</div>`}}
async function sendKakao(i,button,skipConfirm=false){
 const r=getReceipts()[i], c=r?.customer;
 if(!c?.phone)throw new Error(`${r?.nick||'고객'} 전화번호가 없습니다.`);
 const msg=formatSms(r);
 if(!skipConfirm&&!confirm(`${c.name||r.nick}님 (${c.phone})에게 카카오 알림톡을 전송할까요?\n\n${msg}`))return false;
 const old=button?.textContent||'카카오 알림톡 전송';
 try{
  if(button){button.disabled=true;button.textContent='전송 중...'}
  const response=await fetch('/api/kakao/send',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({to:c.phone,text:msg,variables:kakaoVariables(r)})});
  const data=await response.json().catch(()=>({}));
  if(!response.ok||!data.ok)throw new Error(data.error||`카카오 API 오류 ${response.status}`);
  if(button)button.textContent='전송완료';
  if(!skipConfirm)alert(`${c.name||r.nick}님에게 카카오 알림톡 전송을 접수했습니다.`);
  return true;
 }catch(e){
  if(button){button.disabled=false;button.textContent=old}
  if(!skipConfirm)alert('카카오 알림톡 전송 실패: '+e.message);
  throw e;
 }
}
async function sendAllKakao(button){
 const rs=getReceipts().filter(r=>r.customer?.phone);
 if(!rs.length){alert('전화번호가 등록된 고객 정산서가 없습니다.');return}
 if(!confirm(`전화번호가 있는 ${rs.length}명에게 카카오 알림톡을 전체 전송할까요?`))return;
 const old=button.textContent;button.disabled=true;let ok=0,fail=[];
 for(const r of rs){
  const i=getReceipts().findIndex(x=>x.key===r.key);
  button.textContent=`전송 중 ${ok+fail.length+1}/${rs.length}`;
  try{await sendKakao(i,null,true);ok++}catch(e){fail.push(`${r.customer?.name||r.nick}: ${e.message}`)}
  await new Promise(resolve=>setTimeout(resolve,250));
 }
 button.disabled=false;button.textContent=old;
 alert(`카카오 알림톡 전체전송 완료\n성공 ${ok}명 / 실패 ${fail.length}명${fail.length?'\n\n'+fail.slice(0,10).join('\n'):''}`);
}
async function sendAllSms(button){
 const rs=getReceipts().filter(r=>r.customer?.phone);
 if(!rs.length){alert('전화번호가 등록된 고객 정산서가 없습니다.');return}
 if(!confirm(`전화번호가 있는 ${rs.length}명에게 문자를 전체 전송할까요?`))return;
 const old=button.textContent;button.disabled=true;let ok=0,fail=[];
 for(const r of rs){
  const c=r.customer,msg=formatSms(r);
  button.textContent=`전송 중 ${ok+fail.length+1}/${rs.length}`;
  try{
   const response=await fetch('/api/sms/send',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({to:c.phone,text:msg})});
   const data=await response.json().catch(()=>({}));
   if(!response.ok||!data.ok)throw new Error(data.error||`문자 API 오류 ${response.status}`);
   ok++;
  }catch(e){fail.push(`${c.name||r.nick}: ${e.message}`)}
  await new Promise(resolve=>setTimeout(resolve,250));
 }
 button.disabled=false;button.textContent=old;
 alert(`문자 전체전송 완료\n성공 ${ok}명 / 실패 ${fail.length}명${fail.length?'\n\n'+fail.slice(0,10).join('\n'):''}`);
}
async function sendSms(i,button){
 const r=getReceipts()[i], c=r.customer;
 if(!c?.phone){alert('고객 전화번호가 없습니다.');return}
 const msg=formatSms(r);
 if(!confirm(`${c.name||r.nick}님 (${c.phone})에게 정산서 문자를 바로 전송할까요?\n\n${msg}`))return;
 const old=button?.textContent||'솔라피 문자전송';
 try{
  if(button){button.disabled=true;button.textContent='전송 중...'}
  const response=await fetch('/api/sms/send',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({to:c.phone,text:msg})});
  const data=await response.json().catch(()=>({}));
  if(!response.ok||!data.ok)throw new Error(data.error||`문자 API 오류 ${response.status}`);
  if(button)button.textContent='전송완료';
  alert(`${c.name||r.nick}님에게 문자 전송을 접수했습니다.`);
 }catch(e){
  if(button){button.disabled=false;button.textContent=old}
  alert('문자 전송 실패: '+e.message+'\n\nRender 환경변수와 솔라피 발신번호 승인·잔액을 확인해 주세요.');
 }
}

async function openSolapiSettings(){document.getElementById('solapiModal').classList.add('show');await loadSolapiSettings()}
function closeSolapiSettings(){document.getElementById('solapiModal').classList.remove('show')}
async function loadSolapiSettings(){
 const st=document.getElementById('solapiStatus');try{
  const r=await fetch('/api/solapi/config',{cache:'no-store'}),d=await r.json();if(!r.ok)throw new Error(d.error||'설정 조회 실패');
  document.getElementById('solapiApiKey').value=d.apiKey||'';document.getElementById('solapiApiSecret').value='';document.getElementById('solapiSender').value=d.sender||'';document.getElementById('solapiPfId').value=d.pfId||'';document.getElementById('solapiTemplateId').value=d.templateId||'';
  st.textContent=d.configured?'문자 발송 기본설정 완료'+(d.pfId&&d.templateId?' · 카카오 알림톡 설정 완료':' · 카카오 알림톡 ID 미설정'):'아직 솔라피 설정이 완료되지 않았습니다.';
 }catch(e){st.textContent='설정 확인 실패: '+e.message}
}
async function saveSolapiSettings(){
 const body={apiKey:document.getElementById('solapiApiKey').value.trim(),apiSecret:document.getElementById('solapiApiSecret').value.trim(),sender:document.getElementById('solapiSender').value.trim(),pfId:document.getElementById('solapiPfId').value.trim(),templateId:document.getElementById('solapiTemplateId').value.trim()};
 try{const r=await fetch('/api/solapi/config',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)}),d=await r.json();if(!r.ok||!d.ok)throw new Error(d.error||'저장 실패');alert('솔라피 API 설정을 저장했습니다.');await loadSolapiSettings()}catch(e){alert('솔라피 설정 저장 실패: '+e.message)}
}
function printOne(i){
 const html=getReceipts(); const card=document.querySelectorAll('.receipt')[i]; if(!card)return;
 const w=window.open('','_blank');w.document.write(`<html><head><title>정산서</title><style>body{font-family:sans-serif;padding:20px}.receipt{max-width:500px;margin:auto}.receipt table{width:100%;border-collapse:collapse}.receipt th,.receipt td{border:1px solid #ccc;padding:7px}.no-print{display:none}.total{text-align:right;font-size:20px;font-weight:bold}
/* v6.9 화면 밀림·열 겹침 긴급 수정 */
header{position:relative!important;top:auto!important;align-items:flex-start!important;flex-wrap:wrap!important;gap:12px!important}
header>div:first-child{flex:0 0 auto;min-width:230px}
header>.actions{flex:1 1 760px;justify-content:flex-end;align-items:center}
main{padding-top:18px}
.page th{position:static!important;top:auto!important}
#receipts .scroll,#payments .scroll,#customers .scroll{max-height:none!important;overflow-x:auto!important;overflow-y:visible!important}
#receiptCards{display:block!important;width:100%!important}
#receiptCards>.scroll{width:100%!important}
.receipt-list-table{width:1380px!important;min-width:1380px!important;max-width:none!important;table-layout:fixed!important}
.receipt-list-table col:nth-child(1){width:68px}.receipt-list-table col:nth-child(2){width:145px}.receipt-list-table col:nth-child(3){width:115px}.receipt-list-table col:nth-child(4){width:150px}.receipt-list-table col:nth-child(5){width:345px}.receipt-list-table col:nth-child(6){width:125px}.receipt-list-table col:nth-child(7){width:100px}.receipt-list-table col:nth-child(8){width:120px}.receipt-list-table col:nth-child(9){width:105px}.receipt-list-table col:nth-child(10){width:107px}
.receipt-list-table th,.receipt-list-table td{height:auto!important;min-height:52px;line-height:1.35;overflow:hidden!important;text-overflow:ellipsis!important}
.receipt-list-table td:nth-child(5){white-space:normal!important;word-break:keep-all!important;overflow-wrap:anywhere!important}
.receipt-list-table td:nth-child(8),.receipt-list-table td:nth-child(9),.receipt-list-table td:nth-child(10){overflow:visible!important;text-overflow:clip!important}
.receipt-list-table .btn{display:inline-flex!important;align-items:center;justify-content:center;min-width:86px!important;white-space:nowrap!important}
#paymentsTable table{width:100%!important;min-width:820px!important;table-layout:fixed!important}
#paymentsTable col:nth-child(1){width:25%}#paymentsTable col:nth-child(2){width:18%}#paymentsTable col:nth-child(3){width:27%}#paymentsTable col:nth-child(4){width:30%}
#paymentsTable th,#paymentsTable td{position:static!important;white-space:normal!important;word-break:keep-all!important;vertical-align:middle!important}
#customers table{width:1050px!important;min-width:1050px!important;table-layout:fixed!important}
#customers th,#customers td{position:static!important;top:auto!important}
.section-title{flex-wrap:wrap}.receipt-tools{position:relative!important;z-index:1!important}
@media(max-width:900px){header>.actions{justify-content:flex-start}.receipt-list-table{width:1300px!important;min-width:1300px!important}}

</style></head><body>${card.outerHTML}
<div id="solapiModal" class="modal"><div class="modal-box">
 <div class="section-title"><h3>솔라피 API 직접 입력·연결</h3><button class="btn secondary" onclick="closeSolapiSettings()">닫기</button></div>
 <div class="form-grid">
  <div><label>API Key</label><input id="solapiApiKey" autocomplete="off" placeholder="솔라피 API Key"></div>
  <div><label>API Secret</label><input id="solapiApiSecret" type="password" autocomplete="new-password" placeholder="저장된 값은 다시 표시되지 않습니다"></div>
  <div><label>승인 발신번호</label><input id="solapiSender" placeholder="하이픈 없이 숫자만"></div>
  <div><label>카카오 채널 PF ID</label><input id="solapiPfId" placeholder="알림톡 사용 시 입력"></div>
  <div><label>알림톡 템플릿 ID</label><input id="solapiTemplateId" placeholder="승인 템플릿 ID"></div>
 </div>
 <p id="solapiStatus" class="muted"></p>
 <div class="actions"><button class="btn" onclick="saveSolapiSettings()">API 입력값 저장·연결</button><button class="btn secondary" onclick="loadSolapiSettings()">연결상태 확인</button></div>
 <p class="note">API Secret은 관리자 서버에만 저장되며 화면에는 다시 표시하지 않습니다. Render 재배포에도 절대 잃지 않으려면 Render Environment에도 같은 값을 등록하거나 Persistent Disk를 연결하세요.</p>
</div></div>
</body></html>`);w.document.close();w.print();
}
const paymentAliases={payer:['입금자','입금자명','보낸분','거래내용','내용','적요','성명'],amount:['입금액','금액','거래금액','입금'],date:['거래일시','거래일','날짜','시간']};
function importPayments(rows,headerIndex=null,sheetName=''){
 const hi=headerIndex===null?findHeader(rows,paymentAliases):headerIndex, map=colMap(rows[hi]||[],paymentAliases);let n=0;
 rows.slice(hi+1).forEach(r=>{
  const payer=map.payer>=0?r[map.payer]:'', amount=Number(String(map.amount>=0?r[map.amount]:0).replace(/[^0-9.-]/g,''));
  if(!String(payer).trim()||!amount||amount<0)return;
  state.payments.push({id:crypto.randomUUID(),payer:String(payer).trim(),amount,date:String(map.date>=0?r[map.date]:'').trim()});n++;
 });
 saveSilently();renderAll();go('payments');alert((sheetName?sheetName+' 시트에서 ':'')+n+'건 입금내역을 불러왔습니다.');
}
function paymentRowsFromPopulate(workbook){
 const found=[];
 workbook.sheets().forEach(sheet=>{
  let rows=[];try{const used=sheet.usedRange();rows=used?used.value():[]}catch(e){rows=[]}
  if(!Array.isArray(rows))return;
  for(let i=0;i<Math.min(rows.length,150);i++){
   const map=colMap(rows[i]||[],paymentAliases);if(map.payer<0||map.amount<0)continue;
   let count=0;for(const r of rows.slice(i+1)){const payer=cleanBrokenText(r?.[map.payer]);const amount=Number(String(r?.[map.amount]??'').replace(/[^0-9.-]/g,''));if(payer&&amount>0)count++}
   found.push({rows,hi:i,count,sheet:sheet.name()});
  }
 });
 return found.sort((a,b)=>b.count-a.count)[0]||null;
}
function paymentRowsFromSheetJs(buffer){
 const wb=XLSX.read(buffer,{type:'array',cellDates:true});let best=null;
 for(const sn of wb.SheetNames){const ws=wb.Sheets[sn];if(!ws)continue;const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',blankrows:true,raw:false});
  for(let i=0;i<Math.min(rows.length,150);i++){const map=colMap(rows[i]||[],paymentAliases);if(map.payer<0||map.amount<0)continue;let count=0;
   for(const r of rows.slice(i+1)){const payer=cleanBrokenText(r[map.payer]);const amount=Number(String(r[map.amount]??'').replace(/[^0-9.-]/g,''));if(payer&&amount>0)count++;}
   if(!best||count>best.count)best={rows,hi:i,count,sheet:sn};
  }
 }
 return best;
}
async function readEncryptedPaymentWorkbook(buffer,password){
 if(!window.XlsxPopulate)throw new Error('암호 파일 처리 모듈을 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.');
 const wb=await XlsxPopulate.fromDataAsync(buffer,{password});
 return paymentRowsFromPopulate(wb);
}
function readPaymentWorkbook(file){
 const status=$('paymentUploadStatus');if(status)status.textContent='입금명세 파일을 확인하는 중입니다...';
 const fr=new FileReader();fr.onload=async e=>{const buffer=e.target.result;try{
   let best=null;
   try{best=paymentRowsFromSheetJs(buffer)}catch(normalErr){console.warn('일반 엑셀 읽기 실패, 암호 여부 확인',normalErr)}
   if(!best||best.count===0){
    const password=prompt('암호가 설정된 입금명세 파일이면 암호를 입력해 주세요.\n암호가 없는 파일이면 취소를 누르세요.');
    if(password===null||password==='')throw new Error('파일이 암호화되어 있거나 거래일시·거래금액·내용 열을 찾지 못했습니다. 암호 파일은 암호를 입력해야 합니다.');
    if(status)status.textContent='암호를 확인하고 입금내역을 복호화하는 중입니다...';
    best=await readEncryptedPaymentWorkbook(buffer,password);
   }
   if(!best||best.count===0)throw new Error('암호는 확인됐지만 거래일시·거래금액·내용(입금자) 열 또는 입금 건수를 찾지 못했습니다.');
   importPayments(best.rows,best.hi,best.sheet);if(status)status.textContent=`${best.sheet} 시트에서 입금 ${best.count}건을 등록하고 정산서와 대조했습니다.`;
 }catch(err){if(status)status.textContent='불러오기 실패: '+err.message;alert('입금명세 파일을 읽지 못했습니다: '+err.message)}finally{$('paymentFile').value=''}};fr.readAsArrayBuffer(file);
}
$('paymentFile').onchange=e=>{const f=e.target.files&&e.target.files[0];if(f)readPaymentWorkbook(f)};
let paymentFilter='all';
function setPaymentFilter(v,btn){paymentFilter=v;document.querySelectorAll('#paymentFilterBar .btn').forEach(x=>x.classList.toggle('active',x===btn));renderPayments()}
function paymentStatusLabel(st){return st==='paid'?'입금완료':st==='amount-mismatch'?'금액불일치':st==='review'?'정보확인필요':'미입금'}
function setManualPayment(key,status){state.paymentOverrides=state.paymentOverrides||{};const r=getReceipts().find(x=>x.key===key);if(!r)return;state.paymentOverrides[key]={status,payment:r.payment?.payment||null,updatedAt:new Date().toISOString()};saveSilently();renderAll()}
function renderPayments(){
 const rs=getReceipts();let rows=rs;if(paymentFilter!=='all')rows=rows.filter(r=>r.payment.status===paymentFilter);
 const counts={all:rs.length,paid:rs.filter(r=>r.payment.status==='paid').length,unpaid:rs.filter(r=>r.payment.status==='unpaid').length,'amount-mismatch':rs.filter(r=>r.payment.status==='amount-mismatch').length,review:rs.filter(r=>r.payment.status==='review').length};
 document.querySelectorAll('#paymentFilterBar [data-filter]').forEach(b=>{const k=b.dataset.filter;b.textContent=({all:'전체',paid:'입금자',unpaid:'미입금자','amount-mismatch':'금액불일치',review:'정보확인필요'}[k])+' '+counts[k]});
 const box=$('paymentsTable');if(!rows.length){box.innerHTML='<div class="empty">선택한 조건의 고객이 없습니다.</div>';return}
 box.innerHTML='<div class="scroll"><table style="min-width:1250px"><thead><tr><th><input type="checkbox" class="payment-check" onchange="document.querySelectorAll(\'.payment-row-check\').forEach(x=>x.checked=this.checked)"></th><th>고객명</th><th>입금자</th><th>입금액</th><th>거래일시</th><th>대조결과(수기변경)</th><th>문자전송</th><th>상세</th></tr></thead><tbody>'+rows.map(r=>{const p=r.payment?.payment||{};const c=r.customer||{};return `<tr><td><input class="payment-check payment-row-check" type="checkbox" value="${esc(r.key)}"></td><td><b>${esc(c.name||r.nick)}</b><br><span class="muted">${esc(r.nick)}</span></td><td>${esc(p.payer||'-')}</td><td>${p.amount!=null?money(p.amount):'-'}<br><span class="muted">청구 ${money(r.total)}</span></td><td>${esc(p.date||'-')}</td><td><select class="payment-manual-select" onchange="setManualPayment('${esc(r.key)}',this.value)"><option value="paid" ${r.payment.status==='paid'?'selected':''}>입금완료</option><option value="unpaid" ${r.payment.status==='unpaid'?'selected':''}>미입금</option><option value="amount-mismatch" ${r.payment.status==='amount-mismatch'?'selected':''}>금액불일치</option><option value="review" ${r.payment.status==='review'?'selected':''}>정보확인필요</option></select></td><td><button class="btn" onclick="sendPaymentStatusSms('${esc(r.key)}',this)">문자전송</button></td><td><button class="btn secondary" onclick="openReceiptDetail('${esc(r.key)}')">정산서 보기</button></td></tr>`}).join('')+'</tbody></table></div>';
 const t=$('unpaidSmsTemplate');if(t&&!t.value)t.value=state.settings?.unpaidSmsTemplate||'{고객명}님, 현재 대조결과는 {대조결과}입니다. 확인금액 {금액}입니다. 다른 이름으로 입금하셨다면 입금자명을 알려주세요.';
}
async function sendPaymentStatusSms(key,button){const r=getReceipts().find(x=>x.key===key);if(!r)return alert('정산서를 찾을 수 없습니다.');const c=r.customer;if(!c?.phone)return alert('고객 연락처가 없습니다.');const tpl=$('unpaidSmsTemplate')?.value||'';const text=tpl.replaceAll('{고객명}',c.name||r.nick).replaceAll('{닉네임}',r.nick).replaceAll('{금액}',money(r.total)).replaceAll('{대조결과}',paymentStatusLabel(r.payment.status));if(!confirm(`${c.name||r.nick}님에게 아래 문자를 전송할까요?\n\n${text}`))return;const old=button.textContent;button.disabled=true;button.textContent='전송 중...';try{const response=await fetch('/api/sms/send',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({to:c.phone,text,date:r.date,nickname:r.nick,name:c.name||'',total:r.total})});const d=await response.json();if(!response.ok||!d.ok)throw new Error(d.error||'전송 실패');alert('문자를 전송했습니다.')}catch(e){alert('전송 실패: '+e.message)}finally{button.disabled=false;button.textContent=old}}
async function sendSelectedPaymentSms(button){const keys=[...document.querySelectorAll('.payment-row-check:checked')].map(x=>x.value);if(!keys.length)return alert('문자를 보낼 고객을 선택해 주세요.');if(!confirm(`선택한 ${keys.length}명에게 문자를 전송할까요?`))return;button.disabled=true;let ok=0,fail=0;for(const key of keys){const r=getReceipts().find(x=>x.key===key),c=r?.customer;if(!c?.phone){fail++;continue}const tpl=$('unpaidSmsTemplate')?.value||'';const text=tpl.replaceAll('{고객명}',c.name||r.nick).replaceAll('{닉네임}',r.nick).replaceAll('{금액}',money(r.total)).replaceAll('{대조결과}',paymentStatusLabel(r.payment.status));try{const resp=await fetch('/api/sms/send',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({to:c.phone,text,date:r.date,nickname:r.nick,name:c.name||'',total:r.total})});if(resp.ok)ok++;else fail++}catch(e){fail++}}button.disabled=false;alert(`문자전송 완료: 성공 ${ok}명 / 실패 ${fail}명`)}
function saveUnpaidTemplateDraft(){state.settings=state.settings||{};state.settings.unpaidSmsTemplate=$('unpaidSmsTemplate')?.value||'';saveSilently()}
async function sendUnpaidRequestByKey(key,button){
 const r=getReceipts().find(x=>x.key===key);if(!r)return alert('정산서를 찾을 수 없습니다.');const c=r.customer;if(!c?.phone)return alert('고객 연락처가 없습니다.');
 const tpl=($('unpaidSmsTemplate')?.value||state.settings?.unpaidSmsTemplate||'{고객명}님 입금이 확인되지 않아 입금요청 드립니다. 다른 이름으로 입금하신 경우, 입금자명 보내주시기 바랍니다.');
 const text=tpl.replaceAll('{고객명}',c.name||r.nick).replaceAll('{닉네임}',r.nick).replaceAll('{금액}',money(r.total));
 if(!confirm(`${c.name||r.nick}님에게 정산서 이미지와 아래 입금요청 문자를 전송할까요?

${text}`))return;
 const old=button.textContent;button.disabled=true;button.textContent='전송 중...';
 try{const imageBase64=await receiptImageBase64(r);const response=await fetch('/api/mms/send',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({to:c.phone,imageBase64,subject:'땡라이브 입금요청',text,date:r.date,nickname:r.nick,name:c.name||'',total:r.total})});const d=await response.json();if(!response.ok||!d.ok)throw new Error(d.error||'전송 실패');alert('입금 요청문자와 정산서 이미지를 전송했습니다.')}catch(e){alert('전송 실패: '+e.message)}finally{button.disabled=false;button.textContent=old}
}
function shippingCodeFor(g){
 const raw=[g.name,g.nick,g.phone,g.address].join('|');let h=2166136261;
 for(let i=0;i<raw.length;i++){h^=raw.charCodeAt(i);h=Math.imul(h,16777619)}
 return 'FIRST-'+(h>>>0).toString(36).toUpperCase().padStart(7,'0');
}
function renderShipping(){
 const filter=$('shippingFilter').value, bundle=$('bundleMode').value; let rs=getReceipts();if(filter==='paid')rs=rs.filter(r=>r.payment.status==='paid');
 const groups=new Map();rs.forEach(r=>{const k=bundle==='customer'?(r.customer?.id||r.nick):r.key;if(!groups.has(k))groups.set(k,{name:r.customer?.name||'',nick:r.nick,phone:r.customer?.phone||'',address:[r.customer?.postalCode,r.customer?.address,r.customer?.detailAddress].filter(Boolean).join(' '),dates:new Set(),items:[],subtotal:0,fee:0,total:0,status:r.payment.status});const g=groups.get(k);g.dates.add(r.date);g.items.push(...r.items);g.subtotal+=r.subtotal;g.fee+=r.fee;g.total+=r.total});
 const arr=[...groups.values()].map(g=>({...g,code:shippingCodeFor(g)}));state.shippingScans=state.shippingScans||{};
 $('shippingTable').innerHTML=arr.length?'<div class="scroll"><table class="shipping-table"><thead><tr><th>포장</th><th>택배코드</th><th>고객명</th><th>닉네임</th><th>연락처</th><th>주소</th><th>방송일</th><th aria-label="상품내역"></th><th>상품합계</th><th>배송비</th><th>토탈금액</th><th>송장번호</th></tr></thead><tbody>'+arr.map(g=>{const packed=!!state.shippingScans[g.code];return `<tr class="${packed?'packed-row':''}"><td>${packed?'<span class="badge good">포장완료</span><br><small>'+new Date(state.shippingScans[g.code].at).toLocaleString('ko-KR')+'</small>':'<span class="badge warn">대기</span>'}<br><button class="btn secondary" style="padding:5px 8px;margin-top:5px" onclick="togglePackingStatus('${esc(g.code)}')">${packed?'미완료로 변경':'수기 완료'}</button></td><td class="shipping-code">${esc(g.code)}</td><td>${esc(g.name)}</td><td>${esc(g.nick)}</td><td>${esc(g.phone)}</td><td class="address-cell">${esc(g.address)}</td><td>${[...g.dates].join(', ')}</td><td>${g.items.map((x,i)=>`#${esc(x.number||i+1)} ${esc(x.item)} / ${x.qty}개 / ${money(x.unit)} / ${money(x.amount)}`).join('<br>')}</td><td>${money(g.subtotal)}</td><td>${money(g.fee)}</td><td><b>${money(g.total)}</b></td><td><b>${esc(state.shippingScans?.[g.code]?.trackingNumber||'미발행')}</b></td></tr>`}).join('')+'</tbody></table></div>':'<div class="empty">출력할 택배 대상이 없습니다.</div>';window.currentShipping=arr;
}

async function togglePackingStatus(code){
 state.shippingScans=state.shippingScans||{};
 try{
  if(state.shippingScans[code]){
   if(!confirm('포장완료를 미완료로 변경할까요?'))return;
   const r=await fetch('/api/packing/status?code='+encodeURIComponent(code),{method:'DELETE'}),d=await r.json();
   if(!r.ok||!d.ok)throw new Error(d.error||'상태 변경 실패');
  }else{
   const worker=prompt('작업자 이름을 입력해 주세요. (선택사항)','관리자')||'관리자';
   const r=await fetch('/api/packing/status',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({code,worker,source:'admin-manual'})}),d=await r.json();
   if(!r.ok||!d.ok)throw new Error(d.error||'상태 변경 실패');
  }
  await syncShippingScans(true);
 }catch(e){alert('포장상태 변경 실패: '+e.message)}
}
async function processShippingScan(value){
 const input=$('shippingScanInput'),code=String(value||input?.value||'').trim().toUpperCase();
 const status=$('shippingScanStatus');if(!code){status.textContent='스캔할 코드가 없습니다.';return}
 const found=(window.currentShipping||[]).find(g=>g.code.toUpperCase()===code||String(g.phone||'').replace(/\D/g,'')===code.replace(/\D/g,''));
 if(!found){status.textContent='목록에서 일치하는 택배코드를 찾지 못했습니다: '+code;return}
 try{
  const rr=await fetch('/api/packing/status',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({code:found.code,worker:'관리자 스캔',source:'admin-scan'})}),dd=await rr.json();
  if(!rr.ok||!dd.ok)throw new Error(dd.error||'포장완료 저장 실패');
  await syncShippingScans(true);
  status.textContent=`포장완료 처리: ${found.name||found.nick} · ${found.code}`;if(input){input.value='';input.focus()}
 }catch(e){status.textContent='포장완료 저장 실패: '+e.message}
}
$('shippingScanInput')?.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();processShippingScan()}});
let shippingScanStream=null,shippingScanTimer=null;
async function startCameraScanner(){
 const status=$('shippingScanStatus'),video=$('shippingScanVideo');
 if(!('BarcodeDetector' in window)){status.textContent='이 브라우저는 카메라 바코드 인식을 지원하지 않습니다. USB 스캐너 또는 직접 입력을 사용해 주세요.';return}
 try{
  shippingScanStream=await navigator.mediaDevices.getUserMedia({video:{facingMode:{ideal:'environment'}}});video.srcObject=shippingScanStream;video.style.display='block';await video.play();
  const detector=new BarcodeDetector({formats:['qr_code','code_128','code_39','ean_13','ean_8']});status.textContent='카메라를 QR 또는 바코드에 맞춰주세요.';
  const tick=async()=>{try{const codes=await detector.detect(video);if(codes.length){processShippingScan(codes[0].rawValue);stopCameraScanner();return}}catch(e){}shippingScanTimer=requestAnimationFrame(tick)};tick();
 }catch(e){status.textContent='카메라를 열지 못했습니다: '+e.message}
}
function stopCameraScanner(){if(shippingScanTimer)cancelAnimationFrame(shippingScanTimer);shippingScanTimer=null;if(shippingScanStream)shippingScanStream.getTracks().forEach(t=>t.stop());shippingScanStream=null;const v=$('shippingScanVideo');if(v){v.pause();v.srcObject=null;v.style.display='none'}}
async function syncShippingScans(renderAfter=false){
 try{
  const r=await fetch('/api/packing/scans?ts='+Date.now(),{cache:'no-store'}),d=await r.json();
  if(!r.ok||!d.ok)throw new Error(d.error||'포장상태 조회 실패');
  state.shippingScans=d.shippingScans||{};
  localStorage.setItem(KEY,JSON.stringify(state));
  if(renderAfter)renderShipping();
  return true;
 }catch(e){console.warn('포장상태 동기화 실패',e);return false}
}

function normalizeTrackingHeader(v){return String(v??'').trim().toLowerCase().replace(/[\s_()\[\]\-]/g,'')}
function trackingField(row,names){for(const [k,v] of Object.entries(row||{})){const h=normalizeTrackingHeader(k);if(names.some(n=>h===normalizeTrackingHeader(n)||h.includes(normalizeTrackingHeader(n)))){if(String(v??'').trim())return v}}return ''}
function trackingPhone(v){return String(v??'').replace(/\D/g,'').replace(/^82/,'0')}
function trackingAddress(v){return String(v??'').toLowerCase().replace(/[\s,().-]/g,'')}
function downloadTrackingTemplate(){exportXlsx([{'주문번호(선택)':'','고객명':'홍길동','닉네임(선택)':'길동맘','연락처':'010-0000-0000','주소':'서울시 강남구 예시로 1','송장번호':'123456789012','택배사':'CJ대한통운'}],'FIRST_OMS_송장업로드_양식.xlsx','송장업로드')}
async function uploadTrackingFile(input){
 const file=input?.files?.[0];if(!file)return;
 try{
  renderShipping();const buf=await file.arrayBuffer(),wb=XLSX.read(buf,{type:'array',cellDates:true});let rows=[];
  for(const sn of wb.SheetNames){const r=XLSX.utils.sheet_to_json(wb.Sheets[sn],{defval:'',raw:false});if(r.length>rows.length)rows=r}
  if(!rows.length)throw new Error('송장 데이터가 없습니다.');
  const groups=window.currentShipping||[],updates=[],review=[];
  for(const [idx,row] of rows.entries()){
   const orderNo=String(trackingField(row,['주문번호','택배코드','코드'])||'').trim().toUpperCase();
   const name=String(trackingField(row,['고객명','받는분','수하인명','수취인명','이름'])||'').trim();
   const nick=String(trackingField(row,['닉네임','별명'])||'').trim();
   const phone=trackingPhone(trackingField(row,['연락처','전화번호','휴대폰','수하인전화번호','수취인전화번호']));
   const address=String(trackingField(row,['주소','배송지','수하인주소','수취인주소'])||'').trim();
   const trackingNumber=String(trackingField(row,['송장번호','운송장번호','등기번호','trackingnumber','invoice'])||'').replace(/[^0-9A-Za-z-]/g,'').trim();
   const courier=String(trackingField(row,['택배사','배송사','운송사'])||'파일접수').trim();
   if(!trackingNumber){review.push(`${idx+2}행: 송장번호 없음`);continue}
   let candidates=[];
   if(orderNo)candidates=groups.filter(g=>String(g.code||'').toUpperCase()===orderNo);
   if(!candidates.length&&phone)candidates=groups.filter(g=>trackingPhone(g.phone)===phone);
   if(candidates.length>1&&name)candidates=candidates.filter(g=>String(g.name||'').trim()===name);
   if(!candidates.length&&name&&address){const a=trackingAddress(address);candidates=groups.filter(g=>String(g.name||'').trim()===name&&(trackingAddress(g.address).includes(a)||a.includes(trackingAddress(g.address))))}
   if(!candidates.length&&name)candidates=groups.filter(g=>String(g.name||'').trim()===name);
   if(candidates.length>1&&nick)candidates=candidates.filter(g=>String(g.nick||'').trim()===nick);
   if(candidates.length!==1){review.push(`${idx+2}행: ${name||nick||phone||'고객'} 매칭 ${candidates.length?'중복':'실패'}`);continue}
   const g=candidates[0],existing=state.shippingScans?.[g.code]?.trackingNumber||'';
   if(existing&&existing!==trackingNumber){review.push(`${idx+2}행: ${g.name||g.nick} 기존 송장 충돌`);continue}
   updates.push({code:g.code,trackingNumber,courier});
  }
  if(!updates.length)throw new Error('자동 연결 가능한 송장이 없습니다.\n'+review.slice(0,5).join('\n'));
  const msg=`자동 연결 ${updates.length}건 / 확인 필요 ${review.length}건\n적용할까요?`;
  if(!confirm(msg))return;
  const resp=await fetch('/api/packing/tracking/bulk',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({updates})});const d=await resp.json();
  if(!resp.ok||!d.ok)throw new Error(d.error||'송장 저장 실패');
  await syncShippingScans(true);alert(`송장 등록 완료 ${d.updated}건${review.length?`\n확인 필요 ${review.length}건`:''}`);
 }catch(e){alert('송장 파일 업로드 실패: '+e.message)}finally{input.value=''}
}
function downloadShippingList(){renderShipping();const rows=(window.currentShipping||[]).map((g,i)=>({'번호':i+1,'포장상태':state.shippingScans?.[g.code]?'포장완료':'대기','택배코드':g.code,'받는분':g.name,'닉네임':g.nick,'연락처':g.phone,'주소':g.address,'방송일':[...g.dates].join(', '),'상품내역':g.items.map((x,j)=>`#${x.number||j+1} ${x.item} ${x.qty}개 단가 ${x.unit} 금액 ${x.amount}`).join(' / '),'상품합계':g.subtotal,'배송비':g.fee,'토탈금액':g.total,'송장번호':state.shippingScans?.[g.code]?.trackingNumber||'','택배사':state.shippingScans?.[g.code]?.courier||''}));exportXlsx(rows,'FIRST_OMS_택배리스트.xlsx','택배리스트')}
function printShipping(){renderShipping();const arr=window.currentShipping||[];if(!arr.length)return alert('출력할 택배 대상이 없습니다.');const data=JSON.stringify(arr.map(g=>({code:g.code,url:location.origin+'/packing.html?code='+encodeURIComponent(g.code)})));const pages=arr.map((g,i)=>`<section class="page"><header><div><h1>땡라이브 고객별 포장리스트</h1><b>${esc(g.name||g.nick)}</b> · ${esc(g.nick)}</div><div id="qr-${i}" class="qr"></div></header><div class="meta"><div><b>연락처</b> ${esc(g.phone)}</div><div><b>방송일</b> ${[...g.dates].join(', ')}</div><div class="address"><b>주소</b> ${esc(g.address)}</div><div><b>택배코드</b> ${esc(g.code)}</div><div><b>송장번호</b> ${esc(state.shippingScans?.[g.code]?.trackingNumber||'미발행')}</div></div><table><thead><tr><th>번호</th><th>상품명</th><th>수량</th><th>단가</th><th>금액</th><th>확인</th></tr></thead><tbody>${g.items.map((x,j)=>`<tr><td>#${esc(x.number||j+1)}</td><td>${esc(x.item)}</td><td>${x.qty}</td><td>${money(x.unit)}</td><td>${money(x.amount)}</td><td>□</td></tr>`).join('')}</tbody></table><div class="totals"><div>상품 합계 <b>${money(g.subtotal)}</b></div><div>택배비 <b>${money(g.fee)}</b></div><div class="grand">토탈 금액 <b>${money(g.total)}</b></div></div><footer>QR을 휴대폰 기본 카메라로 스캔하면 이 고객의 포장리스트가 열립니다.</footer></section>`).join('');const w=window.open('','_blank');w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>고객별 택배리스트</title><script src="https://cdn.jsdelivr.net/npm/qrcodejs@1.0.0/qrcode.min.js"><\/script><style>*{box-sizing:border-box}body{margin:0;font-family:Arial,'Noto Sans KR',sans-serif;color:#111}.page{width:210mm;min-height:297mm;padding:14mm;page-break-after:always;position:relative}.page:last-child{page-break-after:auto}header{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:3px solid #111;padding-bottom:12px}h1{font-size:25px;margin:0 0 8px}.qr img,.qr canvas{width:110px!important;height:110px!important}.meta{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:18px 0;font-size:14px}.address{grid-column:1/-1;line-height:1.5}table{width:100%;border-collapse:collapse;table-layout:fixed}th,td{border:1px solid #555;padding:11px 8px;text-align:center;word-break:keep-all}th:nth-child(2),td:nth-child(2){text-align:left;width:42%}.totals{margin:18px 0 0 auto;width:330px;font-size:16px;line-height:2;text-align:right}.grand{font-size:23px;border-top:2px solid #111;margin-top:4px;padding-top:5px}footer{position:absolute;bottom:12mm;left:14mm;right:14mm;text-align:center;color:#555}@page{size:A4 portrait;margin:0}</style><style id="v717-final-ui">
#shipping{background:#f7f8fa;color:#111;padding:24px;border-radius:18px}
#shipping .section-title h2{font-size:30px;color:#111}
#shipping .card,#shipping .shipping-scan-box{background:#fff;border:1px solid #e2e5e9;box-shadow:none;color:#111}
#shippingTable{margin-top:14px}.shipping-modern{width:100%;border-collapse:separate;border-spacing:0 7px;table-layout:fixed;font-size:13px}
.shipping-modern thead th{background:#f1f3f5;color:#222;padding:12px 6px;border-top:1px solid #ddd;border-bottom:1px solid #ddd}
.shipping-modern tbody td{background:#fff;border-top:1px solid #e2e5e9;border-bottom:1px solid #e2e5e9;padding:11px 7px;vertical-align:middle;word-break:keep-all;overflow-wrap:anywhere}
.shipping-modern tbody td:first-child{border-left:1px solid #e2e5e9;border-radius:10px 0 0 10px}.shipping-modern tbody td:last-child{border-right:1px solid #e2e5e9;border-radius:0 10px 10px 0}
.shipping-customer strong{display:block;font-size:21px;line-height:1.2;color:#111;font-weight:900}.shipping-customer .realname{font-size:16px;font-weight:800;margin-top:4px}.shipping-customer .phone{font-size:13px;margin-top:6px}
.ship-items{line-height:1.45}.ship-pay{border:1px solid #e5e5e5;border-radius:8px;padding:7px;background:#fafafa;line-height:1.55}.ship-pay .grand{font-weight:900;color:#d62828;font-size:15px}.ship-pay .mismatch{margin-top:5px;padding-top:5px;border-top:1px dashed #ccc;color:#b42318;font-weight:800}
.ship-address{line-height:1.45}.ship-status{display:inline-block;padding:5px 8px;border-radius:6px;font-weight:800;background:#fff3cd;color:#8a5b00}.ship-status.paid{background:#e7f7ea;color:#167a2f}.ship-status.mismatch{background:#ffe6e6;color:#b42318}
.ship-actions{display:grid;gap:5px}.ship-actions .btn{padding:6px 7px;font-size:12px;width:100%}
@media(max-width:1100px){#shipping{padding:12px}.shipping-modern{font-size:11px}.shipping-customer strong{font-size:17px}.shipping-modern tbody td{padding:8px 4px}}
.receipt-sheet{width:100%;max-width:850px;margin:0 auto;background:#fff;color:#111;padding:30px 34px 24px;font-family:Arial,'Noto Sans KR',sans-serif;box-sizing:border-box}
.receipt-sheet *{box-sizing:border-box}.rs-head{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:3px solid #111;padding-bottom:18px}.rs-brand{font-size:24px;font-weight:900}.rs-title{font-size:68px;font-weight:950;letter-spacing:18px;line-height:1;margin-top:12px}.rs-order{border:1px solid #111;border-collapse:collapse;font-size:16px}.rs-order td{border:1px solid #111;padding:10px 14px}.rs-customer-title{display:flex;align-items:center;gap:12px;margin:16px 0}.rs-page-label{background:#111;color:#fff;padding:9px 16px;border-radius:4px;font-size:19px;font-weight:900}.rs-customer-title strong{font-size:28px;font-weight:950}.rs-info{border:1px solid #222;border-radius:6px;display:grid;grid-template-columns:1fr 1.25fr;margin-bottom:16px}.rs-info>div{padding:13px 16px;border-bottom:1px solid #aaa;font-size:16px}.rs-info>div:nth-last-child(-n+2){border-bottom:none}.rs-info .label{display:inline-block;min-width:78px;font-weight:900}.rs-info .customer-name{font-size:22px;font-weight:950}.rs-table{width:100%;border-collapse:collapse;table-layout:fixed}.rs-table th{background:#050505;color:#fff;padding:10px;border:1px solid #555;font-size:16px}.rs-table td{border:1px solid #aaa;padding:10px;text-align:center;font-size:15px}.rs-table td:nth-child(2){text-align:left}.rs-summary{display:grid;grid-template-columns:1fr 1fr;border:1px solid #222;border-top:none}.rs-summary>div{padding:13px 20px;line-height:1.8}.rs-total{display:flex;align-items:center;justify-content:center;gap:20px;border-left:1px solid #222;font-weight:900;font-size:20px}.rs-total strong{font-size:46px}.rs-notice{border:1px solid #222;border-radius:6px;padding:14px 20px;margin-top:16px;font-size:18px}.rs-notice strong{font-size:24px;color:#c91515}.rs-bottom{display:grid;grid-template-columns:1.4fr .6fr;gap:20px;margin-top:18px}.rs-bank{border:1px solid #e0c14e;background:#fff8d8;border-radius:6px;padding:16px;font-size:18px;line-height:1.5}.rs-bank strong{font-size:27px}.rs-qr{text-align:center;border-left:1px solid #aaa;padding:10px}.rs-qr-placeholder{width:110px;height:110px;margin:5px auto;border:5px solid #111;display:grid;place-items:center;font-weight:900;font-size:18px}.rs-footer{border-top:2px solid #222;margin-top:18px;padding-top:13px;display:flex;justify-content:space-between;font-size:16px}.rs-tabs{margin-top:18px;border-top:1px dashed #777;padding-top:14px;display:flex;gap:10px}.rs-tab{border:1px solid #777;padding:8px 16px;border-radius:4px}.rs-tab.active{background:#111;color:#fff;font-weight:900}
@media print{body *{visibility:hidden!important}.receipt-sheet,.receipt-sheet *{visibility:visible!important}.receipt-sheet{position:absolute;left:0;top:0;width:210mm;max-width:none;min-height:297mm;padding:12mm}.no-print{display:none!important}}
</style>
</head><body>${pages}<script>const d=${data};window.onload=()=>{d.forEach((x,i)=>new QRCode(document.getElementById('qr-'+i),{text:x.url,width:180,height:180}));setTimeout(()=>window.print(),700)};<\/script></body></html>`);w.document.close()}
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
