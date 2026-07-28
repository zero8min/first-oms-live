
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

<script id="v717-final-functions">
function v717Esc(v){return esc(String(v??''))}
function v717PaymentAmounts(r){const paid=Number(r?.payment?.payment?.amount||0);const charged=Number(r?.total||0);return {charged,paid,diff:paid-charged}}
function v717StatusText(r){const s=r?.payment?.status;if(s==='paid')return '입금완료';if(s==='amount-mismatch')return '금액불일치';if(s==='review')return '정보확인필요';return '송장발행 대기'}
function renderShipping(){
 const filter=$('shippingFilter')?.value||'paid';let rs=getReceipts();if(filter==='paid')rs=rs.filter(r=>r.payment.status==='paid');
 state.shippingScans=state.shippingScans||{};
 const arr=rs.map(r=>{const c=r.customer||{};const g={key:r.key,name:c.name||'',nick:r.nick,phone:c.phone||'',address:[c.postalCode,c.address,c.detailAddress].filter(Boolean).join(' '),dates:new Set([r.date]),items:r.items,subtotal:r.subtotal,fee:r.fee,total:r.total,status:r.payment.status,payment:r.payment,customer:c};g.code=shippingCodeFor(g);return g});window.currentShipping=arr;
 const box=$('shippingTable');if(!arr.length){box.innerHTML='<div class="empty">출력할 택배 대상이 없습니다.</div>';return}
 box.innerHTML=`<table class="shipping-modern"><colgroup><col style="width:3%"><col style="width:15%"><col style="width:16%"><col style="width:13%"><col style="width:18%"><col style="width:10%"><col style="width:10%"><col style="width:7%"><col style="width:8%"></colgroup><thead><tr><th>선택</th><th>고객정보<br><small>(닉네임 / 이름)</small></th><th>주문정보<br><small>(상품/수량)</small></th><th>결제정보</th><th>배송지 / 연락처</th><th>배송상태</th><th>송장번호</th><th>주문일시</th><th>관리</th></tr></thead><tbody>${arr.map(g=>{const a=v717PaymentAmounts(g);const track=state.shippingScans?.[g.code]?.trackingNumber||'';const st=g.status==='amount-mismatch'?'mismatch':g.status==='paid'?'paid':'';return `<tr><td><input type="checkbox"></td><td class="shipping-customer"><strong>${v717Esc(g.nick)}</strong><div class="realname">(${v717Esc(g.name||'이름 미등록')})</div><div class="phone">☎ ${v717Esc(g.phone||'연락처 없음')}</div></td><td class="ship-items">${g.items.slice(0,4).map(x=>`${v717Esc(x.item)}<br><b>수량 ${Number(x.qty)||0}</b>`).join('<hr style="border:0;border-top:1px dashed #ddd">')}${g.items.length>4?`<br>외 ${g.items.length-4}건`:''}</td><td><div class="ship-pay">상품금액 ${money(g.subtotal)}<br>배송비 ${money(g.fee)}<div class="grand">총 청구 ${money(a.charged)}</div>${g.status==='amount-mismatch'?`<div class="mismatch">실입금 ${money(a.paid)}<br>${a.diff<0?'부족':'초과'} ${money(Math.abs(a.diff))}</div>`:''}</div></td><td class="ship-address">📍 ${v717Esc(g.address||'주소 미등록')}<br><br>☎ ${v717Esc(g.phone||'-')}</td><td><span class="ship-status ${st}">${v717StatusText(g)}</span><br><small>${track?'송장 등록완료':'상품 준비중입니다.'}</small></td><td><b>${v717Esc(track||'-')}</b></td><td>${[...g.dates].join(', ')}<br><small>${v717Esc(g.items[0]?.time||'')}</small></td><td class="ship-actions"><button class="btn secondary" onclick="openReceiptDetail('${v717Esc(g.key)}')">정산서 보기</button><button class="btn secondary" onclick="openCustomerForKey('${v717Esc(g.key)}')">상세보기</button><button class="btn" onclick="togglePackingStatus('${v717Esc(g.code)}')">작업</button></td></tr>`}).join('')}</tbody></table>`;
}
function v717ReceiptHTML(r,i){const c=r.customer||{},a=v717PaymentAmounts(r),account=state.settings?.account||'계좌번호를 설정해 주세요';const orderNo=r.date.replaceAll('-','')+'-'+String(i+1).padStart(4,'0');return `<article class="receipt-sheet"><div class="rs-head"><div><div class="rs-brand">🎁 땡라이브</div><div class="rs-title">정산서</div></div><table class="rs-order"><tr><td>주문일</td><td>${v717Esc(r.date.replaceAll('-','.'))}</td></tr><tr><td>주문번호</td><td>${v717Esc(orderNo)}</td></tr></table></div><div class="rs-customer-title"><span class="rs-page-label">정산서 1</span><strong>${v717Esc(c.name||r.nick)}님 정산서</strong></div><div class="rs-info"><div><span class="label">고객명</span><span class="customer-name">${v717Esc(c.name||'확인필요')}</span></div><div><span class="label">배송지</span>${v717Esc([c.postalCode,c.address,c.detailAddress].filter(Boolean).join(' ')||'주소 미등록')}</div><div><span class="label">연락처</span><b>${v717Esc(c.phone||'연락처 미등록')}</b></div><div><span class="label">메모</span>${v717Esc(c.memo||'')}</div></div><table class="rs-table"><thead><tr><th style="width:8%">NO</th><th>상품명</th><th style="width:14%">수량</th><th style="width:18%">단가</th><th style="width:18%">금액</th></tr></thead><tbody>${r.items.map((x,n)=>`<tr><td>${n+1}</td><td>${v717Esc(x.item)}</td><td>${Number(x.qty)||0}</td><td>${money(x.unit)}</td><td>${money(x.amount)}</td></tr>`).join('')}${Array.from({length:Math.max(0,7-r.items.length)},()=>'<tr><td>&nbsp;</td><td></td><td></td><td></td><td></td></tr>').join('')}</tbody></table><div class="rs-summary"><div>상품금액 합계 <b style="float:right">${money(r.subtotal)}</b><br>배송비 <b style="float:right">${money(r.fee)}</b></div><div class="rs-total"><span>총 결제금액</span><strong>${Number(r.total||0).toLocaleString('ko-KR')}</strong><span>원</span></div></div>${r.payment.status==='amount-mismatch'?`<div class="rs-notice"><strong>금액 불일치</strong> · 청구금액 ${money(a.charged)} / 실제 입금액 ${money(a.paid)} / ${a.diff<0?'부족':'초과'}금액 ${money(Math.abs(a.diff))}</div>`:`<div class="rs-notice"><strong>일요일 12시까지</strong> 입금 바랍니다. 입금자명은 주문자명과 동일하게 부탁드립니다.</div>`}<div class="rs-bottom"><div><b>입금계좌</b><div class="rs-bank">${v717Esc(state.settings?.bank||'카카오뱅크')}<br><strong>${v717Esc(account)}</strong><br>예금주 : ${v717Esc(state.settings?.holder||'김미숙')}</div></div><div class="rs-qr"><b>카카오채널 바로가기</b><div class="rs-qr-placeholder">QR</div><small>${v717Esc(state.settings?.contact||'문의는 카카오채널로 부탁드립니다.')}</small></div></div><div class="rs-footer"><b>🎧 고객센터 ${v717Esc(state.settings?.customerService||state.settings?.phone||'010-2184-2344')}</b><span>감사합니다! 좋은 하루 되세요 :)</span></div><div class="rs-tabs"><span class="rs-tab active">정산서 1</span>${r.items.length>7?'<span class="rs-tab">정산서 2</span>':''}<span style="margin-left:auto">상품이 많으면 여러 장으로 분리 발행됩니다.</span></div></article>`}
async function openReceiptDetail(key){const r=receiptByKey(key);if(!r)return alert('정산서를 찾을 수 없습니다.');let modal=document.getElementById('receiptDetailModal');if(!modal){modal=document.createElement('div');modal.id='receiptDetailModal';modal.className='receipt-detail-modal';modal.innerHTML='<div class="receipt-detail-box" style="max-width:960px"><div class="receipt-detail-actions no-print"><button class="btn warn" id="detailEditBtn">수정</button><button class="btn secondary" id="detailPrintBtn">인쇄</button><button class="btn bad" onclick="closeReceiptDetail()">닫기</button></div><div id="receiptDetailBody"></div></div>';document.body.appendChild(modal)}modal.classList.add('show');document.getElementById('detailEditBtn').onclick=()=>openReceiptEditByKey(key);document.getElementById('detailPrintBtn').onclick=()=>window.print();document.getElementById('receiptDetailBody').innerHTML=v717ReceiptHTML(r,indexOfReceiptKey(key));}
function printOne(i){const r=getReceipts()[i];if(!r)return;const w=window.open('','_blank');w.document.write('<!doctype html><html><head><meta charset="utf-8"><title>정산서</title>'+document.getElementById('v717-final-ui').outerHTML+'</head><body>'+v717ReceiptHTML(r,i)+'<script>window.onload=()=>setTimeout(()=>window.print(),300)<\\/script></body></html>');w.document.close()}
function renderPayments(){const rs=getReceipts();let rows=paymentFilter==='all'?rs:rs.filter(r=>r.payment.status===paymentFilter);const counts={all:rs.length,paid:rs.filter(r=>r.payment.status==='paid').length,unpaid:rs.filter(r=>r.payment.status==='unpaid').length,'amount-mismatch':rs.filter(r=>r.payment.status==='amount-mismatch').length,review:rs.filter(r=>r.payment.status==='review').length};document.querySelectorAll('#paymentFilterBar [data-filter]').forEach(b=>{const k=b.dataset.filter;b.textContent=({all:'전체',paid:'입금자',unpaid:'미입금자','amount-mismatch':'금액불일치',review:'정보확인필요'}[k])+' '+counts[k]});const box=$('paymentsTable');if(!rows.length){box.innerHTML='<div class="empty">선택한 조건의 고객이 없습니다.</div>';return}box.innerHTML='<div class="scroll"><table style="min-width:1100px"><thead><tr><th>고객명</th><th>입금자</th><th>청구금액</th><th>실제 입금액</th><th>차액</th><th>대조결과</th><th>상세</th></tr></thead><tbody>'+rows.map(r=>{const p=r.payment?.payment||{},a=v717PaymentAmounts(r);return `<tr><td><b>${v717Esc(r.customer?.name||r.nick)}</b><br>${v717Esc(r.nick)}</td><td>${v717Esc(p.payer||'-')}</td><td><b>${money(a.charged)}</b></td><td><b>${p.amount!=null?money(a.paid):'-'}</b></td><td style="font-weight:900;color:${a.diff===0?'#187a32':'#c62020'}">${p.amount==null?'-':(a.diff<0?'부족 ':'초과 ')+money(Math.abs(a.diff))}</td><td>${paymentStatusLabel(r.payment.status)}</td><td><button class="btn secondary" onclick="openReceiptDetail('${v717Esc(r.key)}')">정산서 보기</button></td></tr>`}).join('')+'</tbody></table></div>'}
