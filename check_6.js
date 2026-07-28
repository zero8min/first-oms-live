
(async()=>{
 try{
  const r=await fetch('/api/admin/current-tenant',{credentials:'same-origin',cache:'no-store'});
  if(!r.ok)return;
  const d=await r.json(),b=document.getElementById('tenantGuard');
  if(d.actingAs&&b){
   b.style.display='flex';b.style.alignItems='center';b.style.justifyContent='center';b.style.gap='12px';b.style.flexWrap='wrap';
   const text=document.createElement('span');text.textContent=`현재 거래처: ${d.tenant.company} (${d.tenant.code}) · 이 화면의 저장 내용은 해당 거래처에만 반영됩니다.`;
   const back=document.createElement('button');back.type='button';back.textContent='거래처 관리로 돌아가기';back.style.cssText='border:0;border-radius:8px;padding:7px 11px;font-weight:900;cursor:pointer;background:#fff;color:#7b2d2d';back.onclick=()=>location.href='/superadmin.html';
   b.replaceChildren(text,back);
  }
 }catch(e){console.warn('거래처 표시를 불러오지 못했습니다.',e)}
})();
