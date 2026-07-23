const fs=require('fs'),path=require('path'),http=require('http');
function fail(msg){console.error('FAIL:',msg);process.exit(1)}
for(const f of ['index.html','join.html','server.js']){
 if(!fs.existsSync(path.join(__dirname,f)))fail(f+' 없음');
}
const join=fs.readFileSync(path.join(__dirname,'join.html'),'utf8');
for(const id of ['name','nickname','phone','postalCode','address','detailAddress','memo','agree','formBox','success']){
 if(!join.includes(`id="${id}"`))fail('가입폼 필드 누락: '+id);
}
for(const fn of ['findAddress','submitJoin','testServerConnection']){
 if(!join.includes(`function ${fn}`)&&!join.includes(`async function ${fn}`))fail('가입폼 함수 누락: '+fn);
}
console.log('정적 점검 통과');
