const fs=require('fs'),path=require('path'),vm=require('vm');
function fail(m){console.error('FAIL',m);process.exitCode=1}
for(const f of ['index.html','login.html','signup.html','superadmin.html','join.html','server.js','package.json'])if(!fs.existsSync(path.join(__dirname,f)))fail(f+' 없음');
const html=fs.readFileSync(path.join(__dirname,'index.html'),'utf8');
const defs=new Set([...html.matchAll(/(?:async\s+)?function\s+([A-Za-z_$][\w$]*)\s*\(/g)].map(x=>x[1]));
const calls=[...html.matchAll(/onclick="\s*([A-Za-z_$][\w$]*)\s*\(/g)].map(x=>x[1]);
const allowed=new Set(['location','window','document']);const missing=[...new Set(calls.filter(x=>!defs.has(x)&&!allowed.has(x)))];
if(missing.length)fail('onclick 함수 누락: '+missing.join(','));
else console.log('PASS onclick 버튼 함수',new Set(calls).size+'개');
console.log(process.exitCode?'점검 실패':'정적 점검 통과');
