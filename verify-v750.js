const fs=require('fs');
const checks=[
 ['index.html','포장단위 자동환산'],['index.html','printClassifiedReceiptsV750'],['index.html','packing/access-link'],['index.html','포장팀 접속키 재발급'],
 ['packing.html','FIRST_OMS_PACKING_TOKEN'],['packing.html','실제 포장 총'],['packing.html','송장을 스캔해주세요'],
 ['server.js','validPackingTokenV750'],['server.js','/api/packing/access-reset'],['server.js','X-Frame-Options']
];
let bad=0;for(const [f,t] of checks){const ok=fs.readFileSync(f,'utf8').includes(t);console.log(ok?'PASS':'FAIL',f,t);if(!ok)bad++}process.exit(bad?1:0);
