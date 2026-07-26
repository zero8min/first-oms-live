const fs=require('fs'),path=require('path'),crypto=require('crypto');
const ROOT=__dirname,DATA_ROOT=process.env.DATA_DIR||path.join(ROOT,'data'),file=path.join(DATA_ROOT,'accounts.json'),backup=path.join(DATA_ROOT,'accounts-backup.json');
function hash(p,s=crypto.randomBytes(16).toString('hex')){return s+':'+crypto.scryptSync(String(p),s,64).toString('hex')}
if(!fs.existsSync(file))throw new Error('accounts.json이 없습니다. 먼저 서버를 한 번 실행해 주세요.');
const list=JSON.parse(fs.readFileSync(file,'utf8'));const a=list.find(x=>x.role==='superadmin');if(!a)throw new Error('최고관리자 계정을 찾을 수 없습니다.');
fs.copyFileSync(file,backup);a.username='firstadmin';a.passwordHash=hash('12345678');a.mustChangePassword=true;a.passwordResetAt=new Date().toISOString();
const tmp=file+'.tmp';fs.writeFileSync(tmp,JSON.stringify(list,null,2));fs.renameSync(tmp,file);console.log('완료: firstadmin / 12345678 (로그인 후 변경 필요)');
