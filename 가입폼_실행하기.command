#!/bin/bash
cd "$(dirname "$0")"

if ! command -v node >/dev/null 2>&1; then
  osascript -e 'display alert "Node.js가 설치되어 있지 않습니다" message "Node.js LTS를 설치한 뒤 다시 실행해 주세요."'
  open "https://nodejs.org/"
  exit 1
fi

PORT=3010 node server.js > server.log 2>&1 &
PID=$!
sleep 2

if ! curl -s "http://localhost:3010/api/health" | grep -q '"ok":true'; then
  osascript -e 'display alert "가입폼 서버 실행 실패" message "폴더 안 server.log 파일을 확인해 주세요."'
  exit 1
fi

open "http://localhost:3010/join.html"
echo ""
echo "가입폼: http://localhost:3010/join.html"
echo "관리자: http://localhost:3010/index.html"
echo "이 터미널 창을 닫으면 서버가 종료됩니다."
wait $PID
