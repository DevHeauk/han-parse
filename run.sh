#!/bin/bash
# 웹 서버 실행 스크립트

echo "한글 파일 파서 웹 서버를 시작합니다..."
echo "브라우저에서 http://localhost:8080 으로 접속하세요"
echo ""

# 기존 서버 프로세스 종료
echo "기존 서버 프로세스 확인 중..."
OLD_PID=$(lsof -ti:8080 2>/dev/null)
if [ ! -z "$OLD_PID" ]; then
    echo "기존 서버 프로세스(PID: $OLD_PID)를 종료합니다..."
    kill -9 $OLD_PID 2>/dev/null
    sleep 1
    echo "기존 서버가 종료되었습니다."
else
    echo "실행 중인 서버가 없습니다."
fi

# Python app.py 프로세스도 확인
PYTHON_PIDS=$(ps aux | grep "python.*app.py" | grep -v grep | awk '{print $2}')
if [ ! -z "$PYTHON_PIDS" ]; then
    echo "Python app.py 프로세스를 종료합니다..."
    echo "$PYTHON_PIDS" | xargs kill -9 2>/dev/null
    sleep 1
fi

echo ""
echo "새 서버를 시작합니다..."
echo ""

# 새 서버 시작
python3 app.py

