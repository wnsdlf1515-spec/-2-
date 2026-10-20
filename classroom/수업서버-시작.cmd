@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0start-local.cjs" goto incomplete
if not exist "%~dp0server.cjs" goto incomplete
if not exist "%~dp0public\game.html" goto incomplete
if not exist "%~dp0..\city-stage6\src\engine.js" goto incomplete
where node >nul 2>nul
if errorlevel 1 (
  echo Node.js 24 이상이 필요합니다. https://nodejs.org/ 에서 설치하세요.
  pause
  exit /b 1
)
node "%~dp0start-local.cjs"
pause
exit /b

:incomplete
echo.
echo [압축을 모두 풀어 주세요]
echo 실행에 필요한 파일을 찾을 수 없습니다.
echo ZIP 안에서 CMD를 바로 실행하면 이 문제가 발생할 수 있습니다.
echo.
echo 1. 이 창을 닫으세요.
echo 2. 다운로드한 ZIP을 마우스 오른쪽 버튼으로 누르세요.
echo 3. '모두 압축 풀기'를 선택하세요.
echo 4. 풀린 폴더의 classroom 폴더에서 이 CMD를 실행하세요.
echo.
echo classroom과 city-stage6 폴더는 같은 상위 폴더에 두세요.
echo CMD 파일 하나만 바탕화면으로 옮기면 실행되지 않습니다.
echo.
pause
exit /b 1
