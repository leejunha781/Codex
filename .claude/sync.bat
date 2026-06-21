@echo off
echo 🔄 윈도우 프로젝트 환경 최신화 파일 생성 중...
set OUTPUT=claude-context.md

:: 기존 파일이 있으면 삭제
if exist %OUTPUT% del %OUTPUT%

:: 프로젝트 내 소스 코드들을 하나의 파일로 병합 (node_modules, .git 등 제외)
for /r %%i in (*.js *.ts *.tsx *.py *.json *.html *.css *.md) do (
    echo %%i | findstr /i /v "node_modules .git dist build" >nul
    if errorlevel 1 (
        rem 제외 폴더 패스
    ) else (
        echo === FILE: %%~nxi === >> %OUTPUT%
        type "%%i" >> %OUTPUT%
        echo. >> %OUTPUT%
        echo. >> %OUTPUT%
    )
)

echo ✅ 동기화 완료! '%OUTPUT%' 파일이 생성되었습니다.
pause